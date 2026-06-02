# backend/api/v1/endpoints/system_tools.py
"""
System tools endpoints mirroring BackupExportarPage and SystemToolsPage.
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import zipfile
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel, Field
from sqlalchemy import delete, func, select, text
from sqlalchemy.engine import make_url
from sqlalchemy.ext.asyncio import AsyncSession

from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.config.settings import settings
from backend.models.system import SistemaLogDB
from backend.models.user import User
from backend.models.venda_importada import VendaImportadaDB
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()

PROJECT_ROOT = Path(__file__).resolve().parents[4]
BACKUP_DIR = Path(os.getenv("BACKUP_DIR", PROJECT_ROOT / "backup")).resolve()
EXPORT_DIR = Path(os.getenv("ROTA_EXPORT_DIR", PROJECT_ROOT / "exports")).resolve()
INSTALLER_DIR = Path(os.getenv("INSTALLER_DIR", PROJECT_ROOT / "dist_installer")).resolve()
SYSTEM_TABLES = [
    "usuarios",
    "motoristas",
    "vendedores",
    "clientes",
    "veiculos",
    "ajudantes",
    "programacoes",
    "programacao_itens",
    "recebimentos",
    "despesas",
    "vendas_importadas",
]


class BackupInfo(BaseModel):
    arquivo: str
    caminho: str = ""
    tamanho_kb: float = 0
    data_criacao: str = ""


class BackupCreateResponse(BaseModel):
    arquivo: str
    caminho: str
    tamanho_kb: float
    timestamp: str
    mensagem: str


class RestorePayload(BaseModel):
    confirmar: bool = False


class RestoreResponse(BaseModel):
    arquivo_restaurado: str
    backup_anterior: str = ""
    mensagem: str


class SistemaLogResponse(BaseModel):
    id: int
    tipo_acao: str = ""
    descricao: str = ""
    usuario: str = ""
    status: str = ""
    resultado: str = ""
    executado_em: str = ""


class ClearLogsResponse(BaseModel):
    linhas_deletadas: int = 0


class IntegrityResponse(BaseModel):
    ok: bool
    integridade: str
    verificado_em: str


class SystemInfoResponse(BaseModel):
    registros_por_tabela: dict[str, int] = Field(default_factory=dict)
    tamanho_banco_kb: float = 0
    ultima_acao_em: str | None = None
    data_hora_atual: str
    database_path: str = ""
    backup_dir: str = ""


class SystemOverviewResponse(BaseModel):
    backups: list[BackupInfo]
    logs: list[SistemaLogResponse]
    info: SystemInfoResponse


class DiariaConfigItem(BaseModel):
    local_rota: str
    motorista: float = Field(default=0, ge=0)
    ajudante: float = Field(default=0, ge=0)


class DiariaConfigResponse(BaseModel):
    itens: list[DiariaConfigItem]


class DiariaConfigPayload(BaseModel):
    serra_motorista: float = Field(default=0, ge=0)
    serra_ajudante: float = Field(default=0, ge=0)
    sertao_motorista: float = Field(default=0, ge=0)
    sertao_ajudante: float = Field(default=0, ge=0)


def sqlite_db_path() -> Path:
    url = make_url(settings.DATABASE_URL)
    if not url.drivername.startswith("sqlite"):
        raise HTTPException(status_code=409, detail="Backup local disponivel somente para banco SQLite.")
    if not url.database or url.database == ":memory:":
        raise HTTPException(status_code=409, detail="Backup indisponivel para banco SQLite em memoria.")
    db_path = Path(url.database)
    if not db_path.is_absolute():
        db_path = (PROJECT_ROOT / db_path).resolve()
    return db_path


def ensure_backup_dir() -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    return BACKUP_DIR


def ensure_export_dir() -> Path:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    return EXPORT_DIR


def assert_sqlite_file(path: Path) -> None:
    try:
        with path.open("rb") as handle:
            header = handle.read(16)
    except OSError as exc:
        raise HTTPException(status_code=404, detail="Arquivo de banco nao encontrado.") from exc
    if header != b"SQLite format 3\x00":
        raise HTTPException(status_code=422, detail="O arquivo informado nao e um banco SQLite valido.")


def backup_file_path(filename: str) -> Path:
    name = Path(filename).name
    path = (ensure_backup_dir() / name).resolve()
    if path.parent != BACKUP_DIR:
        raise HTTPException(status_code=400, detail="Nome de backup invalido.")
    if not path.exists():
        raise HTTPException(status_code=404, detail="Backup nao encontrado.")
    if path.suffix.lower() != ".db":
        raise HTTPException(status_code=422, detail="Arquivo de backup invalido.")
    return path


def latest_installer_path() -> Path:
    if not INSTALLER_DIR.exists():
        raise HTTPException(status_code=404, detail="Diretorio de instaladores nao encontrado.")
    candidates = [
        path
        for path in INSTALLER_DIR.iterdir()
        if path.is_file() and path.suffix.lower() in {".exe", ".msi"}
    ]
    if not candidates:
        raise HTTPException(status_code=404, detail="Nenhum instalador encontrado.")
    return max(candidates, key=lambda item: item.stat().st_mtime)


def sqlite_backup_copy(src: Path, dst: Path) -> None:
    src_conn = sqlite3.connect(str(src))
    dst_conn = sqlite3.connect(str(dst))
    try:
        with dst_conn:
            src_conn.backup(dst_conn)
    finally:
        dst_conn.close()
        src_conn.close()


def file_copy(src: Path, dst: Path) -> None:
    with src.open("rb") as fsrc, dst.open("wb") as fdst:
        shutil.copyfileobj(fsrc, fdst)


def cleanup_sqlite_sidecars(db_path: Path) -> None:
    for suffix in ("-wal", "-shm"):
        sidecar = Path(f"{db_path}{suffix}")
        if sidecar.exists():
            try:
                sidecar.unlink()
            except OSError:
                pass


def backup_to_info(path: Path) -> BackupInfo:
    stat = path.stat()
    return BackupInfo(
        arquivo=path.name,
        caminho=str(path),
        tamanho_kb=round(stat.st_size / 1024, 2),
        data_criacao=datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
    )


def create_migration_package() -> Path:
    db_path = sqlite_db_path()
    if not db_path.exists():
        raise HTTPException(status_code=404, detail="Banco de dados principal nao encontrado.")
    assert_sqlite_file(db_path)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    export_dir = ensure_export_dir()
    archive_path = export_dir / f"rotahub_migration_{timestamp}.zip"
    backup_db = export_dir / f".rotadb_export_{timestamp}.db"
    photos_dir = Path(
        os.getenv("ROTA_MOBILE_PHOTOS_DIR", PROJECT_ROOT / ".rotahub_runtime" / "fotos_rotas")
    ).expanduser()
    photos_count = 0

    try:
        sqlite_backup_copy(db_path, backup_db)
        manifest = {
            "format": "rotahub-migration-v1",
            "created_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "database": {
                "type": "sqlite",
                "archive_path": "database/rotadb.db",
                "source_name": db_path.name,
                "size_bytes": backup_db.stat().st_size,
            },
            "photos": {
                "archive_path": "fotos_rotas/",
                "included": photos_dir.exists(),
                "files": 0,
            },
            "restore": {
                "rota_db": "/var/rotahub/data/rotadb.db",
                "photos_dir": "/var/rotahub/data/fotos_rotas",
            },
        }
        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(backup_db, "database/rotadb.db")
            if photos_dir.exists():
                for path in sorted(photos_dir.rglob("*")):
                    if not path.is_file():
                        continue
                    archive.write(path, Path("fotos_rotas") / path.relative_to(photos_dir))
                    photos_count += 1
            manifest["photos"]["files"] = photos_count
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
    finally:
        backup_db.unlink(missing_ok=True)

    return archive_path


async def registrar_log(
    db: AsyncSession,
    *,
    tipo_acao: str,
    descricao: str,
    usuario: str,
    status: str = "OK",
    resultado: str = "",
) -> None:
    db.add(
        SistemaLogDB(
            tipo_acao=tipo_acao,
            descricao=descricao,
            usuario=usuario,
            status=status,
            resultado_texto=resultado,
            executado_em=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
    )


def username(current_user: User) -> str:
    return str(getattr(current_user, "nome", None) or getattr(current_user, "username", None) or "ADMIN")


def normalize_local_rota(value: Any) -> str:
    text_value = str(value or "").strip().upper()
    if text_value.startswith("SERRA"):
        return "SERRA"
    if text_value.startswith("SERT"):
        return "SERTAO"
    return text_value


async def ensure_diarias_config_table(db: AsyncSession) -> None:
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS diaria_config (
                local_rota TEXT PRIMARY KEY,
                motorista_valor REAL DEFAULT 0,
                ajudante_valor REAL DEFAULT 0,
                atualizado_em TEXT,
                atualizado_por TEXT
            )
            """
        )
    )
    for local in ("SERRA", "SERTAO"):
        await db.execute(
            text(
                """
                INSERT INTO diaria_config (local_rota, motorista_valor, ajudante_valor, atualizado_em, atualizado_por)
                VALUES (:local_rota, 0, 0, :atualizado_em, 'SISTEMA')
                ON CONFLICT(local_rota) DO NOTHING
                """
            ),
            {"local_rota": local, "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        )


async def diarias_config_data(db: AsyncSession) -> DiariaConfigResponse:
    await ensure_diarias_config_table(db)
    result = await db.execute(
        text(
            """
            SELECT local_rota, motorista_valor, ajudante_valor
              FROM diaria_config
             WHERE local_rota IN ('SERRA', 'SERTAO')
             ORDER BY CASE local_rota WHEN 'SERRA' THEN 1 WHEN 'SERTAO' THEN 2 ELSE 9 END
            """
        )
    )
    return DiariaConfigResponse(
        itens=[
            DiariaConfigItem(
                local_rota=normalize_local_rota(row["local_rota"]),
                motorista=round(float(row["motorista_valor"] or 0), 2),
                ajudante=round(float(row["ajudante_valor"] or 0), 2),
            )
            for row in result.mappings().all()
        ]
    )


def system_log_response(item: SistemaLogDB) -> SistemaLogResponse:
    return SistemaLogResponse(
        id=item.id,
        tipo_acao=item.tipo_acao or "",
        descricao=item.descricao or "",
        usuario=item.usuario or "",
        status=item.status or "",
        resultado=item.resultado_texto or "",
        executado_em=item.executado_em or "",
    )


async def list_backups_data() -> list[BackupInfo]:
    directory = ensure_backup_dir()
    return [
        backup_to_info(path)
        for path in sorted(directory.glob("banco_de_dados_*.db"), key=lambda item: item.stat().st_mtime, reverse=True)
    ]


async def list_logs_data(db: AsyncSession, limit: int = 100) -> list[SistemaLogResponse]:
    result = await db.execute(
        select(SistemaLogDB).order_by(SistemaLogDB.id.desc()).limit(max(min(limit, 500), 1))
    )
    return [system_log_response(item) for item in result.scalars().all()]


async def system_info_data(db: AsyncSession) -> SystemInfoResponse:
    stats: dict[str, int] = {}
    for table_name in SYSTEM_TABLES:
        try:
            result = await db.execute(text(f"SELECT COUNT(*) FROM {table_name}"))
            stats[table_name] = int(result.scalar() or 0)
        except Exception:
            continue

    db_path = sqlite_db_path()
    last_log = await db.execute(select(func.max(SistemaLogDB.executado_em)))
    return SystemInfoResponse(
        registros_por_tabela=stats,
        tamanho_banco_kb=round(db_path.stat().st_size / 1024, 2) if db_path.exists() else 0,
        ultima_acao_em=last_log.scalar(),
        data_hora_atual=datetime.now().isoformat(timespec="seconds"),
        database_path=str(db_path),
        backup_dir=str(ensure_backup_dir()),
    )


@router.get("/overview", response_model=SystemOverviewResponse)
async def system_tools_overview(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    return SystemOverviewResponse(
        backups=await list_backups_data(),
        logs=await list_logs_data(db, 100),
        info=await system_info_data(db),
    )


@router.get("/diarias", response_model=DiariaConfigResponse)
async def obter_config_diarias(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    return await diarias_config_data(db)


@router.put("/diarias", response_model=DiariaConfigResponse)
async def salvar_config_diarias(
    payload: DiariaConfigPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_diarias_config_table(db)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    values = {
        "SERRA": (payload.serra_motorista, payload.serra_ajudante),
        "SERTAO": (payload.sertao_motorista, payload.sertao_ajudante),
    }
    for local, (motorista_valor, ajudante_valor) in values.items():
        await db.execute(
            text(
                """
                INSERT INTO diaria_config (local_rota, motorista_valor, ajudante_valor, atualizado_em, atualizado_por)
                VALUES (:local_rota, :motorista_valor, :ajudante_valor, :atualizado_em, :atualizado_por)
                ON CONFLICT(local_rota) DO UPDATE SET
                    motorista_valor=excluded.motorista_valor,
                    ajudante_valor=excluded.ajudante_valor,
                    atualizado_em=excluded.atualizado_em,
                    atualizado_por=excluded.atualizado_por
                """
            ),
            {
                "local_rota": local,
                "motorista_valor": round(float(motorista_valor or 0), 2),
                "ajudante_valor": round(float(ajudante_valor or 0), 2),
                "atualizado_em": now,
                "atualizado_por": username(current_user),
            },
        )
    await registrar_log(
        db,
        tipo_acao="CONFIG_DIARIAS",
        descricao="Valores de diarias atualizados",
        usuario=username(current_user),
        status="OK",
        resultado=f"SERRA M:{payload.serra_motorista} A:{payload.serra_ajudante} | SERTAO M:{payload.sertao_motorista} A:{payload.sertao_ajudante}",
    )
    record_audit_log(
        db,
        action="config_diarias_atualizada",
        actor_user=current_user,
        entity_type="diaria_config",
        entity_id="SERRA_SERTAO",
        ip_address=client_ip_from_request(request),
        metadata=payload.model_dump(),
    )
    await db.commit()
    return await diarias_config_data(db)


@router.post("/backups", response_model=BackupCreateResponse)
async def criar_backup(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    db_path = sqlite_db_path()
    if not db_path.exists():
        raise HTTPException(status_code=404, detail="Banco de dados principal nao encontrado.")
    assert_sqlite_file(db_path)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = ensure_backup_dir() / f"banco_de_dados_{timestamp}.db"
    try:
        try:
            sqlite_backup_copy(db_path, target)
        except Exception:
            file_copy(db_path, target)
    except Exception as exc:
        await registrar_log(
            db,
            tipo_acao="BACKUP",
            descricao="Erro ao fazer backup",
            usuario=username(current_user),
            status="ERRO",
            resultado=str(exc),
        )
        await db.commit()
        raise HTTPException(status_code=500, detail=f"Erro ao criar backup: {exc}") from exc

    info = backup_to_info(target)
    await registrar_log(
        db,
        tipo_acao="BACKUP",
        descricao=f"Backup criado: {info.arquivo}",
        usuario=username(current_user),
        status="OK",
        resultado=info.caminho,
    )
    record_audit_log(
        db,
        action="system_backup_criado",
        actor_user=current_user,
        entity_type="backup",
        entity_id=info.arquivo,
        ip_address=client_ip_from_request(request),
        metadata={"tamanho_kb": info.tamanho_kb},
    )
    await db.commit()
    return BackupCreateResponse(
        arquivo=info.arquivo,
        caminho=info.caminho,
        tamanho_kb=info.tamanho_kb,
        timestamp=timestamp,
        mensagem="Backup criado com sucesso.",
    )


@router.get("/backups", response_model=list[BackupInfo])
async def listar_backups(
    current_user: User = Depends(require_admin_user),
):
    return await list_backups_data()


@router.get("/backups/{filename}/download")
async def baixar_backup(
    filename: str,
    current_user: User = Depends(require_admin_user),
):
    path = backup_file_path(filename)
    assert_sqlite_file(path)
    return FileResponse(
        path,
        media_type="application/x-sqlite3",
        filename=path.name,
    )


@router.get("/migration/export/download")
async def baixar_pacote_migracao(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    try:
        path = create_migration_package()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Erro ao criar pacote de migracao: {exc}") from exc
    await registrar_log(
        db,
        tipo_acao="EXPORTAR_MIGRACAO",
        descricao=f"Pacote de migracao criado: {path.name}",
        usuario=username(current_user),
        status="OK",
        resultado=str(path),
    )
    record_audit_log(
        db,
        action="system_migration_exportado",
        actor_user=current_user,
        entity_type="migration_package",
        entity_id=path.name,
        ip_address=client_ip_from_request(request),
        metadata={"tamanho_kb": round(path.stat().st_size / 1024, 2)},
    )
    await db.commit()
    return FileResponse(
        path,
        media_type="application/zip",
        filename=path.name,
    )


@router.get("/installer/download")
async def baixar_instalador(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    path = latest_installer_path()
    await registrar_log(
        db,
        tipo_acao="DOWNLOAD_INSTALADOR",
        descricao=f"Download do instalador: {path.name}",
        usuario=username(current_user),
        status="OK",
        resultado=str(path),
    )
    record_audit_log(
        db,
        action="system_instalador_baixado",
        actor_user=current_user,
        entity_type="installer",
        entity_id=path.name,
        ip_address=client_ip_from_request(request),
        metadata={"tamanho_kb": round(path.stat().st_size / 1024, 2)},
    )
    await db.commit()
    return FileResponse(
        path,
        media_type="application/octet-stream",
        filename=path.name,
    )


def restore_backup_file(src: Path, db_path: Path) -> Path:
    assert_sqlite_file(src)
    auto_backup = db_path.with_name(f"auto_backup_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    if db_path.exists():
        try:
            sqlite_backup_copy(db_path, auto_backup)
        except Exception:
            file_copy(db_path, auto_backup)
    cleanup_sqlite_sidecars(db_path)
    try:
        sqlite_backup_copy(src, db_path)
    except Exception:
        file_copy(src, db_path)
    cleanup_sqlite_sidecars(db_path)
    return auto_backup


@router.post("/backups/{filename}/restore", response_model=RestoreResponse)
async def restaurar_backup_salvo(
    filename: str,
    payload: RestorePayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    if not payload.confirmar:
        raise HTTPException(status_code=422, detail="Confirmacao obrigatoria para restaurar backup.")
    src = backup_file_path(filename)
    db_path = sqlite_db_path()
    await registrar_log(
        db,
        tipo_acao="RESTAURACAO",
        descricao=f"Banco restaurado de: {src.name}",
        usuario=username(current_user),
        status="OK",
        resultado="Restauracao solicitada pela web.",
    )
    record_audit_log(
        db,
        action="system_backup_restaurado",
        actor_user=current_user,
        entity_type="backup",
        entity_id=src.name,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"origem": "backup_salvo"},
    )
    await db.commit()
    auto_backup = restore_backup_file(src, db_path)
    return RestoreResponse(
        arquivo_restaurado=src.name,
        backup_anterior=str(auto_backup) if auto_backup.exists() else "",
        mensagem="Backup restaurado. Reinicie o servidor para reabrir conexoes do banco.",
    )


@router.post("/backups/upload-restore", response_model=RestoreResponse)
async def restaurar_backup_upload(
    request: Request,
    confirmar: bool = False,
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    if not confirmar:
        raise HTTPException(status_code=422, detail="Confirmacao obrigatoria para restaurar backup.")
    if not file.filename or not file.filename.lower().endswith(".db"):
        raise HTTPException(status_code=422, detail="Selecione um arquivo .db.")
    upload_dir = ensure_backup_dir()
    temp_path = upload_dir / f"upload_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{Path(file.filename).name}"
    try:
        content = await file.read()
        temp_path.write_bytes(content)
        assert_sqlite_file(temp_path)
        await registrar_log(
            db,
            tipo_acao="RESTAURACAO",
            descricao=f"Banco restaurado de upload: {file.filename}",
            usuario=username(current_user),
            status="OK",
            resultado=str(temp_path),
        )
        record_audit_log(
            db,
            action="system_backup_upload_restaurado",
            actor_user=current_user,
            entity_type="backup",
            entity_id=Path(file.filename).name,
            severity="warning",
            ip_address=client_ip_from_request(request),
            metadata={"origem": "upload"},
        )
        await db.commit()
        auto_backup = restore_backup_file(temp_path, sqlite_db_path())
        return RestoreResponse(
            arquivo_restaurado=Path(file.filename).name,
            backup_anterior=str(auto_backup) if auto_backup.exists() else "",
            mensagem="Backup restaurado. Reinicie o servidor para reabrir conexoes do banco.",
        )
    finally:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass


@router.get("/vendas-importadas/export")
async def exportar_vendas_importadas(
    current_user: User = Depends(require_admin_user),
    db: AsyncSession = Depends(get_db),
):
    try:
        from openpyxl import Workbook
    except Exception as exc:  # pragma: no cover - optional package
        raise HTTPException(status_code=503, detail="Biblioteca openpyxl indisponivel.") from exc

    result = await db.execute(select(VendaImportadaDB).order_by(VendaImportadaDB.id.desc()))
    rows = result.scalars().all()
    if not rows:
        raise HTTPException(status_code=404, detail="Nao ha vendas importadas para exportar.")

    fields = [
        "id",
        "pedido",
        "data_venda",
        "cliente",
        "nome_cliente",
        "vendedor",
        "produto",
        "vr_total",
        "qnt",
        "cidade",
        "valor_unitario",
        "observacao",
        "selecionada",
        "usada",
        "usada_em",
        "codigo_programacao",
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "VENDAS_IMPORTADAS"
    ws.append(fields)
    for item in rows:
        ws.append([getattr(item, field, None) for field in fields])
    for column_cells in ws.columns:
        letter = column_cells[0].column_letter
        max_len = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 46)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    filename = f"VENDAS_IMPORTADAS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/logs", response_model=list[SistemaLogResponse])
async def listar_logs_sistema(
    limit: int = 100,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    return await list_logs_data(db, limit)


@router.delete("/logs", response_model=ClearLogsResponse)
async def limpar_logs_sistema(
    request: Request,
    dias: int = 30,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    dias = max(min(int(dias), 3650), 1)
    cutoff = (datetime.now() - timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")
    result = await db.execute(delete(SistemaLogDB).where(SistemaLogDB.executado_em < cutoff))
    deleted = int(result.rowcount or 0)
    await registrar_log(
        db,
        tipo_acao="LIMPEZA_LOGS",
        descricao=f"Removidos {deleted} logs com mais de {dias} dias",
        usuario=username(current_user),
        status="OK",
        resultado=str(deleted),
    )
    record_audit_log(
        db,
        action="system_logs_limpos",
        actor_user=current_user,
        entity_type="sistema_logs",
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"dias": dias, "linhas_deletadas": deleted},
    )
    await db.commit()
    return ClearLogsResponse(linhas_deletadas=deleted)


@router.get("/integridade", response_model=IntegrityResponse)
async def verificar_integridade_banco(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    try:
        result = await db.execute(text("PRAGMA integrity_check"))
        integridade = str(result.scalar() or "")
        ok = integridade.lower() == "ok"
    except Exception as exc:
        integridade = str(exc)
        ok = False

    await registrar_log(
        db,
        tipo_acao="VERIFICACAO",
        descricao="Integridade do banco verificada" if ok else "Problemas encontrados na integridade",
        usuario=username(current_user),
        status="OK" if ok else "AVISO",
        resultado=integridade,
    )
    record_audit_log(
        db,
        action="system_integridade_verificada",
        actor_user=current_user,
        entity_type="database",
        severity="info" if ok else "warning",
        ip_address=client_ip_from_request(request),
        metadata={"integridade": integridade},
    )
    await db.commit()
    return IntegrityResponse(ok=ok, integridade="OK" if ok else integridade, verificado_em=datetime.now().isoformat(timespec="seconds"))


@router.get("/info", response_model=SystemInfoResponse)
async def informacoes_sistema(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    return await system_info_data(db)
