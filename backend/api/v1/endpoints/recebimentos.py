# backend/api/v1/endpoints/recebimentos.py
"""
Recebimentos endpoints mirroring the desktop RecebimentosPage core flow.
"""
from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Request, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from sqlalchemy import delete, func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import normalize_date, normalize_time, safe_float, safe_int
from backend.api.v1.endpoints.programacao import get_programacao_by_codigo, upper_text
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import AjudanteDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemDB
from backend.models.recebimento import RecebimentoDB
from backend.models.user import User
from backend.services.audit import client_ip_from_request, record_audit_log
from backend.services.roteiro_operacional import registrar_roteiro_operacional

router = APIRouter()

FORMAS_VALIDAS = {"DINHEIRO", "PIX", "CARTAO", "BOLETO", "OUTRO"}
CANCELLED_STATUSES = {"CANCELADA", "CANCELADO"}


class RecebimentoProgramacaoOption(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    veiculo: str = ""
    status: str = ""
    prestacao_status: str = "PENDENTE"
    fechada: bool = False


class RecebimentoCabecalho(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    motorista_nome: str = ""
    veiculo: str = ""
    equipe: str = ""
    equipe_nomes: str = ""
    rota: str = ""
    status: str = ""
    prestacao_status: str = "PENDENTE"
    num_nf: str = ""
    data_saida: str = ""
    hora_saida: str = ""
    data_chegada: str = ""
    hora_chegada: str = ""
    diaria_motorista_valor: float = 0
    fechada: bool = False


class DiariasResumo(BaseModel):
    qtd_diarias: float = 0
    qtd_ajudantes: int = 0
    diaria_motorista: float = 0
    diaria_ajudante: float = 0
    total_motorista: float = 0
    total_ajudantes: float = 0
    total_geral: float = 0
    observacao_motorista: str = ""
    observacao_ajudantes: str = ""


class RecebimentoCliente(BaseModel):
    cod_cliente: str
    nome_cliente: str
    valor: float = 0
    forma_pagamento: str = ""
    observacao: str = ""
    data_registro: str = ""


class RecebimentoBundleResponse(BaseModel):
    cabecalho: RecebimentoCabecalho
    diarias: DiariasResumo
    clientes: list[RecebimentoCliente]
    total_recebido: float = 0


class RecebimentoCabecalhoPayload(BaseModel):
    data_saida: str | None = Field(default=None, max_length=20)
    hora_saida: str | None = Field(default=None, max_length=20)
    data_chegada: str | None = Field(default=None, max_length=20)
    hora_chegada: str | None = Field(default=None, max_length=20)
    diaria_motorista_valor: float = Field(default=0, ge=0)


class RecebimentoPayload(BaseModel):
    cod_cliente: str = Field(min_length=1, max_length=80)
    nome_cliente: str = Field(min_length=1, max_length=180)
    valor: float = Field(gt=0)
    forma_pagamento: str = Field(min_length=1, max_length=40)
    observacao: str | None = Field(default=None, max_length=300)
    num_nf: str | None = Field(default=None, max_length=80)


class ClienteManualPayload(BaseModel):
    cod_cliente: str = Field(min_length=1, max_length=80)
    nome_cliente: str = Field(min_length=1, max_length=180)


def status_ref(programacao: ProgramacaoDB) -> str:
    status_value = upper_text(programacao.status_operacional or programacao.status)
    if not status_value and safe_int(programacao.finalizada_no_app, 0) == 1:
        return "FINALIZADA"
    return status_value or "ATIVA"


def is_closed(programacao: ProgramacaoDB) -> bool:
    return upper_text(programacao.prestacao_status or "PENDENTE") == "FECHADA"


def assert_can_mutate(programacao: ProgramacaoDB) -> None:
    if is_closed(programacao):
        raise HTTPException(status_code=409, detail="Esta prestacao ja esta FECHADA.")
    if status_ref(programacao) in CANCELLED_STATUSES:
        raise HTTPException(status_code=409, detail="Programacao cancelada nao aceita recebimentos.")


def assert_can_open_recebimentos(programacao: ProgramacaoDB) -> None:
    if status_ref(programacao) in CANCELLED_STATUSES:
        raise HTTPException(status_code=409, detail="Programacao cancelada nao aceita recebimentos.")


async def ajudante_map(db: AsyncSession) -> dict[str, str]:
    result = await db.execute(select(AjudanteDB).order_by(AjudanteDB.id.asc()))
    return {
        str(item.id): upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
        for item in result.scalars().all()
    }


def equipe_nomes(equipe_raw: str | None, nomes_por_id: dict[str, str]) -> str:
    raw = str(equipe_raw or "").strip()
    if not raw:
        return ""
    out = []
    seen = set()
    for part in re.split(r"[|,;/]+", raw):
        token = part.strip()
        if not token:
            continue
        nome = nomes_por_id.get(token) if token.isdigit() else ""
        nome = nome or upper_text(token)
        if nome in seen:
            continue
        seen.add(nome)
        out.append(nome)
    return " / ".join(out)


def parse_dt_diaria(data_s: str | None, hora_s: str | None) -> datetime | None:
    data_s = str(data_s or "").strip()
    hora_s = str(hora_s or "").strip()
    if not data_s:
        return None
    normalized_date = normalize_date(data_s)
    if normalized_date is None:
        return None
    normalized_time = normalize_time(hora_s) if hora_s else "00:00:00"
    if normalized_time is None:
        return None
    try:
        y, m, d = normalized_date.split("-")
        hh, mm, *_ = normalized_time.split(":")
        return datetime(int(y), int(m), int(d), int(hh), int(mm))
    except Exception:
        return None


def calc_qtd_diarias(data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str) -> float:
    dt_saida = parse_dt_diaria(data_saida, hora_saida)
    dt_chegada = parse_dt_diaria(data_chegada, hora_chegada)
    if not dt_saida:
        return 0.0
    if not dt_chegada or dt_chegada <= dt_saida:
        return 1.0
    horas = (dt_chegada - dt_saida).total_seconds() / 3600.0
    if horas <= 24.0:
        return 1.0
    rem = horas - 24.0
    full = int(rem // 24.0)
    half = 0.5 if (rem - (full * 24.0)) > 0 else 0.0
    return 1.0 + full + half


def count_ajudantes(equipe_raw: str | None) -> int:
    raw = str(equipe_raw or "").strip()
    if not raw:
        return 0
    parts = [part.strip() for part in re.split(r"[|,;/]+", raw) if part.strip()]
    return len(parts) if len(parts) >= 2 else 2


def normalize_local_rota_diaria(value: Any) -> str:
    text_value = upper_text(value)
    if text_value.startswith("SERRA"):
        return "SERRA"
    if text_value.startswith("SERT"):
        return "SERTAO"
    return text_value


async def diaria_config_for_programacao(db: AsyncSession, programacao: ProgramacaoDB) -> dict[str, float]:
    local_rota = normalize_local_rota_diaria(programacao.local_rota or programacao.tipo_rota)
    if local_rota not in {"SERRA", "SERTAO"}:
        return {}
    try:
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
        result = await db.execute(
            text("SELECT motorista_valor, ajudante_valor FROM diaria_config WHERE local_rota=:local_rota LIMIT 1"),
            {"local_rota": local_rota},
        )
        row = result.mappings().first()
        if not row:
            return {}
        return {
            "motorista": safe_float(row["motorista_valor"], 0.0),
            "ajudante": safe_float(row["ajudante_valor"], 0.0),
        }
    except Exception:
        return {}


async def diarias_for_programacao(db: AsyncSession, programacao: ProgramacaoDB, nomes_por_id: dict[str, str]) -> DiariasResumo:
    config = await diaria_config_for_programacao(db, programacao)
    diaria_motorista_programacao = safe_float(programacao.diaria_motorista_valor, 0.0)
    diaria_motorista = diaria_motorista_programacao or safe_float(config.get("motorista"), 0.0)
    qtd = calc_qtd_diarias(
        programacao.data_saida or "",
        programacao.hora_saida or "",
        programacao.data_chegada or "",
        programacao.hora_chegada or "",
    )
    qtd_ajudantes = count_ajudantes(programacao.equipe)
    diaria_ajudante = safe_float(config.get("ajudante"), 0.0) if config else 0.0
    if diaria_ajudante <= 0:
        diaria_ajudante = max(diaria_motorista - 10.0, 0.0)
    total_motorista = round(qtd * diaria_motorista, 2)
    total_ajudantes = round(qtd * (diaria_ajudante * qtd_ajudantes), 2)
    equipe = equipe_nomes(programacao.equipe, nomes_por_id) or "-"
    motorista = upper_text(programacao.motorista) or "-"
    return DiariasResumo(
        qtd_diarias=qtd,
        qtd_ajudantes=qtd_ajudantes,
        diaria_motorista=round(diaria_motorista, 2),
        diaria_ajudante=round(diaria_ajudante, 2),
        total_motorista=total_motorista,
        total_ajudantes=total_ajudantes,
        total_geral=round(total_motorista + total_ajudantes, 2),
        observacao_motorista=f"QTD DIARIAS: {qtd:g} | MOTORISTA: {motorista}",
        observacao_ajudantes=f"QTD DIARIAS: {qtd:g} | AJUDANTES: {equipe}",
    )


def normalize_date_or_error(value: str | None, field: str) -> str:
    normalized = normalize_date(value or "")
    if normalized is None:
        raise HTTPException(status_code=422, detail=f"Formato invalido em {field}.")
    return normalized


def normalize_time_or_error(value: str | None, field: str) -> str:
    normalized = normalize_time(value or "")
    if normalized is None:
        raise HTTPException(status_code=422, detail=f"Formato invalido em {field}.")
    return normalized


def pdf_line(pdf: Any, y: float, text: Any, *, x: int = 40, size: int = 9, bold: bool = False) -> float:
    pdf.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    pdf.drawString(x, y, str(text or "")[:126])
    return y - 13


def pdf_money(value: Any) -> str:
    return f"R$ {safe_float(value, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def draw_recebimentos_pdf_page(pdf: Any, bundle: RecebimentoBundleResponse) -> None:
    _width, height = pdf._pagesize
    cab = bundle.cabecalho
    diarias = bundle.diarias
    y = height - 52
    y = pdf_line(pdf, y, f"RECEBIMENTOS - PROGRAMACAO {cab.codigo_programacao}", size=13, bold=True)
    y -= 6
    for line in (
        f"Motorista: {cab.motorista_nome or cab.motorista or '-'}",
        f"Veiculo: {cab.veiculo or '-'} | Rota: {cab.rota or '-'} | Status: {cab.status or '-'}",
        f"Equipe: {cab.equipe_nomes or cab.equipe or '-'}",
        f"Saida: {cab.data_saida or '-'} {cab.hora_saida or ''} | Chegada: {cab.data_chegada or '-'} {cab.hora_chegada or ''}",
        f"NF: {cab.num_nf or '-'} | Prestacao: {cab.prestacao_status or '-'}",
    ):
        y = pdf_line(pdf, y, line)
    y -= 8
    y = pdf_line(pdf, y, "DIARIAS", bold=True)
    for line in (
        f"Valor diaria motorista: {pdf_money(diarias.diaria_motorista)}",
        f"Qtd diarias: {diarias.qtd_diarias:g} | Ajudantes: {diarias.qtd_ajudantes}",
        f"Total motorista: {pdf_money(diarias.total_motorista)} | Total ajudantes: {pdf_money(diarias.total_ajudantes)}",
        f"Total geral diarias: {pdf_money(diarias.total_geral)}",
    ):
        y = pdf_line(pdf, y, line)
    y -= 8
    y = pdf_line(pdf, y, "CLIENTES / RECEBIMENTOS", bold=True)
    y = pdf_line(pdf, y, "COD | CLIENTE | VALOR | FORMA | OBS", size=8, bold=True)
    for cliente in bundle.clientes:
        valor = safe_float(cliente.valor, 0.0)
        if valor <= 0:
            continue
        y = pdf_line(
            pdf,
            y,
            f"{cliente.cod_cliente} | {cliente.nome_cliente} | {pdf_money(valor)} | {cliente.forma_pagamento or '-'} | {cliente.observacao or ''}",
            size=8,
        )
        if y < 60:
            pdf.showPage()
            y = height - 52
            y = pdf_line(pdf, y, "CLIENTES / RECEBIMENTOS (continua)", bold=True)
    y -= 8
    pdf_line(pdf, y, f"TOTAL RECEBIDO: {pdf_money(bundle.total_recebido)}", size=11, bold=True)


async def serialize_bundle(db: AsyncSession, programacao: ProgramacaoDB) -> RecebimentoBundleResponse:
    nomes_por_id = await ajudante_map(db)
    codigo = upper_text(programacao.codigo_programacao)
    diaria_config = await diaria_config_for_programacao(db, programacao)
    diaria_motorista_valor = safe_float(programacao.diaria_motorista_valor, 0.0) or safe_float(diaria_config.get("motorista"), 0.0)
    item_result = await db.execute(
        select(ProgramacaoItemDB)
        .where(func.upper(ProgramacaoItemDB.codigo_programacao) == codigo)
        .order_by(ProgramacaoItemDB.nome_cliente.asc(), ProgramacaoItemDB.id.asc())
    )
    clientes_base: dict[str, str] = {}
    for item in item_result.scalars().all():
        cod = upper_text(item.cod_cliente)
        nome = upper_text(item.nome_cliente)
        if cod and cod not in clientes_base:
            clientes_base[cod] = nome

    rec_result = await db.execute(
        select(RecebimentoDB)
        .where(func.upper(RecebimentoDB.codigo_programacao) == codigo)
        .order_by(RecebimentoDB.id.desc())
    )
    rec_map: dict[str, dict[str, Any]] = {}
    for rec in rec_result.scalars().all():
        cod = upper_text(rec.cod_cliente)
        if not cod:
            continue
        clientes_base.setdefault(cod, upper_text(rec.nome_cliente))
        info = rec_map.setdefault(
            cod,
            {"valor": 0.0, "forma": "", "obs": "", "data": "", "nome": upper_text(rec.nome_cliente)},
        )
        info["valor"] += safe_float(rec.valor, 0.0)
        if not info["forma"] and rec.forma_pagamento:
            info["forma"] = upper_text(rec.forma_pagamento)
        if not info["obs"] and rec.observacao:
            info["obs"] = rec.observacao or ""
        if not info["data"] and rec.data_registro:
            info["data"] = rec.data_registro or ""

    clientes = []
    total = 0.0
    for cod, nome in sorted(clientes_base.items(), key=lambda item: (item[1], item[0])):
        info = rec_map.get(cod, {})
        valor = round(safe_float(info.get("valor"), 0.0), 2)
        total += valor
        clientes.append(
            RecebimentoCliente(
                cod_cliente=cod,
                nome_cliente=upper_text(info.get("nome") or nome),
                valor=valor,
                forma_pagamento=upper_text(info.get("forma")),
                observacao=str(info.get("obs") or ""),
                data_registro=str(info.get("data") or "")[:19],
            )
        )

    cabecalho = RecebimentoCabecalho(
        codigo_programacao=codigo,
        motorista=upper_text(programacao.motorista),
        motorista_nome=upper_text(programacao.motorista),
        veiculo=upper_text(programacao.veiculo),
        equipe=programacao.equipe or "",
        equipe_nomes=equipe_nomes(programacao.equipe, nomes_por_id),
        rota=upper_text(programacao.local_rota or programacao.tipo_rota),
        status=status_ref(programacao),
        prestacao_status=upper_text(programacao.prestacao_status or "PENDENTE"),
        num_nf=upper_text(programacao.num_nf or programacao.nf_numero),
        data_saida=programacao.data_saida or "",
        hora_saida=programacao.hora_saida or "",
        data_chegada=programacao.data_chegada or "",
        hora_chegada=programacao.hora_chegada or "",
        diaria_motorista_valor=round(diaria_motorista_valor, 2),
        fechada=is_closed(programacao),
    )
    return RecebimentoBundleResponse(
        cabecalho=cabecalho,
        diarias=await diarias_for_programacao(db, programacao, nomes_por_id),
        clientes=clientes,
        total_recebido=round(total, 2),
    )


async def ensure_recebimento_cliente_item(db: AsyncSession, codigo: str, cod_cliente: str, nome_cliente: str) -> bool:
    item_result = await db.execute(
        select(ProgramacaoItemDB)
        .where(func.upper(ProgramacaoItemDB.codigo_programacao) == codigo, func.upper(ProgramacaoItemDB.cod_cliente) == cod_cliente)
        .limit(1)
    )
    item = item_result.scalar_one_or_none()
    if item:
        if nome_cliente and upper_text(item.nome_cliente) != nome_cliente:
            item.nome_cliente = nome_cliente
        return False
    db.add(
        ProgramacaoItemDB(
            codigo_programacao=codigo,
            cod_cliente=cod_cliente,
            nome_cliente=nome_cliente,
            qnt_caixas=0,
            kg=0,
            preco=0,
            observacao="CLIENTE MANUAL RECEBIMENTOS",
        )
    )
    return True


@router.get("/programacoes", response_model=list[RecebimentoProgramacaoOption])
async def listar_programacoes_recebimentos(
    limit: int = 300,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(select(ProgramacaoDB).order_by(ProgramacaoDB.id.desc()).limit(max(min(limit, 500), 1)))
    out = []
    for programacao in result.scalars().all():
        status = status_ref(programacao)
        if status in CANCELLED_STATUSES:
            continue
        if is_closed(programacao):
            continue
        out.append(
            RecebimentoProgramacaoOption(
                codigo_programacao=upper_text(programacao.codigo_programacao),
                motorista=upper_text(programacao.motorista),
                veiculo=upper_text(programacao.veiculo),
                status=status,
                prestacao_status=upper_text(programacao.prestacao_status or "PENDENTE"),
                fechada=is_closed(programacao),
            )
        )
    return out


@router.get("/{codigo_programacao}/bundle", response_model=RecebimentoBundleResponse)
async def carregar_recebimentos_bundle(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_open_recebimentos(programacao)
    return await serialize_bundle(db, programacao)


@router.get("/{codigo_programacao}/pdf")
async def recebimentos_pdf(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    bundle = await serialize_bundle(db, programacao)
    if not any(safe_float(cliente.valor, 0.0) > 0 for cliente in bundle.clientes):
        raise HTTPException(status_code=409, detail="Nao ha clientes pagantes para imprimir.")
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    draw_recebimentos_pdf_page(pdf, bundle)
    pdf.save()
    buffer.seek(0)
    safe_name = upper_text(programacao.codigo_programacao).replace("/", "_") or "PROGRAMACAO"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="RECEBIMENTOS_{safe_name}.pdf"'},
    )


@router.put("/{codigo_programacao}/cabecalho", response_model=RecebimentoBundleResponse)
async def salvar_cabecalho_recebimentos(
    codigo_programacao: str,
    payload: RecebimentoCabecalhoPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)

    programacao.data_saida = normalize_date_or_error(payload.data_saida, "data_saida")
    programacao.hora_saida = normalize_time_or_error(payload.hora_saida, "hora_saida")
    programacao.data_chegada = normalize_date_or_error(payload.data_chegada, "data_chegada")
    programacao.hora_chegada = normalize_time_or_error(payload.hora_chegada, "hora_chegada")
    programacao.diaria_motorista_valor = safe_float(payload.diaria_motorista_valor, 0.0)
    await registrar_roteiro_operacional(
        db,
        tipo_evento="RECEBIMENTOS_CABECALHO_WEB",
        codigo_programacao=upper_text(programacao.codigo_programacao),
        origem="WEB",
        destino="RECEBIMENTOS",
        motorista_nome=programacao.motorista,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao="Cabecalho de recebimentos salvo",
        payload=payload.model_dump(),
    )

    record_audit_log(
        db,
        action="recebimentos_cabecalho_salvo",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        ip_address=client_ip_from_request(request),
        metadata={"diaria_motorista_valor": programacao.diaria_motorista_valor},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.post("/{codigo_programacao}/recebimentos", response_model=RecebimentoCliente, status_code=status.HTTP_201_CREATED)
async def salvar_recebimento(
    codigo_programacao: str,
    payload: RecebimentoPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)

    forma = upper_text(payload.forma_pagamento)
    if forma not in FORMAS_VALIDAS:
        raise HTTPException(status_code=422, detail="Forma de pagamento invalida.")

    codigo = upper_text(programacao.codigo_programacao)
    cod_cliente = upper_text(payload.cod_cliente)
    nome_cliente = upper_text(payload.nome_cliente)
    cliente_manual_criado = await ensure_recebimento_cliente_item(db, codigo, cod_cliente, nome_cliente)
    await db.execute(
        delete(RecebimentoDB).where(
            func.upper(RecebimentoDB.codigo_programacao) == codigo,
            func.upper(RecebimentoDB.cod_cliente) == cod_cliente,
        )
    )
    rec = RecebimentoDB(
        codigo_programacao=codigo,
        cod_cliente=cod_cliente,
        nome_cliente=nome_cliente,
        valor=round(safe_float(payload.valor, 0.0), 2),
        forma_pagamento=forma,
        observacao=upper_text(payload.observacao),
        num_nf=upper_text(payload.num_nf or programacao.num_nf or programacao.nf_numero),
        data_registro=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    db.add(rec)
    await registrar_roteiro_operacional(
        db,
        tipo_evento="RECEBIMENTO_WEB",
        codigo_programacao=codigo,
        origem="WEB",
        destino=forma,
        motorista_nome=programacao.motorista,
        pedido=getattr(payload, "pedido", "") or "",
        cod_cliente=rec.cod_cliente,
        cliente_nome=rec.nome_cliente,
        nf_numero=rec.num_nf,
        data_hora=rec.data_registro,
        observacao=rec.observacao,
        payload=payload.model_dump(),
    )
    record_audit_log(
        db,
        action="recebimento_salvo",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata={
            "cod_cliente": rec.cod_cliente,
            "valor": rec.valor,
            "forma_pagamento": forma,
            "modo": "substituir_cliente",
            "cliente_manual_criado": cliente_manual_criado,
        },
    )
    await db.commit()
    return RecebimentoCliente(
        cod_cliente=rec.cod_cliente,
        nome_cliente=rec.nome_cliente,
        valor=rec.valor,
        forma_pagamento=rec.forma_pagamento or "",
        observacao=rec.observacao or "",
        data_registro=rec.data_registro or "",
    )


@router.delete("/{codigo_programacao}/recebimentos/{cod_cliente}", response_model=RecebimentoBundleResponse)
async def zerar_recebimento_cliente(
    codigo_programacao: str,
    cod_cliente: str,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    codigo = upper_text(programacao.codigo_programacao)
    cod = upper_text(cod_cliente)
    await db.execute(
        delete(RecebimentoDB).where(
            func.upper(RecebimentoDB.codigo_programacao) == codigo,
            func.upper(RecebimentoDB.cod_cliente) == cod,
        )
    )
    await registrar_roteiro_operacional(
        db,
        tipo_evento="RECEBIMENTO_ZERADO_WEB",
        codigo_programacao=codigo,
        origem="WEB",
        destino="RECEBIMENTOS",
        motorista_nome=programacao.motorista,
        cod_cliente=cod,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao="Recebimento zerado",
        payload={"cod_cliente": cod},
    )
    record_audit_log(
        db,
        action="recebimento_zerado",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"cod_cliente": cod},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.post("/{codigo_programacao}/clientes/manual", response_model=RecebimentoCliente, status_code=status.HTTP_201_CREATED)
async def inserir_cliente_manual(
    codigo_programacao: str,
    payload: ClienteManualPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    codigo = upper_text(programacao.codigo_programacao)
    cod = upper_text(payload.cod_cliente)
    nome = upper_text(payload.nome_cliente)
    await ensure_recebimento_cliente_item(db, codigo, cod, nome)
    record_audit_log(
        db,
        action="recebimento_cliente_manual",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata={"cod_cliente": cod, "nome_cliente": nome},
    )
    await db.commit()
    return RecebimentoCliente(cod_cliente=cod, nome_cliente=nome)
