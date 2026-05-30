# backend/api/v1/endpoints/importar_vendas.py
"""
Importar Vendas endpoints mirroring the desktop import/selection/link flow.
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
from fastapi import APIRouter, Depends, File, HTTPException, Request, UploadFile, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import func, or_, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.excel_helpers import guess_col
from app.utils.formatters import normalize_date, safe_float, safe_int
from backend.api.v1.endpoints.programacao import (
    get_programacao_by_codigo,
    items_for_programacao,
    upper_text,
)
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import ClienteDB, ProdutoDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemDB
from backend.models.user import User
from backend.models.venda_importada import VendaImportadaDB
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()

VINCULAVEL_STATUSES = {"", "ATIVA", "ABERTA", "PROGRAMADA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"}
BLOCKED_STATUSES = {"CANCELADA", "CANCELADO", "FINALIZADA", "FINALIZADO", "EM ENTREGAS", "EM_ENTREGAS"}


def produto_codigo_base(value: Any) -> str:
    text_value = upper_text(value)
    code = "".join(ch if ch.isalnum() else "-" for ch in text_value).strip("-")
    while "--" in code:
        code = code.replace("--", "-")
    return (code or "PRODUTO")[:40]


class VendaImportadaPayload(BaseModel):
    pedido: Any = None
    data_venda: Any = None
    cliente: Any = None
    nome_cliente: Any = None
    vendedor: Any = None
    produto: Any = None
    vr_total: Any = 0
    qnt: Any = 0
    cidade: Any = None
    valor_unitario: Any = None
    observacao: Any = None


class ImportarVendasPayload(BaseModel):
    rows: list[VendaImportadaPayload] = Field(default_factory=list)


class IdsPayload(BaseModel):
    ids: list[int] = Field(default_factory=list)

    @field_validator("ids")
    @classmethod
    def clean_ids(cls, value):
        return sorted({safe_int(item, 0) for item in value if safe_int(item, 0) > 0})


class VincularVendasPayload(IdsPayload):
    codigo_programacao: str = Field(min_length=1, max_length=40)
    caixas_por_venda: dict[str, int] = Field(default_factory=dict)

    @field_validator("codigo_programacao", mode="before")
    @classmethod
    def strip_codigo(cls, value):
        return upper_text(value)


class VendaCaixasPayload(BaseModel):
    qnt_caixas: int = Field(ge=0, le=100000)


class VendaImportadaResponse(BaseModel):
    id: int
    pedido: str = ""
    data_venda: str = ""
    cliente: str = ""
    nome_cliente: str = ""
    vendedor: str = ""
    produto: str = ""
    vr_total: float = 0
    qnt: float = 0
    qnt_caixas: int = 0
    cidade: str = ""
    valor_unitario: float = 0
    observacao: str = ""
    selecionada: int = 0
    usada: int = 0
    usada_em: str = ""
    codigo_programacao: str = ""


class ImportarVendasResult(BaseModel):
    importadas: int = 0
    ignoradas: int = 0
    invalidas: int = 0
    opcionais_ausentes: list[str] = Field(default_factory=list)


class VincularVendasResult(BaseModel):
    codigo_programacao: str
    vendas_vinculadas: int
    itens_adicionados: int
    itens_total: int
    total_caixas: int
    quilos: float


def serialize_venda(venda: VendaImportadaDB) -> VendaImportadaResponse:
    return VendaImportadaResponse(
        id=venda.id,
        pedido=venda.pedido or "",
        data_venda=venda.data_venda or "",
        cliente=venda.cliente or "",
        nome_cliente=venda.nome_cliente or "",
        vendedor=venda.vendedor or "",
        produto=venda.produto or "",
        vr_total=safe_float(venda.vr_total, 0.0),
        qnt=safe_float(venda.qnt, 0.0),
        qnt_caixas=safe_int(getattr(venda, "qnt_caixas", 0), 0),
        cidade=venda.cidade or "",
        valor_unitario=safe_float(venda.valor_unitario, 0.0),
        observacao=venda.observacao or "",
        selecionada=safe_int(venda.selecionada, 0),
        usada=safe_int(venda.usada, 0),
        usada_em=venda.usada_em or "",
        codigo_programacao=venda.codigo_programacao or "",
    )


def excel_text(value: Any) -> str:
    text = str(value or "").strip()
    if upper_text(text) in {"", "NAN", "NAT", "NONE", "NULL", "<NA>"}:
        return ""
    return text


def clean_pedido(value: Any) -> str:
    text = excel_text(value)
    if not text:
        return ""
    try:
        number = float(text.replace(",", "."))
        if abs(number - int(number)) < 1e-9:
            return str(int(number))
        return str(number).rstrip("0").rstrip(".")
    except Exception:
        return text


def normalize_data_venda(value: Any) -> str:
    text = excel_text(value)
    if not text:
        return ""
    normalized = normalize_date(text)
    return normalized if normalized is not None else text


def normalized_import_row(raw: dict[str, Any]) -> dict[str, Any] | None:
    pedido = upper_text(clean_pedido(raw.get("pedido")))
    cliente = upper_text(excel_text(raw.get("cliente")))
    nome_cliente = upper_text(excel_text(raw.get("nome_cliente")))
    produto = upper_text(excel_text(raw.get("produto")))
    if not pedido or not cliente or not nome_cliente or not produto:
        return None

    qnt = safe_float(raw.get("qnt"), 0.0)
    vr_total = safe_float(raw.get("vr_total"), 0.0)
    valor_unitario = safe_float(raw.get("valor_unitario"), 0.0)
    if valor_unitario <= 0 and qnt > 0:
        valor_unitario = vr_total / qnt

    return {
        "pedido": pedido,
        "data_venda": normalize_data_venda(raw.get("data_venda")),
        "cliente": cliente,
        "nome_cliente": nome_cliente,
        "vendedor": upper_text(excel_text(raw.get("vendedor"))),
        "produto": produto,
        "vr_total": vr_total,
        "qnt": qnt,
        "qnt_caixas": max(safe_int(round(qnt), 0), 1 if qnt > 0 else 0),
        "cidade": upper_text(excel_text(raw.get("cidade"))),
        "valor_unitario": safe_float(valor_unitario, 0.0),
        "observacao": upper_text(excel_text(raw.get("observacao"))),
    }


def row_key(row: dict[str, Any] | VendaImportadaDB) -> tuple[str, str, str, str]:
    if isinstance(row, VendaImportadaDB):
        return (
            upper_text(row.pedido),
            upper_text(row.cliente),
            upper_text(row.produto),
            str(row.data_venda or "").strip(),
        )
    return (
        upper_text(row.get("pedido")),
        upper_text(row.get("cliente")),
        upper_text(row.get("produto")),
        str(row.get("data_venda") or "").strip(),
    )


async def existing_venda_keys(db: AsyncSession) -> set[tuple[str, str, str, str]]:
    result = await db.execute(select(VendaImportadaDB))
    return {row_key(item) for item in result.scalars().all()}


async def produto_id_for_nome(db: AsyncSession, nome: str) -> int | None:
    nome_norm = upper_text(nome)
    if not nome_norm:
        return None
    result = await db.execute(select(ProdutoDB).where(func.upper(func.coalesce(ProdutoDB.nome, "")) == nome_norm).limit(1))
    produto = result.scalar_one_or_none()
    if produto:
        return int(produto.id)
    codigo = produto_codigo_base(nome_norm)
    codigo_final = codigo
    for suffix in range(0, 1000):
        candidate = codigo_final if suffix == 0 else f"{codigo[:34]}-{suffix:03d}"
        result = await db.execute(select(ProdutoDB).where(func.upper(func.coalesce(ProdutoDB.codigo, "")) == candidate).limit(1))
        if result.scalar_one_or_none():
            continue
        codigo_final = candidate
        break
    else:
        raise HTTPException(status_code=409, detail="Nao foi possivel gerar codigo unico para o produto.")
    produto = ProdutoDB(
        codigo=codigo_final,
        nome=nome_norm,
        descricao="Cadastro automatico criado pela importacao de vendas.",
        categoria="AVES",
        unidade="KG",
        unidade_estoque="KG",
        controla_estoque_fisico=1,
        controla_estoque_fiscal=1,
        status="ATIVO",
    )
    db.add(produto)
    await db.flush()
    return int(produto.id)


async def import_rows(db: AsyncSession, rows: list[dict[str, Any]]) -> ImportarVendasResult:
    keys = await existing_venda_keys(db)
    result = ImportarVendasResult()
    for raw in rows:
        row = normalized_import_row(raw)
        if not row:
            result.invalidas += 1
            continue
        key = row_key(row)
        if key in keys:
            result.ignoradas += 1
            continue
        produto_id = await produto_id_for_nome(db, row.get("produto"))
        db.add(VendaImportadaDB(**row, produto_id=produto_id, selecionada=0, usada=0, usada_em="", codigo_programacao=""))
        keys.add(key)
        result.importadas += 1
    return result


def rows_from_dataframe(df: pd.DataFrame) -> tuple[list[dict[str, Any]], list[str]]:
    col_pedido = guess_col(df.columns, ["numero pedido", "num pedido", "n pedido", "pedido"])
    col_data = guess_col(df.columns, ["data venda", "data", "dt"])
    col_cliente = guess_col(df.columns, ["cod cliente", "codigo cliente", "cliente", "cod"])
    col_nome = guess_col(df.columns, ["nome completo", "nome cliente", "razao", "nome"])
    col_produto = guess_col(df.columns, ["descricao do produto", "produto", "descr", "item"])
    col_vr_total = guess_col(df.columns, ["vr. total", "vr total", "valor total", "total"])
    col_qnt = guess_col(df.columns, ["qnt", "qtd", "quantidade"])
    col_cidade = guess_col(df.columns, ["cidade", "municipio"])
    col_vendedor = guess_col(df.columns, ["nome do vendedor", "vendedor", "vend"])
    col_obs = guess_col(df.columns, ["obs", "observ", "observacao"])

    missing = []
    if not col_pedido:
        missing.append("Numero Pedido")
    if not col_cliente:
        missing.append("Cliente")
    if not col_nome:
        missing.append("Nome Completo")
    if not col_produto:
        missing.append("Descricao do Produto")
    if not col_vr_total:
        missing.append("Vr. Total")
    if not col_qnt:
        missing.append("Qnt.")
    if missing:
        raise HTTPException(status_code=422, detail="Nao identifiquei as colunas: " + ", ".join(missing))

    opcionais = []
    if not col_data:
        opcionais.append("Data")
    if not col_cidade:
        opcionais.append("Cidade")
    if not col_vendedor:
        opcionais.append("Nome do Vendedor")
    if not col_obs:
        opcionais.append("Observacao")

    rows = []
    for _, source in df.iterrows():
        rows.append(
            {
                "pedido": source.get(col_pedido, ""),
                "data_venda": source.get(col_data, "") if col_data else "",
                "cliente": source.get(col_cliente, ""),
                "nome_cliente": source.get(col_nome, ""),
                "vendedor": source.get(col_vendedor, "") if col_vendedor else "",
                "produto": source.get(col_produto, ""),
                "vr_total": source.get(col_vr_total, 0),
                "qnt": source.get(col_qnt, 0),
                "cidade": source.get(col_cidade, "") if col_cidade else "",
                "observacao": source.get(col_obs, "") if col_obs else "",
            }
        )
    return rows, opcionais


def is_programacao_vinculavel(programacao: ProgramacaoDB) -> bool:
    status_value = upper_text(programacao.status)
    operational = upper_text(programacao.status_operacional)
    prestacao = upper_text(programacao.prestacao_status)
    if prestacao == "FECHADA":
        return False
    if status_value in BLOCKED_STATUSES or operational in BLOCKED_STATUSES:
        return False
    return status_value in VINCULAVEL_STATUSES or operational in VINCULAVEL_STATUSES


async def vendas_by_ids(db: AsyncSession, ids: list[int]) -> list[VendaImportadaDB]:
    if not ids:
        return []
    result = await db.execute(
        select(VendaImportadaDB)
        .where(
            VendaImportadaDB.id.in_(ids),
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
        .order_by(VendaImportadaDB.id.asc())
    )
    return list(result.scalars().all())


async def endereco_map_for_vendas(db: AsyncSession, vendas: list[VendaImportadaDB]) -> dict[str, str]:
    codigos = sorted({upper_text(venda.cliente) for venda in vendas if upper_text(venda.cliente)})
    if not codigos:
        return {}
    result = await db.execute(
        select(ClienteDB).where(func.upper(func.coalesce(ClienteDB.cod_cliente, "")).in_(codigos))
    )
    return {upper_text(item.cod_cliente): upper_text(item.endereco) for item in result.scalars().all()}


def venda_to_programacao_item(venda: VendaImportadaDB, endereco_map: dict[str, str], caixas_por_venda: dict[str, int]) -> dict[str, Any]:
    qnt = safe_float(venda.qnt, 0.0)
    caixa_payload = safe_int(caixas_por_venda.get(str(venda.id)), 0)
    caixa_salva = safe_int(getattr(venda, "qnt_caixas", 0), 0)
    caixas = caixa_payload if caixa_payload > 0 else (caixa_salva if caixa_salva > 0 else max(safe_int(round(qnt), 0), 1 if qnt > 0 else 1))
    preco = (safe_float(venda.vr_total, 0.0) / qnt) if qnt > 0 else safe_float(venda.valor_unitario, 0.0)
    cod_cliente = upper_text(venda.cliente)
    return {
        "cod_cliente": cod_cliente,
        "nome_cliente": upper_text(venda.nome_cliente),
        "produto": upper_text(venda.produto),
        "endereco": upper_text(endereco_map.get(cod_cliente) or venda.cidade),
        "qnt_caixas": caixas,
        "kg": 0.0,
        "preco": safe_float(preco, 0.0),
        "vendedor": upper_text(venda.vendedor),
        "pedido": upper_text(venda.pedido),
        "observacao": upper_text(venda.observacao),
    }


def item_key(item: ProgramacaoItemDB | dict[str, Any]) -> tuple[str, str, str]:
    if isinstance(item, ProgramacaoItemDB):
        return upper_text(item.cod_cliente), upper_text(item.pedido), upper_text(item.produto)
    return upper_text(item.get("cod_cliente")), upper_text(item.get("pedido")), upper_text(item.get("produto"))


@router.get("/", response_model=list[VendaImportadaResponse])
async def list_vendas_importadas(
    busca: str = "",
    skip: int = 0,
    limit: int = 500,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    stmt = select(VendaImportadaDB).where(
        func.coalesce(VendaImportadaDB.usada, 0) == 0,
        func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
    )
    query = upper_text(busca)
    if query:
        pattern = f"%{query}%"
        stmt = stmt.where(
            or_(
                func.upper(func.coalesce(VendaImportadaDB.pedido, "")).like(pattern),
                func.upper(func.coalesce(VendaImportadaDB.cliente, "")).like(pattern),
                func.upper(func.coalesce(VendaImportadaDB.nome_cliente, "")).like(pattern),
                func.upper(func.coalesce(VendaImportadaDB.vendedor, "")).like(pattern),
                func.upper(func.coalesce(VendaImportadaDB.produto, "")).like(pattern),
            )
        )
    result = await db.execute(stmt.order_by(VendaImportadaDB.id.desc()).offset(skip).limit(limit))
    return [serialize_venda(item) for item in result.scalars().all()]


@router.post("/importar", response_model=ImportarVendasResult, status_code=status.HTTP_201_CREATED)
async def importar_vendas_json(
    payload: ImportarVendasPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await import_rows(db, [item.model_dump() for item in payload.rows])
    record_audit_log(
        db,
        action="vendas_importadas_importadas",
        actor_user=current_user,
        entity_type="vendas_importadas",
        ip_address=client_ip_from_request(request),
        metadata=result.model_dump(),
    )
    await db.commit()
    return result


@router.post("/upload", response_model=ImportarVendasResult, status_code=status.HTTP_201_CREATED)
async def importar_vendas_upload(
    request: Request,
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    filename = str(file.filename or "").lower()
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=422, detail="Arquivo vazio.")
    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(BytesIO(contents))
        else:
            engine = "xlrd" if filename.endswith(".xls") and not filename.endswith(".xlsx") else "openpyxl"
            df = pd.read_excel(BytesIO(contents), engine=engine)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Nao foi possivel ler o arquivo: {exc}") from exc

    rows, opcionais = rows_from_dataframe(df)
    result = await import_rows(db, rows)
    result.opcionais_ausentes = opcionais
    record_audit_log(
        db,
        action="vendas_importadas_excel",
        actor_user=current_user,
        entity_type="vendas_importadas",
        ip_address=client_ip_from_request(request),
        metadata={**result.model_dump(), "arquivo": file.filename},
    )
    await db.commit()
    return result


@router.get("/programacoes-vinculo", response_model=list[str])
async def list_programacoes_vinculo(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(select(ProgramacaoDB).order_by(ProgramacaoDB.id.desc()).limit(300))
    codigos = []
    for item in result.scalars().all():
        codigo = upper_text(item.codigo_programacao)
        if codigo and is_programacao_vinculavel(item):
            codigos.append(codigo)
    return list(dict.fromkeys(codigos))


@router.post("/marcar-todas", response_model=dict[str, int])
async def marcar_todas(
    selected: int = 1,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(
        select(VendaImportadaDB).where(
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
    )
    vendas = list(result.scalars().all())
    for venda in vendas:
        venda.selecionada = 1 if selected else 0
    await db.commit()
    return {"updated": len(vendas)}


@router.post("/marcar-ids", response_model=dict[str, int])
async def marcar_ids(
    payload: IdsPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    vendas = await vendas_by_ids(db, payload.ids)
    for venda in vendas:
        venda.selecionada = 1
    await db.commit()
    return {"updated": len(vendas)}


@router.put("/{venda_id}/caixas", response_model=VendaImportadaResponse)
async def atualizar_caixas_venda(
    venda_id: int,
    payload: VendaCaixasPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(
        select(VendaImportadaDB).where(
            VendaImportadaDB.id == venda_id,
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
    )
    venda = result.scalar_one_or_none()
    if not venda:
        raise HTTPException(status_code=404, detail="Venda livre nao encontrada.")
    venda.qnt_caixas = safe_int(payload.qnt_caixas, 0)
    await db.commit()
    await db.refresh(venda)
    return serialize_venda(venda)


@router.delete("/ids", response_model=dict[str, int])
async def delete_ids(
    payload: IdsPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    vendas = await vendas_by_ids(db, payload.ids)
    deleted = len(vendas)
    for venda in vendas:
        await db.delete(venda)
    await db.commit()
    return {"deleted": deleted}


@router.delete("/", response_model=dict[str, int])
async def limpar_tudo(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(
        select(VendaImportadaDB).where(
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
    )
    vendas = list(result.scalars().all())
    for venda in vendas:
        await db.delete(venda)
    await db.commit()
    return {"deleted": len(vendas)}


@router.post("/vincular", response_model=VincularVendasResult)
async def vincular_vendas_programacao(
    payload: VincularVendasPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, payload.codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail=f"Programacao nao encontrada: {payload.codigo_programacao}")
    if not is_programacao_vinculavel(programacao):
        raise HTTPException(status_code=409, detail=f"Programacao {payload.codigo_programacao} nao pode receber vendas.")

    vendas = await vendas_by_ids(db, payload.ids)
    if not vendas:
        raise HTTPException(status_code=422, detail="Selecione uma ou mais vendas livres para vincular.")

    existentes = await items_for_programacao(db, programacao.codigo_programacao)
    seen = {item_key(item) for item in existentes}
    endereco_map = await endereco_map_for_vendas(db, vendas)
    added = 0
    for venda in vendas:
        item_data = venda_to_programacao_item(venda, endereco_map, payload.caixas_por_venda)
        key = item_key(item_data)
        if not item_data["cod_cliente"] or not item_data["nome_cliente"] or key in seen:
            continue
        db.add(
            ProgramacaoItemDB(
                codigo_programacao=upper_text(programacao.codigo_programacao),
                cod_cliente=item_data["cod_cliente"],
                nome_cliente=item_data["nome_cliente"],
                produto_id=int(getattr(venda, "produto_id", 0) or 0) or await produto_id_for_nome(db, item_data["produto"]),
                produto=item_data["produto"],
                endereco=item_data["endereco"] or None,
                qnt_caixas=safe_int(item_data["qnt_caixas"], 0),
                kg=safe_float(item_data["kg"], 0.0),
                preco=safe_float(item_data["preco"], 0.0),
                vendedor=item_data["vendedor"] or None,
                pedido=item_data["pedido"] or None,
                observacao=item_data["observacao"] or None,
            )
        )
        seen.add(key)
        added += 1

    usada_em = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for venda in vendas:
        venda.usada = 1
        venda.usada_em = usada_em
        venda.codigo_programacao = upper_text(programacao.codigo_programacao)
        venda.selecionada = 0

    await db.flush()
    final_items = await items_for_programacao(db, programacao.codigo_programacao)
    total_caixas = sum(safe_int(item.qnt_caixas, 0) for item in final_items)
    total_quilos = round(sum(safe_float(item.kg, 0.0) for item in final_items), 2)
    programacao.total_caixas = total_caixas
    programacao.caixas_carregadas = total_caixas
    programacao.qnt_cx_carregada = total_caixas
    programacao.quilos = total_quilos
    programacao.nf_caixas = total_caixas
    if upper_text(programacao.tipo_estimativa) == "KG" and total_quilos > 0:
        programacao.nf_kg = total_quilos
    programacao.usuario_ultima_edicao = upper_text(current_user.nome or current_user.username) or "ADMIN"

    record_audit_log(
        db,
        action="vendas_importadas_vinculadas",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        ip_address=client_ip_from_request(request),
        metadata={"ids": payload.ids, "vendas": len(vendas), "itens_adicionados": added},
    )
    await db.commit()
    return VincularVendasResult(
        codigo_programacao=upper_text(programacao.codigo_programacao),
        vendas_vinculadas=len(vendas),
        itens_adicionados=added,
        itens_total=len(final_items),
        total_caixas=total_caixas,
        quilos=total_quilos,
    )


@router.post("/{venda_id}/toggle-selecao", response_model=VendaImportadaResponse)
async def toggle_selecao(
    venda_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(
        select(VendaImportadaDB).where(
            VendaImportadaDB.id == venda_id,
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
    )
    venda = result.scalar_one_or_none()
    if not venda:
        raise HTTPException(status_code=404, detail="Venda importada nao encontrada ou ja vinculada.")
    venda.selecionada = 0 if safe_int(venda.selecionada, 0) == 1 else 1
    await db.commit()
    await db.refresh(venda)
    return serialize_venda(venda)
