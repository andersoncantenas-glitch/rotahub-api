# backend/api/v1/endpoints/programacao.py
"""
Programacao endpoints that mirror the desktop ProgramacaoPage core flow.
"""
from __future__ import annotations

import re
import unicodedata
import json
from datetime import datetime, timedelta
from io import BytesIO
from math import asin, cos, radians, sin, sqrt
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query, Request, status
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import delete, func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import normalize_time, safe_float, safe_int
from app.utils.validators import normalize_phone
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import AjudanteDB, ClienteDB, EscalaFolgaDB, MotoristaDB, ProdutoDB, VeiculoDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemDB
from backend.models.user import User
from backend.models.venda_importada import VendaImportadaDB
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()


EDITABLE_STATUSES = {"", "ATIVA", "EM ROTA", "EM_ROTA"}
DELETABLE_STATUSES = {"", "ATIVA"}
RESOURCE_FINAL_STATUSES = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}


def is_active_status(value: Any) -> bool:
    status_value = upper_text(value)
    return not status_value or status_value == "ATIVO"


def programacao_reserva_recurso(programacao: ProgramacaoDB) -> bool:
    prestacao = upper_text(getattr(programacao, "prestacao_status", None) or "PENDENTE")
    if prestacao == "FECHADA":
        return False
    status_value = upper_text(getattr(programacao, "status", None))
    status_operacional = upper_text(getattr(programacao, "status_operacional", None))
    if status_value in RESOURCE_FINAL_STATUSES or status_operacional in RESOURCE_FINAL_STATUSES:
        return False
    if safe_int(getattr(programacao, "finalizada_no_app", 0), 0) == 1:
        return False
    return True


def produto_codigo_base(value: Any) -> str:
    text_value = upper_text(value)
    code = "".join(ch if ch.isalnum() else "-" for ch in text_value).strip("-")
    while "--" in code:
        code = code.replace("--", "-")
    return (code or "PRODUTO")[:40]


class ProgramacaoItemPayload(BaseModel):
    cod_cliente: str = Field(min_length=1, max_length=80)
    nome_cliente: str = Field(min_length=1, max_length=180)
    produto_id: int | None = Field(default=None, ge=1)
    produto: str | None = Field(default=None, max_length=140)
    endereco: str | None = Field(default=None, max_length=220)
    qnt_caixas: int = Field(default=0, ge=0)
    kg: float = Field(default=0, ge=0)
    preco: float = Field(default=0, ge=0)
    vendedor: str | None = Field(default=None, max_length=160)
    pedido: str | None = Field(default=None, max_length=120)
    obs: str | None = Field(default=None, max_length=300)
    ordem_sugerida: int | None = Field(default=None, ge=0)
    distancia: float | None = Field(default=None, ge=0)
    confianca_localizacao: float | None = Field(default=None, ge=0)
    carga_raiz_programacao: str | None = Field(default=None, max_length=40)
    carga_origem_imediata: str | None = Field(default=None, max_length=40)
    transferencia_origem_id: str | None = Field(default=None, max_length=80)

    @field_validator(
        "cod_cliente",
        "nome_cliente",
        "produto",
        "endereco",
        "vendedor",
        "pedido",
        "obs",
        "carga_raiz_programacao",
        "carga_origem_imediata",
        "transferencia_origem_id",
        mode="before",
    )
    @classmethod
    def strip_text(cls, value):
        if value is None:
            return None
        return str(value).strip()


class ProgramacaoPayload(BaseModel):
    codigo_programacao: str | None = Field(default=None, max_length=40)
    motorista: str = Field(min_length=1, max_length=180)
    motorista_codigo: str | None = Field(default=None, max_length=80)
    veiculo: str = Field(min_length=1, max_length=40)
    ajudantes: list[str] = Field(default_factory=list, max_length=2)
    equipe: str | None = Field(default=None, max_length=220)
    local_rota: str = Field(min_length=1, max_length=40)
    tipo_estimativa: str = Field(default="KG", max_length=2)
    kg_estimado: float = Field(default=0, ge=0)
    caixas_estimado: int = Field(default=0, ge=0)
    operacao_tipo: str | None = Field(default=None, max_length=40)
    transbordo_modalidade: str | None = Field(default=None, max_length=40)
    transbordo_observacao: str | None = Field(default=None, max_length=300)
    local_carregamento: str = Field(min_length=1, max_length=160)
    adiantamento: float = Field(default=0, ge=0)
    adiantamento_origem: str | None = Field(default=None, max_length=160)
    itens: list[ProgramacaoItemPayload] | None = None
    venda_ids: list[int] = Field(default_factory=list)

    @field_validator(
        "codigo_programacao",
        "motorista",
        "motorista_codigo",
        "veiculo",
        "equipe",
        "local_rota",
        "tipo_estimativa",
        "operacao_tipo",
        "transbordo_modalidade",
        "transbordo_observacao",
        "local_carregamento",
        "adiantamento_origem",
        mode="before",
    )
    @classmethod
    def strip_header_text(cls, value):
        if value is None:
            return None
        return str(value).strip()


class ProgramacaoItemResponse(BaseModel):
    id: int | None = None
    cod_cliente: str
    nome_cliente: str
    produto_id: int | None = None
    produto: str | None = None
    endereco: str | None = None
    qnt_caixas: int = 0
    kg: float = 0
    preco: float = 0
    vendedor: str | None = None
    pedido: str | None = None
    obs: str | None = None
    ordem_sugerida: int | None = None
    distancia: float | None = None
    confianca_localizacao: float | None = None
    carga_raiz_programacao: str | None = None
    carga_origem_imediata: str | None = None
    transferencia_origem_id: str | None = None


class ProgramacaoResponse(BaseModel):
    id: int
    codigo_programacao: str
    data_criacao: str | None = None
    motorista: str
    motorista_codigo: str | None = None
    veiculo: str
    equipe: str | None = None
    ajudantes: list[str] = Field(default_factory=list)
    kg_estimado: float = 0
    tipo_estimativa: str = "KG"
    caixas_estimado: int = 0
    operacao_tipo: str = "VENDA"
    transbordo_modalidade: str | None = None
    transbordo_observacao: str | None = None
    status: str | None = None
    status_operacional: str | None = None
    prestacao_status: str | None = None
    local_rota: str | None = None
    local_carregamento: str | None = None
    adiantamento: float = 0
    adiantamento_origem: str | None = None
    total_caixas: int = 0
    quilos: float = 0
    itens: list[ProgramacaoItemResponse] = Field(default_factory=list)


class ProgramacaoOptionsResponse(BaseModel):
    motoristas: list[dict[str, Any]]
    veiculos: list[dict[str, Any]]
    ajudantes: list[dict[str, Any]]
    proximo_codigo: str


class ProgramacaoVendaSelecionadaItem(ProgramacaoItemResponse):
    venda_id: int


class ProgramacaoVendasSelecionadasResponse(BaseModel):
    ids: list[int] = Field(default_factory=list)
    itens: list[ProgramacaoVendaSelecionadaItem] = Field(default_factory=list)
    invalidas: int = 0


class ProgramacaoSugestaoPayload(BaseModel):
    veiculo: str | None = Field(default=None, max_length=40)
    local_rota: str | None = Field(default=None, max_length=40)
    itens: list[ProgramacaoItemPayload] = Field(default_factory=list)


class ProgramacaoSugestaoResponse(BaseModel):
    veiculo: str = ""
    capacidade_cx: int = 0
    total_caixas: int = 0
    caixas_dentro_capacidade: int = 0
    caixas_excedentes: int = 0
    clientes_com_localizacao: int = 0
    clientes_sem_localizacao: int = 0
    distancia_estimativa_km: float = 0
    resumo: str = ""
    alertas: list[str] = Field(default_factory=list)
    itens: list[dict[str, Any]] = Field(default_factory=list)


class ProgramacaoRankingItem(BaseModel):
    id: str
    nome: str
    codigo: str | None = None
    display: str
    status: str = "ATIVO"
    total_horas_trabalhadas: float = 0
    total_km_rodado: float = 0
    total_viagens: int = 0
    dias_desde_ultima_programacao: int = 0
    score_exibicao: float = 0
    posicao_ranking: int = 0
    motivo_resumido: str = ""
    disponivel: bool = True


class ProgramacaoRankingsResponse(BaseModel):
    periodo_dias: int
    motoristas: list[ProgramacaoRankingItem] = Field(default_factory=list)
    ajudantes: list[ProgramacaoRankingItem] = Field(default_factory=list)
    resumo_motoristas: str = ""
    resumo_ajudantes: str = ""


def upper_text(value: Any) -> str:
    return str(value or "").strip().upper()


def parse_date_safe(value: Any):
    txt = str(value or "").strip()
    if not txt:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:10], fmt).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(txt.replace(" ", "T")).date()
    except Exception:
        return None


def date_ranges_overlap(inicio: Any, fim: Any, ref_inicio: Any, ref_fim: Any) -> bool:
    inicio_dt = parse_date_safe(inicio)
    fim_dt = parse_date_safe(fim) or inicio_dt
    ref_inicio_dt = parse_date_safe(ref_inicio)
    ref_fim_dt = parse_date_safe(ref_fim) or ref_inicio_dt
    if not inicio_dt or not fim_dt or not ref_inicio_dt or not ref_fim_dt:
        return False
    if fim_dt < inicio_dt:
        inicio_dt, fim_dt = fim_dt, inicio_dt
    if ref_fim_dt < ref_inicio_dt:
        ref_inicio_dt, ref_fim_dt = ref_fim_dt, ref_inicio_dt
    return inicio_dt <= ref_fim_dt and fim_dt >= ref_inicio_dt


async def folgas_ativas_programacao(db: AsyncSession, ref_inicio: str | None = None, ref_fim: str | None = None) -> dict[str, set[str]]:
    hoje = datetime.now().date().isoformat()
    ref_inicio = ref_inicio or hoje
    ref_fim = ref_fim or ref_inicio
    result = await db.execute(
        select(EscalaFolgaDB)
        .where(func.upper(func.coalesce(EscalaFolgaDB.status, "ATIVA")) == "ATIVA")
        .order_by(EscalaFolgaDB.data_inicio.asc())
    )
    out = {
        "motoristas_codigos": set(),
        "motoristas_nomes": set(),
        "ajudantes_ids": set(),
        "ajudantes_nomes": set(),
    }
    for row in result.scalars().all():
        if not date_ranges_overlap(row.data_inicio, row.data_fim, ref_inicio, ref_fim):
            continue
        tipo = upper_text(row.tipo)
        if tipo == "MOTORISTA":
            if upper_text(row.pessoa_codigo):
                out["motoristas_codigos"].add(upper_text(row.pessoa_codigo))
            if upper_text(row.pessoa_nome):
                out["motoristas_nomes"].add(upper_text(row.pessoa_nome))
        elif tipo == "AJUDANTE":
            if str(row.pessoa_id or "").strip():
                out["ajudantes_ids"].add(str(row.pessoa_id or "").strip())
            if upper_text(row.pessoa_nome):
                out["ajudantes_nomes"].add(upper_text(row.pessoa_nome))
    return out


def normalize_ascii(value: Any) -> str:
    text_value = upper_text(value)
    decomposed = unicodedata.normalize("NFKD", text_value)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_operacao_tipo(value: Any, tipo_estimativa: str = "") -> str:
    raw = normalize_ascii(value).replace("-", "_").replace(" ", "_")
    if raw in {"TRANSBORDO", "TRANSFERENCIA_CARGA", "REDISTRIBUICAO"}:
        return "TRANSBORDO"
    return "TRANSBORDO" if upper_text(tipo_estimativa) == "CX" else "VENDA"


def normalize_transbordo_modalidade(value: Any) -> str:
    raw = normalize_ascii(value).replace("-", "_").replace(" ", "_")
    if raw in {"CIF", "FORNECEDOR_ENTREGA", "FORNECEDOR_VEM", "RECEBER_FORNECEDOR", "FORNECEDOR"}:
        return "CIF"
    if raw in {"FOB", "EMPRESA_BUSCA", "NOS_BUSCAMOS", "BUSCAR_FORNECEDOR", "BUSCA_EMPRESA", "EMPRESA"}:
        return "EMPRESA_BUSCA"
    return "EMPRESA_BUSCA"


def is_transbordo_programacao(programacao: ProgramacaoDB) -> bool:
    return normalize_operacao_tipo(getattr(programacao, "operacao_tipo", ""), getattr(programacao, "tipo_estimativa", "")) == "TRANSBORDO"


def is_blank_import_text(value: Any) -> bool:
    return upper_text(value) in {"", "NAN", "NAT", "NONE", "NULL", "<NA>"}


def clean_import_pedido(value: Any) -> str:
    text_value = str(value or "").strip()
    if is_blank_import_text(text_value):
        return ""
    try:
        number = float(text_value.replace(",", "."))
        if abs(number - int(number)) < 1e-9:
            return str(int(number))
        return str(number).rstrip("0").rstrip(".")
    except Exception:
        return upper_text(text_value)


def parse_venda_obs_quantities(value: Any) -> tuple[int, float]:
    caixas = 1
    kg = 0.0
    text_value = str(value or "").lower()
    match_kg = re.search(r"(\d+[\.,]?\d*)\s*kg", text_value)
    if match_kg:
        kg = safe_float(match_kg.group(1), 0.0)
    match_caixas = re.search(r"(\d+)\s*cx", text_value)
    if match_caixas:
        caixas = safe_int(match_caixas.group(1), 1)
    return max(caixas, 1), max(kg, 0.0)


def geo_distance_km(lat1: Any, lon1: Any, lat2: Any, lon2: Any) -> float:
    try:
        a_lat = float(lat1)
        a_lon = float(lon1)
        b_lat = float(lat2)
        b_lon = float(lon2)
    except Exception:
        return 0.0
    r = 6371.0
    d_lat = radians(b_lat - a_lat)
    d_lon = radians(b_lon - a_lon)
    x = sin(d_lat / 2) ** 2 + cos(radians(a_lat)) * cos(radians(b_lat)) * sin(d_lon / 2) ** 2
    return round(2 * r * asin(sqrt(x)), 2)


def normalize_local_rota(value: Any) -> str:
    text_value = normalize_ascii(value)
    if text_value == "SERTAO":
        return "SERTAO"
    if text_value == "SERRA":
        return "SERRA"
    return text_value


def split_equipe(equipe: str | None) -> list[str]:
    return [upper_text(part) for part in re.split(r"[|,;/]+", str(equipe or "")) if upper_text(part)]


async def resolve_equipe_nomes(db: AsyncSession, equipe_raw: str | None) -> str:
    raw = str(equipe_raw or "").strip()
    if not raw:
        return ""
    result = await db.execute(select(AjudanteDB))
    nomes = {
        str(item.id): upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
        for item in result.scalars().all()
    }
    out: list[str] = []
    seen: set[str] = set()
    for part in re.split(r"[|,;/]+", raw):
        token = part.strip()
        if not token:
            continue
        nome = nomes.get(token) if token.isdigit() else ""
        nome = nome or upper_text(token)
        if nome in seen:
            continue
        seen.add(nome)
        out.append(nome)
    return " / ".join(out)


def motorista_display(nome: Any, codigo: Any) -> str:
    nome_text = upper_text(nome)
    codigo_text = upper_text(codigo)
    return f"{nome_text} ({codigo_text})" if codigo_text else nome_text


def ajudante_display(nome: Any, sobrenome: Any, fallback: Any = "") -> str:
    parts = [upper_text(nome), upper_text(sobrenome)]
    display = " ".join(part for part in parts if part)
    return display or upper_text(fallback)


async def recursos_ocupados_em_rotas_abertas(db: AsyncSession, *, exclude_codigo: str = "") -> dict[str, set[str]]:
    exclude = upper_text(exclude_codigo)
    result = await db.execute(
        select(ProgramacaoDB).where(func.trim(func.coalesce(ProgramacaoDB.codigo_programacao, ProgramacaoDB.codigo, "")) != "")
    )
    ocupados = {"motoristas_codigos": set(), "motoristas_nomes": set(), "veiculos": set(), "ajudantes": set()}
    for programacao in result.scalars().all():
        codigo = upper_text(programacao.codigo_programacao or programacao.codigo)
        if exclude and codigo == exclude:
            continue
        if not programacao_reserva_recurso(programacao):
            continue
        motorista_codigo = upper_text(programacao.motorista_codigo or programacao.codigo_motorista)
        motorista_nome = upper_text(programacao.motorista)
        veiculo = upper_text(programacao.veiculo)
        if motorista_codigo:
            ocupados["motoristas_codigos"].add(motorista_codigo)
        if motorista_nome:
            ocupados["motoristas_nomes"].add(motorista_nome)
        if veiculo:
            ocupados["veiculos"].add(veiculo)
        for ajudante in split_equipe(programacao.equipe):
            ocupados["ajudantes"].add(ajudante)
    return ocupados


def ranking_status(value: Any) -> str:
    return upper_text(value)


def ranking_is_ativa(status_raw: Any) -> bool:
    return ranking_status(status_raw) in {"ATIVA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"}


def ranking_is_cancelada(status_raw: Any) -> bool:
    return ranking_status(status_raw) in {"CANCELADA", "CANCELADO"}


def ranking_parse_data(raw: Any) -> datetime | None:
    text_value = str(raw or "").strip()
    if not text_value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d/%m/%y %H:%M:%S", "%d/%m/%y"):
        try:
            return datetime.strptime(text_value[:19], fmt) if "H" in fmt else datetime.strptime(text_value[:10], fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(text_value.replace(" ", "T"))
    except Exception:
        return None


def ranking_parse_data_hora(data_raw: Any, hora_raw: Any) -> datetime | None:
    dt_data = ranking_parse_data(data_raw)
    if dt_data is None:
        return None
    hora_text = str(hora_raw or "").strip()
    if not hora_text:
        return dt_data.replace(hour=0, minute=0, second=0, microsecond=0)
    normalized = normalize_time(hora_text)
    if not normalized:
        return dt_data.replace(hour=0, minute=0, second=0, microsecond=0)
    parts = (normalized + ":00:00").split(":")
    return dt_data.replace(
        hour=safe_int(parts[0], 0),
        minute=safe_int(parts[1], 0),
        second=safe_int(parts[2], 0),
        microsecond=0,
    )


def ranking_calc_horas_trabalhadas(data_saida: Any, hora_saida: Any, data_chegada: Any, hora_chegada: Any) -> float:
    dt_saida = ranking_parse_data_hora(data_saida, hora_saida)
    dt_chegada = ranking_parse_data_hora(data_chegada, hora_chegada)
    if not dt_saida or not dt_chegada:
        return 0.0
    diff = (dt_chegada - dt_saida).total_seconds() / 3600.0
    if diff <= 0:
        return 0.0
    return min(round(diff, 2), 72.0)


def normalizar_indicador(values: list[Any]) -> list[float]:
    numbers = [safe_float(value, 0.0) for value in values]
    if not numbers:
        return []
    minimum = min(numbers)
    maximum = max(numbers)
    if maximum == minimum:
        return [0.0 for _value in numbers]
    return [(value - minimum) / (maximum - minimum) for value in numbers]


def score_motorista(metricas: dict[str, Any], contexto: dict[str, dict[Any, float]]) -> float:
    horas_score = 1.0 - safe_float(contexto["horas_norm"].get(metricas["id"], 0.0), 0.0)
    km_score = 1.0 - safe_float(contexto["km_norm"].get(metricas["id"], 0.0), 0.0)
    viagens_score = 1.0 - safe_float(contexto["viagens_norm"].get(metricas["id"], 0.0), 0.0)
    descanso_score = safe_float(contexto["dias_norm"].get(metricas["id"], 0.0), 0.0)
    metricas["horas_score"] = horas_score
    metricas["km_score"] = km_score
    metricas["viagens_score"] = viagens_score
    metricas["descanso_score"] = descanso_score
    return round((horas_score * 0.35) + (km_score * 0.30) + (viagens_score * 0.20) + (descanso_score * 0.15), 6)


def score_ajudante(metricas: dict[str, Any], contexto: dict[str, dict[Any, float]]) -> float:
    horas_score = 1.0 - safe_float(contexto["horas_norm"].get(metricas["id"], 0.0), 0.0)
    viagens_score = 1.0 - safe_float(contexto["viagens_norm"].get(metricas["id"], 0.0), 0.0)
    descanso_score = safe_float(contexto["dias_norm"].get(metricas["id"], 0.0), 0.0)
    metricas["horas_score"] = horas_score
    metricas["viagens_score"] = viagens_score
    metricas["descanso_score"] = descanso_score
    return round((horas_score * 0.50) + (viagens_score * 0.30) + (descanso_score * 0.20), 6)


def motivo_motorista(item: dict[str, Any], periodo_dias: int) -> str:
    descanso = safe_float(item.get("descanso_score"), 0.0)
    if descanso >= max(safe_float(item.get("horas_score"), 0.0), safe_float(item.get("km_score"), 0.0), safe_float(item.get("viagens_score"), 0.0)):
        return "Maior intervalo desde a ultima programacao"
    if safe_float(item.get("horas_score"), 0.0) >= 0.60 and safe_float(item.get("viagens_score"), 0.0) >= 0.60:
        return "Menos horas e menos viagens no periodo"
    return f"Menor carga acumulada nos ultimos {safe_int(periodo_dias, 30)} dias"


def motivo_ajudante(item: dict[str, Any], periodo_dias: int) -> str:
    descanso = safe_float(item.get("descanso_score"), 0.0)
    if descanso >= max(safe_float(item.get("horas_score"), 0.0), safe_float(item.get("viagens_score"), 0.0)):
        return "Maior intervalo desde a ultima programacao"
    if safe_float(item.get("horas_score"), 0.0) >= 0.60 and safe_float(item.get("viagens_score"), 0.0) >= 0.60:
        return "Menos horas e menos viagens no periodo"
    return f"Menor carga acumulada nos ultimos {safe_int(periodo_dias, 30)} dias"


def ranking_summary(title: str, ranking: list[ProgramacaoRankingItem], limit: int = 3) -> str:
    if not ranking:
        return f"{title}: sem candidatos elegiveis."
    top = ranking[:max(safe_int(limit, 3), 1)]
    parts = [f"{item.posicao_ranking}. {item.nome} ({item.score_exibicao:.2f})" for item in top]
    motivo = str(top[0].motivo_resumido or "").strip()
    if len(motivo) > 44:
        motivo = motivo[:41].rstrip() + "..."
    return f"{title}: " + " | ".join(parts) + (f" | Motivo lider: {motivo}" if motivo else "")


def item_to_response(item: ProgramacaoItemDB) -> ProgramacaoItemResponse:
    return ProgramacaoItemResponse(
        id=item.id,
        cod_cliente=item.cod_cliente or "",
        nome_cliente=item.nome_cliente or "",
        produto_id=int(getattr(item, "produto_id", 0) or 0) or None,
        produto=item.produto or "",
        endereco=item.endereco or "",
        qnt_caixas=safe_int(item.qnt_caixas, 0),
        kg=safe_float(item.kg, 0.0),
        preco=safe_float(item.preco, 0.0),
        vendedor=item.vendedor or "",
        pedido=item.pedido or "",
        obs=item.observacao or "",
        ordem_sugerida=safe_int(item.ordem_sugerida, 0) or None,
        distancia=safe_float(item.distancia, 0.0),
        confianca_localizacao=safe_float(item.confianca_localizacao, 0.0),
        carga_raiz_programacao=getattr(item, "carga_raiz_programacao", None) or "",
        carga_origem_imediata=getattr(item, "carga_origem_imediata", None) or "",
        transferencia_origem_id=getattr(item, "transferencia_origem_id", None) or "",
    )


async def items_for_programacao(db: AsyncSession, codigo_programacao: str) -> list[ProgramacaoItemDB]:
    result = await db.execute(
        select(ProgramacaoItemDB)
        .where(func.upper(ProgramacaoItemDB.codigo_programacao) == upper_text(codigo_programacao))
        .order_by(ProgramacaoItemDB.id.asc())
    )
    return list(result.scalars().all())


async def item_totals_for_programacao(db: AsyncSession, codigo_programacao: str) -> tuple[int, float]:
    result = await db.execute(
        select(
            func.coalesce(func.sum(ProgramacaoItemDB.qnt_caixas), 0),
            func.coalesce(func.sum(ProgramacaoItemDB.kg), 0),
        ).where(func.upper(ProgramacaoItemDB.codigo_programacao) == upper_text(codigo_programacao))
    )
    row = result.one_or_none()
    if not row:
        return 0, 0.0
    return safe_int(row[0], 0), safe_float(row[1], 0.0)


def transferencia_qtd_total(row: Any) -> int:
    return max(safe_int(row.get("qtd_caixas"), 0), safe_int(row.get("qtd_convertida"), 0), 0)


def transferencia_qtd_convertida(row: Any) -> int:
    qtd_convertida = max(safe_int(row.get("qtd_convertida"), 0), 0)
    if qtd_convertida > 0:
        return qtd_convertida
    if upper_text(row.get("status")) == "CONVERTIDA":
        return transferencia_qtd_total(row)
    return 0


def carga_raiz_from_snapshot(value: Any, fallback: Any = "") -> str:
    try:
        data = json.loads(str(value or "{}"))
        if not isinstance(data, dict):
            data = {}
    except Exception:
        data = {}
    return upper_text(data.get("carga_raiz_programacao") or data.get("carga_origem_programacao") or fallback)


async def kg_transferido_convertido_destino(db: AsyncSession, codigo_programacao: str) -> float:
    codigo = upper_text(codigo_programacao)
    if not codigo:
        return 0.0
    try:
        result = await db.execute(
            text(
                """
                SELECT codigo_origem, codigo_destino, qtd_caixas, qtd_convertida, status, snapshot
                  FROM transferencias
                 WHERE UPPER(COALESCE(codigo_destino, ''))=:codigo
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        return 0.0
    rows = [dict(row) for row in result.mappings().all()]
    roots = {
        carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
        for row in rows
        if carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
    }
    root_map: dict[str, ProgramacaoDB] = {}
    if roots:
        root_result = await db.execute(select(ProgramacaoDB).where(func.upper(ProgramacaoDB.codigo_programacao).in_(roots)))
        root_map = {upper_text(item.codigo_programacao): item for item in root_result.scalars().all()}
    total = 0.0
    for row in rows:
        status_value = upper_text(row.get("status"))
        if status_value in {"CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"}:
            continue
        qtd = transferencia_qtd_convertida(row)
        if qtd <= 0:
            continue
        origem = upper_text(row.get("codigo_origem"))
        root = root_map.get(carga_raiz_from_snapshot(row.get("snapshot"), origem))
        kg_base = (
            safe_float(getattr(root, "nf_kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "nf_kg", 0), 0.0)
            or safe_float(getattr(root, "kg_nf", 0), 0.0)
        )
        caixas_base = (
            safe_int(getattr(root, "nf_caixas", 0), 0)
            or safe_int(getattr(root, "total_caixas", 0), 0)
            or safe_int(getattr(root, "caixas_carregadas", 0), 0)
        )
        if kg_base > 0 and caixas_base > 0:
            total += qtd * (kg_base / caixas_base)
    return round(total, 2)


async def caixas_transferidas_destino(db: AsyncSession, codigo_programacao: str) -> int:
    codigo = upper_text(codigo_programacao)
    if not codigo:
        return 0
    try:
        result = await db.execute(
            text(
                """
                SELECT qtd_caixas, qtd_convertida, status
                  FROM transferencias
                 WHERE UPPER(COALESCE(codigo_destino, ''))=:codigo
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        return 0
    total = 0
    for row in result.mappings().all():
        status_value = upper_text(row.get("status"))
        if status_value in {"CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"}:
            continue
        total += transferencia_qtd_total(row)
    return total


async def serialize_programacao(db: AsyncSession, programacao: ProgramacaoDB, *, include_items: bool) -> ProgramacaoResponse:
    codigo = upper_text(programacao.codigo_programacao or programacao.codigo)
    itens = await items_for_programacao(db, codigo) if include_items and codigo else []
    total_caixas = safe_int(programacao.total_caixas, 0)
    quilos = safe_float(programacao.quilos, 0.0)
    if codigo and (total_caixas <= 0 or quilos <= 0):
        item_caixas, item_kg = await item_totals_for_programacao(db, codigo)
        if total_caixas <= 0:
            total_caixas = item_caixas
        if quilos <= 0:
            quilos = item_kg
    if codigo and quilos <= 0:
        quilos = await kg_transferido_convertido_destino(db, codigo)
    if codigo and total_caixas <= 0:
        total_caixas = await caixas_transferidas_destino(db, codigo)
    return ProgramacaoResponse(
        id=programacao.id,
        codigo_programacao=codigo or f"ID {programacao.id}",
        data_criacao=programacao.data_criacao,
        motorista=programacao.motorista or "",
        motorista_codigo=programacao.motorista_codigo or programacao.codigo_motorista or "",
        veiculo=programacao.veiculo or "",
        equipe=programacao.equipe or "",
        ajudantes=split_equipe(programacao.equipe),
        kg_estimado=safe_float(programacao.kg_estimado, 0.0),
        tipo_estimativa=upper_text(programacao.tipo_estimativa or "KG"),
        caixas_estimado=safe_int(programacao.caixas_estimado, 0),
        operacao_tipo=normalize_operacao_tipo(programacao.operacao_tipo, programacao.tipo_estimativa),
        transbordo_modalidade=programacao.transbordo_modalidade or "",
        transbordo_observacao=programacao.transbordo_observacao or "",
        status=programacao.status or "",
        status_operacional=programacao.status_operacional or "",
        prestacao_status=programacao.prestacao_status or "",
        local_rota=programacao.local_rota or programacao.tipo_rota or "",
        local_carregamento=programacao.local_carregamento or programacao.granja_carregada or programacao.local_carregado or programacao.local_carreg or "",
        adiantamento=safe_float(programacao.adiantamento, 0.0),
        adiantamento_origem=programacao.adiantamento_origem or "",
        total_caixas=total_caixas,
        quilos=quilos,
        itens=[item_to_response(item) for item in itens],
    )


async def get_programacao_by_codigo(db: AsyncSession, codigo_programacao: str) -> ProgramacaoDB | None:
    result = await db.execute(
        select(ProgramacaoDB)
        .where(func.upper(ProgramacaoDB.codigo_programacao) == upper_text(codigo_programacao))
        .order_by(ProgramacaoDB.id.desc())
        .limit(1)
    )
    return result.scalar_one_or_none()


async def cliente_endereco_map(db: AsyncSession, codigos: set[str]) -> dict[str, str]:
    if not codigos:
        return {}
    result = await db.execute(
        select(ClienteDB).where(func.upper(func.coalesce(ClienteDB.cod_cliente, "")).in_(codigos))
    )
    return {upper_text(item.cod_cliente): upper_text(item.endereco) for item in result.scalars().all()}


async def produto_id_for_item(db: AsyncSession, produto_id: int | None, produto_nome: str | None) -> int | None:
    if produto_id:
        found = await db.get(ProdutoDB, int(produto_id))
        if found:
            return int(found.id)
    nome_norm = upper_text(produto_nome)
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
        descricao="Cadastro automatico criado pela programacao.",
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


def venda_importada_to_programacao_item(venda: VendaImportadaDB, endereco_map: dict[str, str]) -> ProgramacaoVendaSelecionadaItem | None:
    cod_cliente = upper_text(venda.cliente)
    pedido = clean_import_pedido(venda.pedido)
    nome_cliente = upper_text(venda.nome_cliente)
    produto = upper_text(venda.produto)
    if is_blank_import_text(cod_cliente) or is_blank_import_text(pedido) or is_blank_import_text(nome_cliente) or is_blank_import_text(produto):
        return None

    qnt = safe_float(venda.qnt, 0.0)
    vr_total = safe_float(venda.vr_total, 0.0)
    valor_unitario = safe_float(venda.valor_unitario, 0.0)
    preco = (vr_total / qnt) if qnt > 0 else valor_unitario
    caixas_obs, kg = parse_venda_obs_quantities(venda.observacao)
    caixas = caixas_obs or safe_int(getattr(venda, "qnt_caixas", 0), 0)
    cidade = upper_text(venda.cidade)
    endereco = cidade or endereco_map.get(cod_cliente, "")
    return ProgramacaoVendaSelecionadaItem(
        venda_id=safe_int(venda.id, 0),
        cod_cliente=cod_cliente,
        nome_cliente=nome_cliente,
        produto_id=int(getattr(venda, "produto_id", 0) or 0) or None,
        produto=produto,
        endereco=endereco,
        qnt_caixas=caixas,
        kg=kg,
        preco=safe_float(preco, 0.0),
        vendedor=upper_text(venda.vendedor),
        pedido=upper_text(pedido),
        obs=upper_text(venda.observacao),
    )


async def selected_vendas_importadas_items(db: AsyncSession) -> ProgramacaoVendasSelecionadasResponse:
    result = await db.execute(
        select(VendaImportadaDB)
        .where(
            func.coalesce(VendaImportadaDB.selecionada, 0) == 1,
            func.coalesce(VendaImportadaDB.usada, 0) == 0,
            func.trim(func.coalesce(VendaImportadaDB.codigo_programacao, "")) == "",
        )
        .order_by(VendaImportadaDB.id.asc())
    )
    vendas = list(result.scalars().all())
    endereco_map = await cliente_endereco_map(db, {upper_text(venda.cliente) for venda in vendas if upper_text(venda.cliente)})
    response = ProgramacaoVendasSelecionadasResponse()
    seen: set[tuple[str, str, str]] = set()
    for venda in vendas:
        item = venda_importada_to_programacao_item(venda, endereco_map)
        if not item:
            response.invalidas += 1
            continue
        key = (upper_text(item.pedido), upper_text(item.cod_cliente), upper_text(item.produto))
        if key in seen:
            continue
        seen.add(key)
        response.ids.append(item.venda_id)
        response.itens.append(item)
    return response


async def capacidade_veiculo(db: AsyncSession, placa: str | None) -> int:
    placa_norm = upper_text(placa)
    if not placa_norm:
        return 0
    result = await db.execute(select(VeiculoDB).where(func.upper(func.trim(VeiculoDB.placa)) == placa_norm).limit(1))
    veiculo = result.scalar_one_or_none()
    return safe_int(veiculo.capacidade_cx if veiculo else 0, 0)


async def require_veiculo_com_capacidade(db: AsyncSession, placa: str | None) -> tuple[str, int]:
    placa_norm = upper_text(placa)
    if not placa_norm:
        raise HTTPException(status_code=422, detail="Selecione um veiculo antes de sugerir rota.")
    result = await db.execute(select(VeiculoDB).where(func.upper(func.trim(VeiculoDB.placa)) == placa_norm).limit(1))
    veiculo = result.scalar_one_or_none()
    if not veiculo:
        raise HTTPException(status_code=422, detail=f"Veiculo nao encontrado no cadastro: {placa_norm}")
    capacidade = safe_int(veiculo.capacidade_cx, 0)
    if capacidade <= 0:
        raise HTTPException(status_code=422, detail=f"Cadastre a capacidade em caixas do veiculo {placa_norm} antes de sugerir rota.")
    return placa_norm, capacidade


async def localizacoes_clientes(db: AsyncSession, codigos: set[str]) -> dict[str, dict[str, Any]]:
    codigos_norm = [upper_text(codigo) for codigo in codigos if upper_text(codigo)]
    if not codigos_norm:
        return {}
    out: dict[str, dict[str, Any]] = {}
    params = {f"c{i}": codigo for i, codigo in enumerate(codigos_norm)}
    in_clause = ", ".join(f":c{i}" for i in range(len(codigos_norm)))
    try:
        result = await db.execute(
            text(
                f"""
                SELECT cla.cod_cliente,
                       cla.latitude,
                       cla.longitude,
                       cla.endereco,
                       cla.cidade,
                       cla.bairro,
                       cla.origem,
                       cla.registrado_em,
                       (
                           SELECT COUNT(*)
                             FROM cliente_localizacao_amostras cla2
                            WHERE UPPER(TRIM(cla2.cod_cliente)) = UPPER(TRIM(cla.cod_cliente))
                              AND cla2.latitude IS NOT NULL
                              AND cla2.longitude IS NOT NULL
                       ) AS amostras
                  FROM cliente_localizacao_amostras cla
                 WHERE UPPER(TRIM(cla.cod_cliente)) IN ({in_clause})
                   AND cla.latitude IS NOT NULL
                   AND cla.longitude IS NOT NULL
                 ORDER BY cla.cod_cliente, cla.registrado_em DESC, cla.id DESC
                """
            ),
            params,
        )
        for row in result.mappings().all():
            cod = upper_text(row.get("cod_cliente"))
            if cod in out:
                continue
            out[cod] = {
                "lat": safe_float(row.get("latitude"), 0.0),
                "lon": safe_float(row.get("longitude"), 0.0),
                "endereco": str(row.get("endereco") or ""),
                "cidade": str(row.get("cidade") or ""),
                "bairro": str(row.get("bairro") or ""),
                "origem": str(row.get("origem") or "APP"),
                "amostras": safe_int(row.get("amostras"), 1),
                "registrado_em": str(row.get("registrado_em") or ""),
            }
    except Exception:
        pass

    missing = [codigo for codigo in codigos_norm if codigo not in out]
    if missing:
        params_ctrl = {f"p{i}": codigo for i, codigo in enumerate(missing)}
        in_clause_ctrl = ", ".join(f":p{i}" for i in range(len(missing)))
        try:
            result = await db.execute(
                text(
                    f"""
                    SELECT pc.cod_cliente,
                           COALESCE(pc.lat_entrega, pc.lat_evento) AS latitude,
                           COALESCE(pc.lon_entrega, pc.lon_evento) AS longitude,
                           pc.endereco_evento AS endereco,
                           pc.cidade_evento AS cidade,
                           pc.bairro_evento AS bairro,
                           pc.updated_at AS registrado_em,
                           (
                               SELECT COUNT(*)
                                 FROM programacao_itens_controle pc2
                                WHERE UPPER(TRIM(pc2.cod_cliente)) = UPPER(TRIM(pc.cod_cliente))
                                  AND COALESCE(pc2.lat_entrega, pc2.lat_evento) IS NOT NULL
                                  AND COALESCE(pc2.lon_entrega, pc2.lon_evento) IS NOT NULL
                           ) AS amostras
                      FROM programacao_itens_controle pc
                     WHERE UPPER(TRIM(pc.cod_cliente)) IN ({in_clause_ctrl})
                       AND COALESCE(pc.lat_entrega, pc.lat_evento) IS NOT NULL
                       AND COALESCE(pc.lon_entrega, pc.lon_evento) IS NOT NULL
                     ORDER BY pc.cod_cliente, pc.updated_at DESC, pc.id DESC
                    """
                ),
                params_ctrl,
            )
            for row in result.mappings().all():
                cod = upper_text(row.get("cod_cliente"))
                if cod in out:
                    continue
                out[cod] = {
                    "lat": safe_float(row.get("latitude"), 0.0),
                    "lon": safe_float(row.get("longitude"), 0.0),
                    "endereco": str(row.get("endereco") or ""),
                    "cidade": str(row.get("cidade") or ""),
                    "bairro": str(row.get("bairro") or ""),
                    "origem": "ENTREGA",
                    "amostras": safe_int(row.get("amostras"), 1),
                    "registrado_em": str(row.get("registrado_em") or ""),
                }
        except Exception:
            pass

    missing = [codigo for codigo in codigos_norm if codigo not in out]
    if missing:
        params2 = {f"m{i}": codigo for i, codigo in enumerate(missing)}
        in_clause2 = ", ".join(f":m{i}" for i in range(len(missing)))
        try:
            result = await db.execute(
                text(
                    f"""
                    SELECT cod_cliente, latitude, longitude, endereco, cidade, bairro
                    FROM clientes
                    WHERE UPPER(TRIM(cod_cliente)) IN ({in_clause2})
                    """
                ),
                params2,
            )
            for row in result.mappings().all():
                lat = safe_float(row.get("latitude"), 0.0)
                lon = safe_float(row.get("longitude"), 0.0)
                if lat == 0 and lon == 0:
                    continue
                cod = upper_text(row.get("cod_cliente"))
                out[cod] = {
                    "lat": lat,
                    "lon": lon,
                    "endereco": str(row.get("endereco") or ""),
                    "cidade": str(row.get("cidade") or ""),
                    "bairro": str(row.get("bairro") or ""),
                    "origem": "CADASTRO",
                    "amostras": 0,
                    "registrado_em": "",
                }
        except Exception:
            pass
    return out


async def historico_pares_clientes(db: AsyncSession, codigos: set[str], limit_programacoes: int = 80) -> dict[tuple[str, str], int]:
    codigos_norm = {upper_text(codigo) for codigo in codigos if upper_text(codigo)}
    if len(codigos_norm) < 2:
        return {}
    try:
        result = await db.execute(
            select(ProgramacaoItemDB)
            .where(ProgramacaoItemDB.cod_cliente.in_(codigos_norm))
            .order_by(ProgramacaoItemDB.codigo_programacao.desc(), ProgramacaoItemDB.id.asc())
            .limit(max(safe_int(limit_programacoes, 80), 10) * 80)
        )
    except Exception:
        return {}
    por_programacao: dict[str, list[str]] = {}
    for item in result.scalars().all():
        codigo = upper_text(item.codigo_programacao)
        cliente = upper_text(item.cod_cliente)
        if codigo and cliente:
            por_programacao.setdefault(codigo, []).append(cliente)

    pares: dict[tuple[str, str], int] = {}
    for clientes in list(por_programacao.values())[:limit_programacoes]:
        sequencia = [cliente for index, cliente in enumerate(clientes) if cliente and cliente not in clientes[:index]]
        for atual, proximo in zip(sequencia, sequencia[1:]):
            if atual == proximo:
                continue
            pares[(atual, proximo)] = pares.get((atual, proximo), 0) + 1
            pares[(proximo, atual)] = pares.get((proximo, atual), 0) + 1
    return pares


def ordenar_por_vizinho_mais_proximo(items: list[dict[str, Any]], historico: dict[tuple[str, str], int] | None = None) -> list[dict[str, Any]]:
    historico = historico or {}
    com_geo = [item for item in items if item.get("lat") not in (None, "") and item.get("lon") not in (None, "")]
    sem_geo = [item for item in items if item not in com_geo]
    if len(com_geo) <= 1:
        return com_geo + sorted(sem_geo, key=lambda it: (upper_text(it.get("cidade")), upper_text(it.get("endereco")), upper_text(it.get("nome_cliente"))))
    start = min(com_geo, key=lambda it: (upper_text(it.get("cidade")), upper_text(it.get("bairro")), upper_text(it.get("nome_cliente"))))
    ordered = [start]
    remaining = [item for item in com_geo if item is not start]
    while remaining:
        last = ordered[-1]
        last_cod = upper_text(last.get("cod_cliente"))
        nxt = min(
            remaining,
            key=lambda it: (
                geo_distance_km(last.get("lat"), last.get("lon"), it.get("lat"), it.get("lon"))
                - min(historico.get((last_cod, upper_text(it.get("cod_cliente"))), 0) * 2.0, 12.0),
                upper_text(it.get("cidade")),
                upper_text(it.get("bairro")),
                upper_text(it.get("nome_cliente")),
            ),
        )
        ordered.append(nxt)
        remaining.remove(nxt)
    ordered.extend(sorted(sem_geo, key=lambda it: (upper_text(it.get("cidade")), upper_text(it.get("endereco")), upper_text(it.get("nome_cliente")))))
    return ordered


async def mark_vendas_importadas_used(db: AsyncSession, ids: list[int], codigo_programacao: str) -> int:
    venda_ids = sorted({safe_int(item, 0) for item in ids if safe_int(item, 0) > 0})
    if not venda_ids:
        return 0
    result = await db.execute(select(VendaImportadaDB).where(VendaImportadaDB.id.in_(venda_ids)))
    vendas = list(result.scalars().all())
    by_id = {safe_int(venda.id, 0): venda for venda in vendas}
    unavailable = []
    codigo = upper_text(codigo_programacao)
    for venda_id in venda_ids:
        venda = by_id.get(venda_id)
        if not venda:
            unavailable.append(venda_id)
            continue
        linked_same = safe_int(venda.usada, 0) == 1 and upper_text(venda.codigo_programacao) == codigo
        free = safe_int(venda.usada, 0) == 0 and not upper_text(venda.codigo_programacao)
        if not (linked_same or free):
            unavailable.append(venda_id)
    if unavailable:
        raise HTTPException(status_code=409, detail=f"Venda(s) importada(s) indisponivel(is): {', '.join(map(str, unavailable))}")

    usada_em = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for venda in vendas:
        venda.usada = 1
        venda.usada_em = usada_em
        venda.codigo_programacao = codigo
        venda.selecionada = 0
    return len(vendas)


async def programacoes_para_score(db: AsyncSession, periodo_dias: int) -> list[dict[str, Any]]:
    result = await db.execute(select(ProgramacaoDB).order_by(ProgramacaoDB.id.desc()).limit(5000))
    rows = list(result.scalars().all())
    cutoff = datetime.now() - timedelta(days=max(safe_int(periodo_dias, 30), 1))
    out = []
    for programacao in rows:
        data_ref = programacao.data_saida or programacao.data_criacao or programacao.data or ""
        dt_ref = ranking_parse_data(data_ref)
        if dt_ref is not None and dt_ref < cutoff:
            continue
        status = ranking_status(programacao.status_operacional or programacao.status)
        if not status and safe_int(programacao.finalizada_no_app, 0) == 1:
            status = "FINALIZADA"
        if ranking_is_cancelada(status):
            continue
        out.append(
            {
                "codigo_programacao": upper_text(programacao.codigo_programacao),
                "data_ref": dt_ref,
                "motorista": upper_text(programacao.motorista),
                "motorista_codigo": upper_text(programacao.motorista_codigo or programacao.codigo_motorista),
                "equipe": str(programacao.equipe or ""),
                "status": status,
                "km_rodado": safe_float(programacao.km_rodado, 0.0),
                "data_saida": str(programacao.data_saida or ""),
                "hora_saida": str(programacao.hora_saida or ""),
                "data_chegada": str(programacao.data_chegada or ""),
                "hora_chegada": str(programacao.hora_chegada or ""),
            }
        )
    return out


async def candidatos_motoristas(db: AsyncSession) -> list[dict[str, Any]]:
    folgas = await folgas_ativas_programacao(db)
    result = await db.execute(select(MotoristaDB).order_by(MotoristaDB.nome.asc(), MotoristaDB.codigo.asc()))
    candidatos = []
    for motorista in result.scalars().all():
        nome = upper_text(motorista.nome)
        codigo = upper_text(motorista.codigo)
        status_item = ranking_status(motorista.status) or "ATIVO"
        em_folga = (codigo and codigo in folgas["motoristas_codigos"]) or (nome and nome in folgas["motoristas_nomes"])
        candidatos.append(
            {
                "id": str(safe_int(motorista.id, 0)),
                "nome": nome,
                "codigo": codigo,
                "display": motorista_display(nome, codigo),
                "status": "FOLGA" if em_folga else status_item,
                "elegivel": bool(nome) and status_item == "ATIVO" and not em_folga,
            }
        )
    return candidatos


async def candidatos_ajudantes(db: AsyncSession) -> list[dict[str, Any]]:
    folgas = await folgas_ativas_programacao(db)
    result = await db.execute(select(AjudanteDB).order_by(AjudanteDB.nome.asc(), AjudanteDB.sobrenome.asc()))
    candidatos = []
    for ajudante in result.scalars().all():
        ajudante_id = str(safe_int(ajudante.id, 0))
        display = ajudante_display(ajudante.nome, ajudante.sobrenome, ajudante_id)
        status_item = ranking_status(ajudante.status) or "ATIVO"
        em_folga = (ajudante_id and ajudante_id in folgas["ajudantes_ids"]) or (
            display and upper_text(display) in folgas["ajudantes_nomes"]
        )
        candidatos.append(
            {
                "id": ajudante_id,
                "nome": display,
                "display": display,
                "codigo": None,
                "status": "FOLGA" if em_folga else status_item,
                "elegivel": bool(display) and status_item == "ATIVO" and not em_folga,
            }
        )
    return candidatos


def resolver_motorista_score_key(programacao: dict[str, Any], por_codigo: dict[str, str], por_nome: dict[str, str]) -> str:
    codigo = upper_text(programacao.get("motorista_codigo"))
    nome = upper_text(programacao.get("motorista"))
    if codigo and codigo in por_codigo:
        return por_codigo[codigo]
    if nome and nome in por_nome:
        return por_nome[nome]
    return ""


def resolver_ajudantes_score_keys(equipe_raw: Any, por_id: dict[str, str], por_nome: dict[str, str]) -> list[str]:
    out = []
    seen = set()
    for part in re.split(r"[|,;/]+", str(equipe_raw or "")):
        token = part.strip()
        if not token:
            continue
        if token in por_id and token not in seen:
            seen.add(token)
            out.append(token)
            continue
        nome = upper_text(token)
        ajudante_id = por_nome.get(nome, "")
        if ajudante_id and ajudante_id not in seen:
            seen.add(ajudante_id)
            out.append(ajudante_id)
    return out


async def metricas_motoristas(db: AsyncSession, periodo_dias: int) -> dict[str, dict[str, Any]]:
    candidatos = await candidatos_motoristas(db)
    programacoes = await programacoes_para_score(db, periodo_dias)
    now = datetime.now()
    por_codigo = {item["codigo"]: item["id"] for item in candidatos if item.get("codigo")}
    por_nome = {item["nome"]: item["id"] for item in candidatos if item.get("nome")}
    conflitos = set()
    for programacao in programacoes:
        if ranking_is_ativa(programacao.get("status")):
            key = resolver_motorista_score_key(programacao, por_codigo, por_nome)
            if key:
                conflitos.add(key)

    metricas = {}
    for candidato in candidatos:
        if not candidato.get("elegivel"):
            continue
        candidato_id = str(candidato["id"])
        metricas[candidato_id] = {
            "id": candidato_id,
            "nome": candidato["nome"],
            "codigo": candidato.get("codigo", ""),
            "display": candidato.get("display") or candidato["nome"],
            "status": candidato.get("status", "ATIVO"),
            "total_horas_trabalhadas": 0.0,
            "total_km_rodado": 0.0,
            "total_viagens": 0,
            "dias_desde_ultima_programacao": max(safe_int(periodo_dias, 30), 1),
            "score_final": 0.0,
            "motivo_resumido": "",
            "disponivel": candidato_id not in conflitos,
            "sem_base_historica": True,
            "_last_dt": None,
        }

    for programacao in programacoes:
        candidato_id = resolver_motorista_score_key(programacao, por_codigo, por_nome)
        if not candidato_id or candidato_id not in metricas:
            continue
        horas = ranking_calc_horas_trabalhadas(
            programacao.get("data_saida"),
            programacao.get("hora_saida"),
            programacao.get("data_chegada"),
            programacao.get("hora_chegada"),
        )
        item = metricas[candidato_id]
        item["total_horas_trabalhadas"] += horas
        item["total_km_rodado"] += safe_float(programacao.get("km_rodado"), 0.0)
        item["total_viagens"] += 1
        dt_ref = programacao.get("data_ref")
        if isinstance(dt_ref, datetime) and (item["_last_dt"] is None or dt_ref > item["_last_dt"]):
            item["_last_dt"] = dt_ref
            item["dias_desde_ultima_programacao"] = max((now - dt_ref).days, 0)

    for item in metricas.values():
        item["total_horas_trabalhadas"] = round(safe_float(item.get("total_horas_trabalhadas"), 0.0), 2)
        item["total_km_rodado"] = round(safe_float(item.get("total_km_rodado"), 0.0), 2)
        item["sem_base_historica"] = (
            item["total_horas_trabalhadas"] <= 0
            and item["total_km_rodado"] <= 0
            and safe_int(item.get("total_viagens"), 0) <= 0
            and item["_last_dt"] is None
        )
    return metricas


async def metricas_ajudantes(db: AsyncSession, periodo_dias: int) -> dict[str, dict[str, Any]]:
    candidatos = await candidatos_ajudantes(db)
    programacoes = await programacoes_para_score(db, periodo_dias)
    now = datetime.now()
    por_id = {str(item["id"]): str(item["id"]) for item in candidatos if item.get("id")}
    por_nome = {upper_text(item.get("display")): str(item["id"]) for item in candidatos if item.get("display")}
    conflitos = set()
    for programacao in programacoes:
        if ranking_is_ativa(programacao.get("status")):
            conflitos.update(resolver_ajudantes_score_keys(programacao.get("equipe"), por_id, por_nome))

    metricas = {}
    for candidato in candidatos:
        if not candidato.get("elegivel"):
            continue
        candidato_id = str(candidato["id"])
        metricas[candidato_id] = {
            "id": candidato_id,
            "nome": candidato.get("display") or candidato.get("nome") or "",
            "codigo": None,
            "display": candidato.get("display") or candidato.get("nome") or "",
            "status": candidato.get("status", "ATIVO"),
            "total_horas_trabalhadas": 0.0,
            "total_km_rodado": 0.0,
            "total_viagens": 0,
            "dias_desde_ultima_programacao": max(safe_int(periodo_dias, 30), 1),
            "score_final": 0.0,
            "motivo_resumido": "",
            "disponivel": candidato_id not in conflitos,
            "sem_base_historica": True,
            "_last_dt": None,
        }

    for programacao in programacoes:
        ajudante_ids = resolver_ajudantes_score_keys(programacao.get("equipe"), por_id, por_nome)
        if not ajudante_ids:
            continue
        horas = ranking_calc_horas_trabalhadas(
            programacao.get("data_saida"),
            programacao.get("hora_saida"),
            programacao.get("data_chegada"),
            programacao.get("hora_chegada"),
        )
        dt_ref = programacao.get("data_ref")
        for candidato_id in ajudante_ids:
            if candidato_id not in metricas:
                continue
            item = metricas[candidato_id]
            item["total_horas_trabalhadas"] += horas
            item["total_viagens"] += 1
            if isinstance(dt_ref, datetime) and (item["_last_dt"] is None or dt_ref > item["_last_dt"]):
                item["_last_dt"] = dt_ref
                item["dias_desde_ultima_programacao"] = max((now - dt_ref).days, 0)

    for item in metricas.values():
        item["total_horas_trabalhadas"] = round(safe_float(item.get("total_horas_trabalhadas"), 0.0), 2)
        item["sem_base_historica"] = (
            item["total_horas_trabalhadas"] <= 0
            and safe_int(item.get("total_viagens"), 0) <= 0
            and item["_last_dt"] is None
        )
    return metricas


def ranking_item_from_metricas(item: dict[str, Any], posicao: int) -> ProgramacaoRankingItem:
    return ProgramacaoRankingItem(
        id=str(item.get("id") or ""),
        nome=upper_text(item.get("nome") or item.get("display")),
        codigo=upper_text(item.get("codigo")) or None,
        display=upper_text(item.get("display") or item.get("nome")),
        status=upper_text(item.get("status") or "ATIVO"),
        total_horas_trabalhadas=round(safe_float(item.get("total_horas_trabalhadas"), 0.0), 2),
        total_km_rodado=round(safe_float(item.get("total_km_rodado"), 0.0), 2),
        total_viagens=safe_int(item.get("total_viagens"), 0),
        dias_desde_ultima_programacao=safe_int(item.get("dias_desde_ultima_programacao"), 0),
        score_exibicao=round(safe_float(item.get("score_final"), 0.0) * 100.0, 2),
        posicao_ranking=posicao,
        motivo_resumido=str(item.get("motivo_resumido") or ""),
        disponivel=bool(item.get("disponivel", True)),
    )


async def ranquear_motoristas(db: AsyncSession, periodo_dias: int) -> list[ProgramacaoRankingItem]:
    metricas = await metricas_motoristas(db, periodo_dias)
    elegiveis = [dict(value) for value in metricas.values() if value.get("disponivel")]
    if not elegiveis:
        return []
    if all(value.get("sem_base_historica") for value in elegiveis):
        for item in elegiveis:
            item["score_final"] = 0.0
            item["motivo_resumido"] = "Sem base historica suficiente"
    else:
        ids = [item["id"] for item in elegiveis]
        contexto = {
            "horas_norm": dict(zip(ids, normalizar_indicador([item["total_horas_trabalhadas"] for item in elegiveis]))),
            "km_norm": dict(zip(ids, normalizar_indicador([item["total_km_rodado"] for item in elegiveis]))),
            "viagens_norm": dict(zip(ids, normalizar_indicador([item["total_viagens"] for item in elegiveis]))),
            "dias_norm": dict(zip(ids, normalizar_indicador([item["dias_desde_ultima_programacao"] for item in elegiveis]))),
        }
        for item in elegiveis:
            item["score_final"] = score_motorista(item, contexto)
            item["motivo_resumido"] = motivo_motorista(item, periodo_dias)
    elegiveis.sort(
        key=lambda item: (
            -round(safe_float(item.get("score_final"), 0.0), 8),
            safe_float(item.get("total_horas_trabalhadas"), 0.0),
            safe_int(item.get("total_viagens"), 0),
            safe_float(item.get("total_km_rodado"), 0.0),
            -safe_int(item.get("dias_desde_ultima_programacao"), 0),
            upper_text(item.get("nome")),
        )
    )
    return [ranking_item_from_metricas(item, posicao) for posicao, item in enumerate(elegiveis, start=1)]


async def ranquear_ajudantes(db: AsyncSession, periodo_dias: int) -> list[ProgramacaoRankingItem]:
    metricas = await metricas_ajudantes(db, periodo_dias)
    elegiveis = [dict(value) for value in metricas.values() if value.get("disponivel")]
    if not elegiveis:
        return []
    if all(value.get("sem_base_historica") for value in elegiveis):
        for item in elegiveis:
            item["score_final"] = 0.0
            item["motivo_resumido"] = "Sem base historica suficiente"
    else:
        ids = [item["id"] for item in elegiveis]
        contexto = {
            "horas_norm": dict(zip(ids, normalizar_indicador([item["total_horas_trabalhadas"] for item in elegiveis]))),
            "viagens_norm": dict(zip(ids, normalizar_indicador([item["total_viagens"] for item in elegiveis]))),
            "dias_norm": dict(zip(ids, normalizar_indicador([item["dias_desde_ultima_programacao"] for item in elegiveis]))),
        }
        for item in elegiveis:
            item["score_final"] = score_ajudante(item, contexto)
            item["motivo_resumido"] = motivo_ajudante(item, periodo_dias)
    elegiveis.sort(
        key=lambda item: (
            -round(safe_float(item.get("score_final"), 0.0), 8),
            safe_float(item.get("total_horas_trabalhadas"), 0.0),
            safe_int(item.get("total_viagens"), 0),
            -safe_int(item.get("dias_desde_ultima_programacao"), 0),
            upper_text(item.get("nome")),
        )
    )
    return [ranking_item_from_metricas(item, posicao) for posicao, item in enumerate(elegiveis, start=1)]


async def next_programacao_codigo(db: AsyncSession) -> str:
    prefix = f"PG{datetime.now().strftime('%Y')}"
    result = await db.execute(
        select(ProgramacaoDB.codigo_programacao).where(ProgramacaoDB.codigo_programacao.like(f"{prefix}%"))
    )
    max_suffix = 0
    for codigo in result.scalars().all():
        codigo_up = upper_text(codigo)
        if not codigo_up.startswith(prefix):
            continue
        digits = "".join(ch for ch in codigo_up[len(prefix):] if ch.isdigit())
        max_suffix = max(max_suffix, safe_int(digits, 0) if digits else 0)
    return f"{prefix}{max_suffix + 1:02d}"


def assert_programacao_editable(programacao: ProgramacaoDB) -> None:
    prestacao = upper_text(programacao.prestacao_status)
    if prestacao == "FECHADA":
        raise HTTPException(status_code=409, detail="Programacao com prestacao FECHADA nao pode ser alterada.")
    status_ref = upper_text(programacao.status_operacional or programacao.status)
    if status_ref not in EDITABLE_STATUSES:
        raise HTTPException(status_code=409, detail=f"Programacao com status {status_ref or '-'} nao pode ser alterada.")


def assert_programacao_deletable(programacao: ProgramacaoDB) -> None:
    prestacao = upper_text(programacao.prestacao_status)
    if prestacao == "FECHADA":
        raise HTTPException(status_code=409, detail="Programacao com prestacao FECHADA nao pode ser excluida.")
    status_ref = upper_text(programacao.status_operacional or programacao.status)
    if status_ref not in DELETABLE_STATUSES:
        raise HTTPException(status_code=409, detail="Somente programacoes ATIVAS podem ser excluidas.")


def pdf_text(value: Any) -> str:
    return str(value or "").strip()


def wrap_pdf_text(value: Any, max_chars: int) -> list[str]:
    words = pdf_text(value).split()
    if not words:
        return [""]
    lines: list[str] = []
    line = ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if len(candidate) <= max_chars:
            line = candidate
            continue
        if line:
            lines.append(line)
        line = word[:max_chars]
    if line:
        lines.append(line)
    return lines


def draw_pdf_label_value(pdf: Any, x: int, y: int, label: str, value: Any, *, max_chars: int = 92, line_height: int = 12) -> int:
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(x, y, label)
    pdf.setFont("Helvetica", 9)
    text_x = x + 92
    lines = wrap_pdf_text(value, max_chars)
    for index, line in enumerate(lines):
        pdf.drawString(text_x, y - (index * line_height), line[:max_chars])
    return y - max(line_height, len(lines) * line_height)


def pdf_money_br(value: Any) -> str:
    return f"R$ {safe_float(value, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pdf_number_br(value: Any, places: int = 2) -> str:
    return f"{safe_float(value, 0.0):,.{places}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_local_rota_pdf(value: Any) -> str:
    local = normalize_local_rota(value)
    if local == "SERTAO":
        return "SERTAO"
    if local == "SERRA":
        return "SERRA"
    return upper_text(value)


def local_carregamento_pdf(programacao: ProgramacaoDB) -> str:
    return upper_text(
        programacao.local_carregamento
        or programacao.granja_carregada
        or programacao.local_carregado
        or programacao.local_carreg
    )


def data_saida_pdf(programacao: ProgramacaoDB) -> str:
    data = str(programacao.saida_data or programacao.data_saida or "").strip()
    hora = str(programacao.saida_hora or programacao.hora_saida or "").strip()
    return " ".join(part for part in (data, hora) if part) or "-"


def data_chegada_pdf(programacao: ProgramacaoDB) -> str:
    data = str(programacao.data_chegada or "").strip()
    hora = str(programacao.hora_chegada or "").strip()
    return " ".join(part for part in (data, hora) if part) or "-"


def programacao_estimativa_pdf(programacao: ProgramacaoDB) -> str:
    tipo = upper_text(programacao.tipo_estimativa or "KG")
    if tipo == "CX":
        return f"Estimado (EMPRESA BUSCA): {safe_int(programacao.caixas_estimado, 0)} CX"
    return f"Estimado (CIF): {pdf_number_br(programacao.kg_estimado, 2)} KG"


def adiantamento_valor_pdf(programacao: ProgramacaoDB) -> float:
    return safe_float(programacao.adiantamento, 0.0) or safe_float(programacao.adiantamento_rota, 0.0)


def draw_pdf_footer(pdf: Any, width: float, page_number: int) -> None:
    pdf.setFont("Helvetica-Oblique", 8)
    pdf.drawCentredString(width / 2, 26, '"Tudo posso naquele que me fortalece." (Filipenses 4:13)')
    pdf.setFont("Helvetica", 7)
    pdf.drawRightString(width - 40, 26, f"Pag. {page_number}")


def draw_programacao_pdf(
    pdf: Any,
    width: float,
    height: float,
    programacao: ProgramacaoDB,
    itens: list[ProgramacaoItemDB],
    equipe_display: str,
    *,
    reimpressao: bool = False,
    reimpressao_info: str = "",
) -> None:
    page = 1
    y = height - 58

    def new_page() -> None:
        nonlocal page, y
        draw_pdf_footer(pdf, width, page)
        pdf.showPage()
        page += 1
        y = height - 58
        draw_header(compact=True)

    def ensure_space(needed: int) -> None:
        if y < needed:
            new_page()

    def draw_header(*, compact: bool = False) -> None:
        nonlocal y
        pdf.setFont("Helvetica-Bold", 14 if not compact else 11)
        pdf.drawString(40, y, f"PROGRAMACAO: {upper_text(programacao.codigo_programacao)}")
        y -= 18 if compact else 22
        pdf.setFont("Helvetica", 9)
        if compact:
            pdf.drawString(40, y, f"Motorista: {upper_text(programacao.motorista)} | Veiculo: {upper_text(programacao.veiculo)}")
            y -= 16

    draw_header()
    emissao = datetime.now().strftime("%d/%m/%Y %H:%M")
    local_rota = format_local_rota_pdf(programacao.local_rota or programacao.tipo_rota)
    local_carreg = local_carregamento_pdf(programacao)
    adiantamento = adiantamento_valor_pdf(programacao)

    header_lines = [
        f"Data: {emissao}",
        f"Motorista: {upper_text(programacao.motorista)} | Codigo: {upper_text(programacao.motorista_codigo or programacao.codigo_motorista) or '-'} | Veiculo: {upper_text(programacao.veiculo)}",
        f"Equipe: {equipe_display or '-'}",
        f"Local da Rota: {local_rota or '-'} | Carregamento: {local_carreg or '-'}",
        programacao_estimativa_pdf(programacao),
        f"NF: {upper_text(programacao.nf_numero or programacao.num_nf) or '-'} | Saida: {data_saida_pdf(programacao)} | Prestacao: {upper_text(programacao.prestacao_status or 'PENDENTE')}",
        f"Adiantamento: {pdf_money_br(adiantamento)} | Origem: {upper_text(programacao.adiantamento_origem) or '-'}",
        f"Criado por: {upper_text(programacao.usuario_criacao) or '-'} | Ultima edicao: {upper_text(programacao.usuario_ultima_edicao) or '-'}",
    ]
    pdf.setFont("Helvetica", 9)
    for line in header_lines:
        pdf.drawString(40, y, line[:126])
        y -= 14
        if line.startswith("Data:") and reimpressao:
            pdf.setFont("Helvetica-Bold", 8)
            info = reimpressao_info or datetime.now().strftime("%d/%m/%Y %H:%M")
            pdf.drawString(40, y, f"REIMPRESSAO - gerada em {info}")
            y -= 12
            pdf.setFont("Helvetica", 9)

    y -= 8
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(40, y, "CLIENTE / ENDERECO")
    pdf.drawRightString(338, y, "CX")
    pdf.drawRightString(404, y, "KG")
    pdf.drawRightString(458, y, "PRECO")
    pdf.drawString(470, y, "VENDEDOR")
    pdf.drawString(528, y, "PEDIDO")
    y -= 8
    pdf.line(40, y, width - 40, y)
    y -= 12

    total_cx = 0
    total_kg = 0.0
    total_valor = 0.0
    pdf.setFont("Helvetica", 7.5)
    for item in itens:
        ensure_space(95)
        caixas = safe_int(item.qnt_caixas, 0)
        kg = safe_float(item.kg, 0.0)
        preco = safe_float(item.preco, 0.0)
        total_cx += caixas
        total_kg += kg
        total_valor += kg * preco
        cliente = f"{upper_text(item.cod_cliente)} - {upper_text(item.nome_cliente)}"
        endereco = upper_text(item.endereco)
        if endereco:
            cliente = f"{cliente} | {endereco}"
        pdf.drawString(40, y, cliente[:70])
        pdf.drawRightString(338, y, str(caixas))
        pdf.drawRightString(404, y, pdf_number_br(kg, 2))
        pdf.drawRightString(458, y, pdf_number_br(preco, 2))
        pdf.drawString(470, y, upper_text(item.vendedor)[:11])
        pdf.drawString(528, y, upper_text(item.pedido)[:14])
        y -= 11
        obs = upper_text(item.observacao)
        pdf.setFont("Helvetica-Oblique", 7)
        pdf.drawString(50, y, (f"OBS: {obs}" if obs else "OBS: ________________________________________________")[:112])
        pdf.setFont("Helvetica", 7.5)
        y -= 13

    ensure_space(70)
    pdf.line(40, y, width - 40, y)
    y -= 14
    pdf.setFont("Helvetica-Bold", 9)
    pdf.drawString(
        40,
        y,
        (
            f"TOTAL CLIENTES: {len(itens)} | TOTAL CX: {total_cx} | "
            f"TOTAL KG: {pdf_number_br(total_kg, 2)} | VALOR PREVISTO: {pdf_money_br(total_valor)}"
        )[:126],
    )
    draw_pdf_footer(pdf, width, page)


def normalize_adiantamento_origem(value: Any) -> str:
    origem = upper_text(value)
    return origem or "NAO INFORMADA"


def draw_adiantamento_receipt_pdf(
    pdf: Any,
    page_width: float,
    page_height: float,
    *,
    codigo: str,
    motorista: str,
    veiculo: str,
    equipe: str,
    valor: float,
    origem: str,
    local_rota: str,
    local_carregamento: str,
    usuario: str,
    data_emissao: str | None = None,
    reimpressao: bool = False,
) -> None:
    from reportlab.lib.units import mm

    valor = safe_float(valor, 0.0)
    if valor <= 0:
        return
    margin = 12 * mm
    gap = 8 * mm
    copies = 2
    block_h = (page_height - (2 * margin) - gap) / copies
    data_txt = data_emissao or datetime.now().strftime("%d/%m/%Y %H:%M")
    labels = ("VIA CAIXA", "VIA MOTORISTA")

    def draw_kv(x: float, y: float, label: str, value: Any, max_chars: int = 48) -> None:
        pdf.setFont("Helvetica-Bold", 8.5)
        pdf.drawString(x, y, label)
        pdf.setFont("Helvetica", 8.5)
        pdf.drawString(x + 33 * mm, y, pdf_text(value)[:max_chars])

    for idx, via in enumerate(labels):
        top = page_height - margin - (idx * (block_h + gap))
        bottom = top - block_h
        if idx > 0:
            pdf.setDash(3, 3)
            pdf.line(margin, top + (gap / 2), page_width - margin, top + (gap / 2))
            pdf.setDash()
        pdf.setLineWidth(0.8)
        pdf.rect(margin, bottom, page_width - (2 * margin), block_h)

        y = top - 7 * mm
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawCentredString(page_width / 2, y, "RECIBO DE ADIANTAMENTO DE ROTA")
        pdf.setFont("Helvetica", 7.5)
        pdf.drawRightString(page_width - margin - 4 * mm, y, via)
        y -= 8 * mm
        if reimpressao:
            pdf.setFont("Helvetica-Bold", 7.5)
            pdf.drawCentredString(page_width / 2, y + 3 * mm, f"REIMPRESSAO - gerada em {data_txt}")
        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(page_width / 2, y, pdf_money_br(valor))
        y -= 8 * mm

        x1 = margin + 5 * mm
        x2 = margin + (page_width - (2 * margin)) / 2 + 4 * mm
        draw_kv(x1, y, "Programacao:", codigo, 36)
        draw_kv(x2, y, "Emissao:", data_txt, 32)
        y -= 6 * mm
        draw_kv(x1, y, "Motorista:", motorista, 42)
        draw_kv(x2, y, "Veiculo:", veiculo or "-", 30)
        y -= 6 * mm
        draw_kv(x1, y, "Equipe:", equipe or "-", 42)
        draw_kv(x2, y, "Origem:", normalize_adiantamento_origem(origem), 30)
        y -= 6 * mm
        draw_kv(x1, y, "Rota:", format_local_rota_pdf(local_rota) or "-", 42)
        draw_kv(x2, y, "Carregamento:", local_carregamento or "-", 30)
        y -= 8 * mm

        pdf.setFont("Helvetica", 8.4)
        receipt_text = (
            "Declaro que o motorista identificado recebeu o valor acima como adiantamento vinculado "
            "exclusivamente a esta programacao, devendo prestar contas no fechamento da rota."
        )
        line = ""
        max_width = page_width - (2 * margin) - (10 * mm)
        for word in receipt_text.split():
            candidate = word if not line else f"{line} {word}"
            if pdf.stringWidth(candidate, "Helvetica", 8.4) <= max_width:
                line = candidate
            else:
                pdf.drawString(x1, y, line)
                y -= 4.2 * mm
                line = word
        if line:
            pdf.drawString(x1, y, line)
            y -= 6 * mm

        pdf.line(x1, bottom + 12 * mm, x1 + 70 * mm, bottom + 12 * mm)
        pdf.line(x2, bottom + 12 * mm, x2 + 70 * mm, bottom + 12 * mm)
        pdf.setFont("Helvetica", 7.5)
        pdf.drawCentredString(x1 + 35 * mm, bottom + 8 * mm, "Assinatura do motorista")
        pdf.drawCentredString(x2 + 35 * mm, bottom + 8 * mm, f"Responsavel: {upper_text(usuario) or '-'}"[:58])


def item_valor_venda(item: ProgramacaoItemDB) -> float:
    return safe_float(item.kg, 0.0) * safe_float(item.preco, 0.0)


def draw_romaneio_page(pdf: Any, width: float, height: float, programacao: ProgramacaoDB, item: ProgramacaoItemDB, index: int, total: int) -> None:
    block_gap = 12
    margin_x = 28
    block_h = (height - 72 - block_gap) / 2
    top_y = height - 36
    labels = ("VIA CLIENTE", "VIA EMPRESA")

    for copy_index, via_label in enumerate(labels):
        y_top = top_y - (copy_index * (block_h + block_gap))
        y_base = y_top - block_h
        draw_romaneio_block(pdf, margin_x, y_base, width - (2 * margin_x), block_h, programacao, item, index, total, via_label)


def draw_romaneio_block(
    pdf: Any,
    x: float,
    y_base: float,
    largura: float,
    altura: float,
    programacao: ProgramacaoDB,
    item: ProgramacaoItemDB,
    index: int,
    total: int,
    via_label: str,
) -> None:
    codigo = upper_text(programacao.codigo_programacao)
    pad = 8
    y = y_base + altura - pad
    caixas = safe_int(item.qnt_caixas, 0)
    aves_por_caixa = safe_int(programacao.qnt_aves_por_cx, 0) or 6
    total_aves = caixas * aves_por_caixa
    kg = safe_float(item.kg, 0.0)
    preco = safe_float(item.preco, 0.0)
    valor_venda = kg * preco
    media_kg_ave = (kg / total_aves) if total_aves > 0 else 0.0

    pdf.setLineWidth(0.8)
    pdf.rect(x, y_base, largura, altura)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawCentredString(x + largura / 2, y - 3, "ROMANEIO DE ENTREGA")
    pdf.setFont("Helvetica", 7)
    pdf.drawRightString(x + largura - pad, y - 3, f"{via_label} | {codigo} {index}/{total}")
    y -= 18
    pdf.line(x + pad, y, x + largura - pad, y)
    y -= 12

    pdf.setFont("Helvetica", 7.5)
    pdf.drawString(x + pad, y, f"COD CLIENTE: {upper_text(item.cod_cliente)}")
    pdf.drawString(x + 145, y, f"PEDIDO: {upper_text(item.pedido) or '-'}")
    pdf.drawRightString(x + largura - pad, y, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")
    y -= 12
    pdf.drawString(x + pad, y, f"RAZAO SOCIAL: {upper_text(item.nome_cliente)[:80]}")
    y -= 11
    pdf.drawString(x + pad, y, f"NOME FANTASIA: {upper_text(item.nome_cliente)[:78]}")
    y -= 11
    pdf.drawString(x + pad, y, f"ENDERECO: {upper_text(item.endereco)[:86]}")
    y -= 13
    pdf.line(x + pad, y, x + largura - pad, y)
    y -= 12

    pdf.setFont("Helvetica-Bold", 7.5)
    pdf.drawString(x + pad, y, f"PRODUTO: {upper_text(item.produto or 'CARGA')[:24]}")
    pdf.drawString(x + 145, y, f"PRECO KG: {pdf_money_br(preco)}")
    pdf.drawString(x + 265, y, f"PESO MEDIO: {pdf_number_br(media_kg_ave, 3)}")
    pdf.drawRightString(x + largura - pad, y, f"LOCAL: {format_local_rota_pdf(programacao.local_rota or programacao.tipo_rota)[:18]}")
    y -= 18

    def box(label: str, value: Any, bx: float, by: float, bw: float = 92, bh: float = 16) -> None:
        pdf.setFont("Helvetica", 6.8)
        pdf.drawString(bx, by + bh + 3, label)
        pdf.rect(bx, by, bw, bh)
        if value not in (None, ""):
            pdf.setFont("Helvetica-Bold", 7.4)
            pdf.drawRightString(bx + bw - 4, by + 5, str(value))

    left_x = x + pad
    mid_x = x + pad + 112
    right_x = x + pad + 250
    row_y = y - 16
    step = 25
    box("Qtd. de Caixas:", caixas, left_x, row_y)
    box("Itens por Caixa:", aves_por_caixa, left_x, row_y - step)
    box("Total de Itens:", total_aves, left_x, row_y - (2 * step))
    box("Peso Total:", pdf_number_br(kg, 2), left_x, row_y - (3 * step))
    box("Valor da Venda:", pdf_money_br(valor_venda), mid_x, row_y, bw=120)
    box("Ocorrencias (und):", "", mid_x, row_y - step, bw=120)
    box("Desc. Ocorr. (R$):", "", mid_x, row_y - (2 * step), bw=120)
    box("Valor Final da venda:", pdf_money_br(valor_venda), mid_x, row_y - (3 * step), bw=120)
    box("Deb. Anterior Cliente:", "", right_x, row_y, bw=largura - (right_x - x) - pad)
    box("Valor recebido:", "", right_x, row_y - step, bw=largura - (right_x - x) - pad)
    box("Forma recebimento:", "", right_x, row_y - (2 * step), bw=largura - (right_x - x) - pad)
    box("Recebido Total:", "", right_x, row_y - (3 * step), bw=largura - (right_x - x) - pad)

    footer_y = y_base + 22
    pdf.setFont("Helvetica-Bold", 6.4)
    pdf.drawCentredString(x + largura / 2, footer_y, "CONTA PARA DEPOSITO BANCO DO BRASIL AGENCIA 0532-0 CONTA CORRENTE 25.852-0")
    pdf.drawCentredString(x + largura / 2, footer_y - 9, "CHAVE PIX: 37.752.738/0001-15 (CNPJ)")
    pdf.setFont("Helvetica", 6.5)
    pdf.drawString(x + pad, y_base + 5, f"PROGRAMACAO: {codigo} | CARREGOU EM: {local_carregamento_pdf(programacao) or '-'}")
    pdf.drawRightString(x + largura - pad, y_base + 5, f"MOTORISTA: {upper_text(programacao.motorista) or '-'}")


async def resolve_motorista(db: AsyncSession, nome: str, codigo: str | None) -> tuple[str, str, int | None]:
    nome_norm = upper_text(nome)
    codigo_norm = upper_text(codigo)
    stmt = select(MotoristaDB)
    if codigo_norm:
        stmt = stmt.where(func.upper(func.coalesce(MotoristaDB.codigo, "")) == codigo_norm)
    else:
        stmt = stmt.where(func.upper(func.coalesce(MotoristaDB.nome, "")) == nome_norm)
    result = await db.execute(stmt.limit(1))
    motorista = result.scalar_one_or_none()
    if motorista:
        return upper_text(motorista.nome), upper_text(motorista.codigo), motorista.id
    return nome_norm, codigo_norm, None


async def validate_recursos_programacao(
    db: AsyncSession,
    payload: ProgramacaoPayload,
    ajudantes: list[str],
    *,
    codigo_programacao: str,
) -> None:
    motorista_nome = upper_text(payload.motorista)
    motorista_codigo = upper_text(payload.motorista_codigo)
    motorista_stmt = select(MotoristaDB)
    if motorista_codigo:
        motorista_stmt = motorista_stmt.where(func.upper(func.coalesce(MotoristaDB.codigo, "")) == motorista_codigo)
    else:
        motorista_stmt = motorista_stmt.where(func.upper(func.coalesce(MotoristaDB.nome, "")) == motorista_nome)
    motorista_result = await db.execute(motorista_stmt.limit(1))
    motorista = motorista_result.scalar_one_or_none()
    if not motorista and motorista_codigo and motorista_nome:
        motorista_result = await db.execute(
            select(MotoristaDB).where(func.upper(func.coalesce(MotoristaDB.nome, "")) == motorista_nome).limit(1)
        )
        motorista = motorista_result.scalar_one_or_none()
    if not motorista or not is_active_status(motorista.status):
        raise HTTPException(status_code=422, detail="Selecione um motorista ativo do cadastro.")
    folgas = await folgas_ativas_programacao(db)
    motorista_nome_db = upper_text(motorista.nome)
    motorista_codigo_db = upper_text(motorista.codigo)
    if (motorista_codigo_db and motorista_codigo_db in folgas["motoristas_codigos"]) or motorista_nome_db in folgas["motoristas_nomes"]:
        raise HTTPException(status_code=409, detail=f"Motorista {motorista_display(motorista.nome, motorista.codigo)} esta em folga.")

    veiculo_placa = upper_text(payload.veiculo)
    veiculo_result = await db.execute(select(VeiculoDB).where(func.upper(func.coalesce(VeiculoDB.placa, "")) == veiculo_placa).limit(1))
    veiculo = veiculo_result.scalar_one_or_none()
    if not veiculo or not is_active_status(getattr(veiculo, "status", "ATIVO")):
        raise HTTPException(status_code=422, detail="Selecione um veiculo ativo do cadastro.")

    ajudantes_result = await db.execute(select(AjudanteDB))
    ajudantes_por_id: dict[str, AjudanteDB] = {}
    ajudantes_por_nome: dict[str, AjudanteDB] = {}
    for item in ajudantes_result.scalars().all():
        display = ajudante_display(item.nome, item.sobrenome, item.id)
        ajudantes_por_id[str(item.id)] = item
        ajudantes_por_nome[upper_text(display)] = item
    for ajudante in ajudantes:
        cadastro = ajudantes_por_id.get(str(ajudante)) or ajudantes_por_nome.get(upper_text(ajudante))
        if not cadastro or not is_active_status(cadastro.status):
            raise HTTPException(status_code=422, detail="Selecione somente ajudantes ativos do cadastro.")
        display = ajudante_display(cadastro.nome, cadastro.sobrenome, cadastro.id)
        if str(cadastro.id) in folgas["ajudantes_ids"] or upper_text(display) in folgas["ajudantes_nomes"]:
            raise HTTPException(status_code=409, detail=f"Ajudante {display} esta em folga.")

    ocupados = await recursos_ocupados_em_rotas_abertas(db, exclude_codigo=codigo_programacao)
    if (motorista_codigo_db and motorista_codigo_db in ocupados["motoristas_codigos"]) or motorista_nome_db in ocupados["motoristas_nomes"]:
        raise HTTPException(status_code=409, detail=f"Motorista {motorista_display(motorista.nome, motorista.codigo)} ja esta vinculado a uma rota aberta.")
    if veiculo_placa in ocupados["veiculos"]:
        raise HTTPException(status_code=409, detail=f"Veiculo {veiculo_placa} ja esta vinculado a uma rota aberta.")
    for ajudante in ajudantes:
        cadastro = ajudantes_por_id.get(str(ajudante)) or ajudantes_por_nome.get(upper_text(ajudante))
        display = ajudante_display(cadastro.nome, cadastro.sobrenome, cadastro.id) if cadastro else upper_text(ajudante)
        if str(ajudante) in ocupados["ajudantes"] or upper_text(display) in ocupados["ajudantes"]:
            raise HTTPException(status_code=409, detail=f"Ajudante {display} ja esta vinculado a uma rota aberta.")


async def validate_capacity(db: AsyncSession, veiculo: str, caixas_para_validar: int) -> None:
    if caixas_para_validar <= 0:
        return
    result = await db.execute(
        select(VeiculoDB).where(func.upper(func.coalesce(VeiculoDB.placa, "")) == upper_text(veiculo)).limit(1)
    )
    vehicle = result.scalar_one_or_none()
    if not vehicle:
        raise HTTPException(status_code=422, detail=f"Veiculo nao encontrado no cadastro: {upper_text(veiculo)}")
    capacidade = safe_int(vehicle.capacidade_cx, -1)
    if capacidade >= 0 and caixas_para_validar > capacidade:
        raise HTTPException(
            status_code=422,
            detail=f"Capacidade excedida para o veiculo {upper_text(veiculo)}. Caixas: {caixas_para_validar}. Capacidade: {capacidade}.",
        )


async def upsert_cliente_from_item(db: AsyncSession, item: ProgramacaoItemPayload) -> None:
    cod = upper_text(item.cod_cliente)
    nome = upper_text(item.nome_cliente)
    if not cod or not nome:
        return
    result = await db.execute(select(ClienteDB).where(func.upper(func.coalesce(ClienteDB.cod_cliente, "")) == cod).limit(1))
    cliente = result.scalar_one_or_none()
    if cliente:
        cliente.nome = nome
        cliente.nome_cliente = nome
        if upper_text(item.endereco):
            cliente.endereco = upper_text(item.endereco)
        if upper_text(item.vendedor):
            cliente.vendedor = upper_text(item.vendedor)
    else:
        db.add(
            ClienteDB(
                cod_cliente=cod,
                nome=nome,
                nome_cliente=nome,
                endereco=upper_text(item.endereco) or None,
                vendedor=upper_text(item.vendedor) or None,
            )
        )


async def apply_programacao_payload(
    db: AsyncSession,
    programacao: ProgramacaoDB,
    payload: ProgramacaoPayload,
    *,
    codigo_programacao: str,
    current_user: User,
    creating: bool,
) -> tuple[int, float, list[ProgramacaoItemPayload]]:
    local_rota = normalize_local_rota(payload.local_rota)
    if local_rota not in {"SERRA", "SERTAO"}:
        raise HTTPException(status_code=422, detail="Selecione o Local da Rota (SERRA ou SERTAO).")

    tipo_estimativa = upper_text(payload.tipo_estimativa or "KG")
    if tipo_estimativa not in {"KG", "CX"}:
        tipo_estimativa = "KG"
    operacao_tipo = normalize_operacao_tipo(payload.operacao_tipo, tipo_estimativa)
    transbordo_modalidade = normalize_transbordo_modalidade(payload.transbordo_modalidade) if operacao_tipo == "TRANSBORDO" else "CIF"

    ajudantes = [upper_text(item) for item in (payload.ajudantes or []) if upper_text(item)]
    if not ajudantes and payload.equipe:
        ajudantes = split_equipe(payload.equipe)
    if len(ajudantes) != 2:
        raise HTTPException(status_code=422, detail="Selecione exatamente 2 ajudantes da programacao.")
    if ajudantes[0] == ajudantes[1]:
        raise HTTPException(status_code=422, detail="Os ajudantes selecionados devem ser diferentes.")
    equipe = "|".join(ajudantes)

    itens = list(payload.itens or [])
    if tipo_estimativa == "CX":
        if safe_int(payload.caixas_estimado, 0) <= 0:
            raise HTTPException(status_code=422, detail="Informe a estimativa em caixas (CX) para EMPRESA BUSCA.")
    elif safe_float(payload.kg_estimado, 0.0) <= 0:
        raise HTTPException(status_code=422, detail="Informe o KG estimado para CIF.")

    total_caixas = sum(safe_int(item.qnt_caixas, 0) for item in itens)
    if total_caixas <= 0 and tipo_estimativa == "CX":
        total_caixas = safe_int(payload.caixas_estimado, 0)
    total_quilos = round(sum(safe_float(item.kg, 0.0) for item in itens), 2)
    if total_quilos <= 0 and tipo_estimativa == "KG":
        total_quilos = round(safe_float(payload.kg_estimado, 0.0), 2)

    caixas_para_validar = safe_int(payload.caixas_estimado, 0) if tipo_estimativa == "CX" else total_caixas
    await validate_recursos_programacao(db, payload, ajudantes, codigo_programacao=codigo_programacao)
    await validate_capacity(db, payload.veiculo, caixas_para_validar)

    motorista_nome, motorista_codigo, motorista_id = await resolve_motorista(db, payload.motorista, payload.motorista_codigo)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    actor = upper_text(current_user.nome or current_user.username) or "ADMIN"

    programacao.codigo_programacao = codigo_programacao
    programacao.codigo = codigo_programacao
    if creating:
        programacao.data = now
        programacao.data_criacao = now
        programacao.usuario_criacao = actor
    elif not programacao.data_criacao:
        programacao.data_criacao = now
    programacao.motorista = motorista_nome
    programacao.motorista_id = motorista_id
    programacao.motorista_codigo = motorista_codigo
    programacao.codigo_motorista = motorista_codigo
    programacao.veiculo = upper_text(payload.veiculo)
    programacao.equipe = equipe
    programacao.kg_estimado = safe_float(payload.kg_estimado, 0.0)
    programacao.tipo_estimativa = tipo_estimativa
    programacao.caixas_estimado = safe_int(payload.caixas_estimado, 0)
    programacao.operacao_tipo = operacao_tipo
    if operacao_tipo == "TRANSBORDO":
        programacao.transbordo_modalidade = transbordo_modalidade
        programacao.transbordo_observacao = upper_text(payload.transbordo_observacao) or None
        programacao.transbordo_grupo = programacao.transbordo_grupo or codigo_programacao
    else:
        programacao.transbordo_modalidade = transbordo_modalidade or "CIF"
        programacao.transbordo_observacao = None
        programacao.transbordo_grupo = None
    if creating:
        programacao.status = "ATIVA"
        programacao.status_operacional = None
        programacao.finalizada_no_app = 0
    programacao.prestacao_status = programacao.prestacao_status or "PENDENTE"
    programacao.local_rota = local_rota
    programacao.tipo_rota = local_rota
    programacao.local_carregamento = upper_text(payload.local_carregamento)
    programacao.granja_carregada = upper_text(payload.local_carregamento)
    programacao.local_carregado = upper_text(payload.local_carregamento)
    programacao.local_carreg = upper_text(payload.local_carregamento)
    programacao.adiantamento = safe_float(payload.adiantamento, 0.0)
    programacao.adiantamento_rota = safe_float(payload.adiantamento, 0.0)
    programacao.adiantamento_origem = upper_text(payload.adiantamento_origem) or None
    programacao.usuario_ultima_edicao = actor
    programacao.total_caixas = safe_int(total_caixas, 0)
    programacao.caixas_carregadas = safe_int(total_caixas, 0)
    programacao.qnt_cx_carregada = safe_int(total_caixas, 0)
    programacao.quilos = safe_float(total_quilos, 0.0)
    programacao.nf_kg = programacao.quilos if tipo_estimativa == "KG" else 0.0
    programacao.nf_caixas = programacao.total_caixas

    return safe_int(total_caixas, 0), safe_float(total_quilos, 0.0), itens


@router.get("/options", response_model=ProgramacaoOptionsResponse)
async def get_programacao_options(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    ocupados = await recursos_ocupados_em_rotas_abertas(db)
    folgas = await folgas_ativas_programacao(db)
    motoristas_result = await db.execute(
        select(MotoristaDB)
        .where(func.upper(func.coalesce(MotoristaDB.status, "ATIVO")) == "ATIVO")
        .order_by(MotoristaDB.nome.asc())
    )
    veiculos_result = await db.execute(
        select(VeiculoDB)
        .where(func.upper(func.coalesce(VeiculoDB.status, "ATIVO")) == "ATIVO")
        .order_by(VeiculoDB.placa.asc())
    )
    ajudantes_result = await db.execute(
        select(AjudanteDB)
        .where(func.upper(func.coalesce(AjudanteDB.status, "ATIVO")) == "ATIVO")
        .order_by(AjudanteDB.nome.asc(), AjudanteDB.sobrenome.asc())
    )
    motoristas = [
        {
            "id": item.id,
            "nome": item.nome or "",
            "codigo": item.codigo or "",
            "display": f"{upper_text(item.nome)} ({upper_text(item.codigo)})" if item.codigo else upper_text(item.nome),
        }
        for item in motoristas_result.scalars().all()
        if upper_text(item.codigo) not in ocupados["motoristas_codigos"] and upper_text(item.nome) not in ocupados["motoristas_nomes"]
        and upper_text(item.codigo) not in folgas["motoristas_codigos"] and upper_text(item.nome) not in folgas["motoristas_nomes"]
    ]
    veiculos = [
        {
            "id": item.id,
            "placa": upper_text(item.placa),
            "modelo": item.modelo or "",
            "capacidade_cx": safe_int(item.capacidade_cx, 0),
            "status": upper_text(getattr(item, "status", "ATIVO")) or "ATIVO",
        }
        for item in veiculos_result.scalars().all()
        if upper_text(item.placa) and upper_text(item.placa) not in ocupados["veiculos"]
    ]
    ajudantes = [
        {
            "id": str(item.id),
            "nome": upper_text(item.nome),
            "sobrenome": upper_text(item.sobrenome),
            "telefone": normalize_phone(item.telefone),
            "display": " ".join(part for part in (upper_text(item.nome), upper_text(item.sobrenome)) if part),
        }
        for item in ajudantes_result.scalars().all()
        if str(item.id) not in ocupados["ajudantes"]
        and " ".join(part for part in (upper_text(item.nome), upper_text(item.sobrenome)) if part) not in ocupados["ajudantes"]
        and str(item.id) not in folgas["ajudantes_ids"]
        and " ".join(part for part in (upper_text(item.nome), upper_text(item.sobrenome)) if part) not in folgas["ajudantes_nomes"]
    ]
    return ProgramacaoOptionsResponse(
        motoristas=motoristas,
        veiculos=veiculos,
        ajudantes=ajudantes,
        proximo_codigo=await next_programacao_codigo(db),
    )


@router.get("/", response_model=list[ProgramacaoResponse])
async def list_programacoes(
    skip: int = 0,
    limit: int = 100,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(
        select(ProgramacaoDB)
        .where(func.trim(func.coalesce(ProgramacaoDB.codigo_programacao, ProgramacaoDB.codigo, "")) != "")
        .order_by(ProgramacaoDB.id.desc())
        .offset(skip)
        .limit(limit)
    )
    return [await serialize_programacao(db, item, include_items=False) for item in result.scalars().all()]


@router.get("/vendas-selecionadas", response_model=ProgramacaoVendasSelecionadasResponse)
async def programacao_vendas_selecionadas(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    return await selected_vendas_importadas_items(db)


@router.post("/sugestao", response_model=ProgramacaoSugestaoResponse)
async def sugerir_programacao(
    payload: ProgramacaoSugestaoPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    itens_payload = payload.itens or []
    venda_ids_by_key: dict[tuple[str, str, str], int] = {}
    if not itens_payload:
        vendas = await selected_vendas_importadas_items(db)
        venda_ids_by_key = {
            (upper_text(item.cod_cliente), upper_text(item.pedido), upper_text(item.produto)): safe_int(item.venda_id, 0)
            for item in vendas.itens
        }
        itens_payload = [
            ProgramacaoItemPayload(
                cod_cliente=item.cod_cliente,
                nome_cliente=item.nome_cliente,
                produto_id=item.produto_id,
                produto=item.produto,
                endereco=item.endereco,
                qnt_caixas=item.qnt_caixas,
                kg=item.kg,
                preco=item.preco,
                vendedor=item.vendedor,
                pedido=item.pedido,
                obs=item.obs,
            )
            for item in vendas.itens
        ]
    if not itens_payload:
        raise HTTPException(status_code=400, detail="Inclua itens ou marque vendas antes de solicitar sugestao.")

    veiculo_placa, capacidade = await require_veiculo_com_capacidade(db, payload.veiculo)
    codigos = {upper_text(item.cod_cliente) for item in itens_payload}
    locs = await localizacoes_clientes(db, codigos)
    enriched: list[dict[str, Any]] = []
    for item in itens_payload:
        cod = upper_text(item.cod_cliente)
        loc = locs.get(cod, {})
        enriched.append(
            {
                "cod_cliente": cod,
                "nome_cliente": upper_text(item.nome_cliente),
                "produto": item.produto or "",
                "endereco": str(item.endereco or loc.get("endereco") or ""),
                "qnt_caixas": safe_int(item.qnt_caixas, 0),
                "kg": safe_float(item.kg, 0.0),
                "preco": safe_float(item.preco, 0.0),
                "vendedor": item.vendedor or "",
                "pedido": item.pedido or "",
                "obs": item.obs or "",
                "lat": loc.get("lat"),
                "lon": loc.get("lon"),
                "cidade": loc.get("cidade") or "",
                "bairro": loc.get("bairro") or "",
                "origem_localizacao": loc.get("origem") or "",
                "amostras_localizacao": safe_int(loc.get("amostras"), 0),
                "ultima_localizacao_em": loc.get("registrado_em") or "",
                "venda_id": venda_ids_by_key.get((cod, upper_text(item.pedido), upper_text(item.produto)), 0),
            }
        )

    historico = await historico_pares_clientes(db, codigos)
    ordered = ordenar_por_vizinho_mais_proximo(enriched, historico)
    total_caixas = sum(max(safe_int(item.get("qnt_caixas"), 0), 0) for item in ordered)
    caixas_acum = 0
    distancia_total = 0.0
    prev = None
    alertas: list[str] = []
    for idx, item in enumerate(ordered, start=1):
        cx = max(safe_int(item.get("qnt_caixas"), 0), 0)
        caixas_acum += cx
        dentro = capacidade <= 0 or caixas_acum <= capacidade
        distancia = 0.0
        if prev and item.get("lat") not in (None, "") and prev.get("lat") not in (None, ""):
            distancia = geo_distance_km(prev.get("lat"), prev.get("lon"), item.get("lat"), item.get("lon"))
            distancia_total += distancia
        item["ordem_sugerida"] = idx
        item["distancia_anterior_km"] = distancia
        item["distancia"] = distancia
        item["dentro_capacidade"] = dentro
        item["confianca_localizacao"] = (
            min(100, 60 + (safe_int(item.get("amostras_localizacao"), 0) * 10))
            if item.get("lat") not in (None, "")
            else 0
        )
        tags = []
        if not item.get("lat") or not item.get("lon"):
            tags.append("SEM GPS")
        if not dentro:
            tags.append("EXCEDE VEICULO")
        if idx > 1 and historico.get((upper_text((ordered[idx - 2] or {}).get("cod_cliente")), upper_text(item.get("cod_cliente"))), 0) > 0:
            item["base_historica"] = historico.get((upper_text((ordered[idx - 2] or {}).get("cod_cliente")), upper_text(item.get("cod_cliente"))), 0)
        item["recomendacao"] = " | ".join(tags) if tags else "OK"
        if item.get("lat") not in (None, ""):
            prev = item

    com_loc = sum(1 for item in ordered if item.get("lat") not in (None, "") and item.get("lon") not in (None, ""))
    sem_loc = len(ordered) - com_loc
    excedente = max(total_caixas - capacidade, 0) if capacidade > 0 else 0
    if excedente > 0:
        alertas.append(f"Excesso de {excedente} caixa(s) em relacao a capacidade do veiculo.")
    if sem_loc:
        alertas.append(f"{sem_loc} cliente(s) sem localizacao historica; ficaram ao final da sugestao.")
    if historico:
        alertas.append("Sugestao considerou programacoes anteriores como apoio para aproximar clientes recorrentes.")

    resumo = (
        f"{len(ordered)} cliente(s), {total_caixas} caixa(s), "
        f"{com_loc} com GPS, distancia estimada {round(distancia_total, 2)} km."
    )
    return ProgramacaoSugestaoResponse(
        veiculo=veiculo_placa,
        capacidade_cx=capacidade,
        total_caixas=total_caixas,
        caixas_dentro_capacidade=min(total_caixas, capacidade) if capacidade > 0 else total_caixas,
        caixas_excedentes=excedente,
        clientes_com_localizacao=com_loc,
        clientes_sem_localizacao=sem_loc,
        distancia_estimativa_km=round(distancia_total, 2),
        resumo=resumo,
        alertas=alertas,
        itens=ordered,
    )


@router.get("/rankings", response_model=ProgramacaoRankingsResponse)
async def programacao_rankings(
    periodo: int = 30,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_dias = max(min(safe_int(periodo, 30), 365), 1)
    motoristas = await ranquear_motoristas(db, periodo_dias)
    ajudantes = await ranquear_ajudantes(db, periodo_dias)
    return ProgramacaoRankingsResponse(
        periodo_dias=periodo_dias,
        motoristas=motoristas,
        ajudantes=ajudantes,
        resumo_motoristas=ranking_summary("Top motoristas", motoristas),
        resumo_ajudantes=ranking_summary("Top ajudantes", ajudantes),
    )


@router.get("/{codigo_programacao}/pdf")
async def programacao_pdf(
    codigo_programacao: str,
    reimpressao: bool = Query(False),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    itens = await items_for_programacao(db, programacao.codigo_programacao)
    if not itens:
        raise HTTPException(status_code=409, detail="Programacao sem itens para gerar PDF.")

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=0)
    width, height = A4
    equipe_display = await resolve_equipe_nomes(db, programacao.equipe)
    draw_programacao_pdf(
        pdf,
        width,
        height,
        programacao,
        itens,
        equipe_display,
        reimpressao=reimpressao,
    )

    adiantamento = adiantamento_valor_pdf(programacao)
    if adiantamento > 0:
        pdf.showPage()
        draw_adiantamento_receipt_pdf(
            pdf,
            width,
            height,
            codigo=upper_text(programacao.codigo_programacao),
            motorista=upper_text(programacao.motorista),
            veiculo=upper_text(programacao.veiculo),
            equipe=equipe_display,
            valor=adiantamento,
            origem=upper_text(programacao.adiantamento_origem),
            local_rota=format_local_rota_pdf(programacao.local_rota or programacao.tipo_rota),
            local_carregamento=local_carregamento_pdf(programacao),
            usuario=upper_text(programacao.usuario_ultima_edicao or programacao.usuario_criacao or current_user.username),
            reimpressao=reimpressao,
        )

    pdf.save()
    buffer.seek(0)
    safe_name = upper_text(programacao.codigo_programacao).replace("/", "_") or "PROGRAMACAO"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="PROGRAMACAO_{safe_name}.pdf"'},
    )


@router.get("/{codigo_programacao}/recibo-adiantamento-pdf")
async def programacao_recibo_adiantamento_pdf(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    adiantamento = adiantamento_valor_pdf(programacao)
    if adiantamento <= 0:
        raise HTTPException(status_code=409, detail="Programacao sem adiantamento para gerar recibo.")

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=0)
    width, height = A4
    draw_adiantamento_receipt_pdf(
        pdf,
        width,
        height,
        codigo=upper_text(programacao.codigo_programacao),
        motorista=upper_text(programacao.motorista),
        veiculo=upper_text(programacao.veiculo),
        equipe=await resolve_equipe_nomes(db, programacao.equipe),
        valor=adiantamento,
        origem=upper_text(programacao.adiantamento_origem),
        local_rota=format_local_rota_pdf(programacao.local_rota or programacao.tipo_rota),
        local_carregamento=local_carregamento_pdf(programacao),
        usuario=upper_text(programacao.usuario_ultima_edicao or programacao.usuario_criacao or current_user.username),
    )
    pdf.save()
    buffer.seek(0)
    safe_name = upper_text(programacao.codigo_programacao).replace("/", "_") or "PROGRAMACAO"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="RECIBO_ADIANTAMENTO_{safe_name}.pdf"'},
    )


@router.get("/{codigo_programacao}/romaneios-pdf")
async def programacao_romaneios_pdf(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    itens = await items_for_programacao(db, programacao.codigo_programacao)
    if not itens:
        raise HTTPException(status_code=409, detail="Programacao sem itens para imprimir romaneios.")

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=0)
    width, height = A4
    total = len(itens)
    for index, item in enumerate(itens, start=1):
        if index > 1:
            pdf.showPage()
        draw_romaneio_page(pdf, width, height, programacao, item, index, total)
    adiantamento = adiantamento_valor_pdf(programacao)
    if adiantamento > 0:
        pdf.showPage()
        draw_adiantamento_receipt_pdf(
            pdf,
            width,
            height,
            codigo=upper_text(programacao.codigo_programacao),
            motorista=upper_text(programacao.motorista),
            veiculo=upper_text(programacao.veiculo),
            equipe=await resolve_equipe_nomes(db, programacao.equipe),
            valor=adiantamento,
            origem=upper_text(programacao.adiantamento_origem),
            local_rota=format_local_rota_pdf(programacao.local_rota or programacao.tipo_rota),
            local_carregamento=local_carregamento_pdf(programacao),
            usuario=upper_text(programacao.usuario_ultima_edicao or programacao.usuario_criacao or current_user.username),
        )
    pdf.save()
    buffer.seek(0)
    safe_name = upper_text(programacao.codigo_programacao).replace("/", "_") or "PROGRAMACAO"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="ROMANEIOS_{safe_name}.pdf"'},
    )


@router.get("/{codigo_programacao}", response_model=ProgramacaoResponse)
async def get_programacao(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    return await serialize_programacao(db, programacao, include_items=True)


@router.post("/", response_model=ProgramacaoResponse, status_code=status.HTTP_201_CREATED)
async def save_programacao(
    payload: ProgramacaoPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    codigo_input = upper_text(payload.codigo_programacao)
    existing = await get_programacao_by_codigo(db, codigo_input) if codigo_input else None
    creating = existing is None
    codigo = codigo_input or await next_programacao_codigo(db)
    programacao = existing or ProgramacaoDB(codigo_programacao=codigo)
    if existing:
        assert_programacao_editable(existing)

    _total_caixas, _total_quilos, itens = await apply_programacao_payload(
        db,
        programacao,
        payload,
        codigo_programacao=codigo,
        current_user=current_user,
        creating=creating,
    )
    if creating:
        db.add(programacao)
    await db.flush()

    if payload.itens is not None:
        await db.execute(delete(ProgramacaoItemDB).where(func.upper(ProgramacaoItemDB.codigo_programacao) == codigo))
        for index, item in enumerate(itens, start=1):
            db.add(
                ProgramacaoItemDB(
                    codigo_programacao=codigo,
                    cod_cliente=upper_text(item.cod_cliente),
                    nome_cliente=upper_text(item.nome_cliente),
                    qnt_caixas=safe_int(item.qnt_caixas, 0),
                    kg=safe_float(item.kg, 0.0),
                    preco=safe_float(item.preco, 0.0),
                    endereco=upper_text(item.endereco) or None,
                    vendedor=upper_text(item.vendedor) or None,
                    pedido=upper_text(item.pedido) or None,
                    produto_id=await produto_id_for_item(db, item.produto_id, item.produto),
                    produto=upper_text(item.produto) or None,
                    observacao=upper_text(item.obs) or None,
                    ordem_sugerida=safe_int(item.ordem_sugerida, 0) or index,
                    distancia=safe_float(item.distancia, 0.0),
                    confianca_localizacao=safe_float(item.confianca_localizacao, 0.0),
                    carga_raiz_programacao=upper_text(item.carga_raiz_programacao) or None,
                    carga_origem_imediata=upper_text(item.carga_origem_imediata) or None,
                    transferencia_origem_id=str(item.transferencia_origem_id or "").strip() or None,
                )
            )
            await upsert_cliente_from_item(db, item)

    vendas_vinculadas = await mark_vendas_importadas_used(db, payload.venda_ids, codigo)

    record_audit_log(
        db,
        action="programacao_criada" if creating else "programacao_alterada",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata={
            "codigo_programacao": codigo,
            "itens": len(itens),
            "total_caixas": programacao.total_caixas,
            "quilos": programacao.quilos,
            "vendas_vinculadas": vendas_vinculadas,
        },
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_programacao(db, programacao, include_items=True)


@router.delete("/{codigo_programacao}", response_model=ProgramacaoResponse)
async def delete_programacao(
    codigo_programacao: str,
    request: Request,
    devolver_vendas: bool = True,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_programacao_deletable(programacao)
    response = await serialize_programacao(db, programacao, include_items=True)
    codigo = upper_text(programacao.codigo_programacao)

    if devolver_vendas:
        try:
            await db.execute(
                text(
                    """
                    UPDATE vendas_importadas
                       SET usada=0, usada_em='', codigo_programacao='', selecionada=0
                     WHERE UPPER(COALESCE(codigo_programacao,''))=:codigo
                    """
                ),
                {"codigo": codigo},
            )
        except Exception:
            pass
    else:
        try:
            await db.execute(
                text("DELETE FROM vendas_importadas WHERE UPPER(COALESCE(codigo_programacao,''))=:codigo"),
                {"codigo": codigo},
            )
        except Exception:
            pass

    for table_name in (
        "programacao_itens_log",
        "programacao_itens_controle",
        "programacao_itens",
        "recebimentos",
        "despesas",
        "rota_gps_pings",
        "rota_substituicoes",
        "cliente_localizacao_amostras",
    ):
        try:
            await db.execute(
                text(f"DELETE FROM {table_name} WHERE UPPER(COALESCE(codigo_programacao,''))=:codigo"),
                {"codigo": codigo},
            )
        except Exception:
            pass

    try:
        await db.execute(
            text(
                """
                DELETE FROM transferencias
                 WHERE UPPER(COALESCE(codigo_origem,''))=:codigo
                    OR UPPER(COALESCE(codigo_destino,''))=:codigo
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        pass

    await db.delete(programacao)
    record_audit_log(
        db,
        action="programacao_excluida",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"codigo_programacao": codigo, "itens": len(response.itens), "devolver_vendas": devolver_vendas},
    )
    await db.commit()
    return response
