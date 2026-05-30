# backend/api/v1/endpoints/centro_custos.py
"""
Centro de Custos endpoints mirroring the desktop CentroCustosPage read model.
"""
from __future__ import annotations

from datetime import datetime, timedelta
import json
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query, status
from pydantic import BaseModel, Field
from sqlalchemy import func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import safe_float, safe_int
from backend.api.v1.endpoints.programacao import normalize_ascii, upper_text
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import VeiculoDB
from backend.models.despesa import DespesaDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemControleDB, ProgramacaoItemDB
from backend.models.recebimento import RecebimentoDB
from backend.models.user import User
from backend.models.venda_importada import VendaImportadaDB

router = APIRouter()

VALID_PERIODOS = {"7", "15", "30", "60", "90", "180", "TODAS"}
VALID_METRICS = {"CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"}
DESPESA_VEICULO_PERFIS = {
    "COMBUSTIVEL": "Combustivel",
    "OLEO_FILTRO": "Oleo e filtros",
    "PNEUS": "Pneus",
    "CORREIAS": "Correias",
    "FREIOS": "Freios",
    "BATERIA": "Bateria",
    "REVISAO": "Revisao",
    "MANUTENCAO": "Manutencao",
    "DOCUMENTACAO": "Documentacao",
    "MULTA": "Multa",
    "SEGURO": "Seguro",
    "LAVAGEM": "Lavagem",
    "OUTROS": "Outros",
}
DESPESA_CONTROLES = {"SEM_CONTROLE", "DATA", "KM", "DATA_KM"}
DESPESA_PRIORIDADES = {"BAIXA", "NORMAL", "ALTA", "CRITICA"}


class CentroCustosOptions(BaseModel):
    periodos: list[str]
    metricas: list[str]
    veiculos: list[str]
    despesa_veiculo_perfis: list[dict[str, str]] = Field(default_factory=list)
    despesa_controles: list[dict[str, str]] = Field(default_factory=list)
    prioridades: list[str] = Field(default_factory=list)


class CentroCustosKpis(BaseModel):
    veiculos: int = 0
    rotas: int = 0
    km_total: float = 0
    kg_carregado: float = 0
    despesas_total: float = 0
    custo_km_global: float = 0
    custo_kg_global: float = 0


class CentroCustosRow(BaseModel):
    veiculo: str
    rotas: int = 0
    km_rodado: float = 0
    kg_carregado: float = 0
    despesas: float = 0
    compra: float = 0
    venda: float = 0
    lucro_liquido: float = 0
    custo_km: float = 0
    custo_kg: float = 0
    ticket_rota: float = 0
    classificacao: str = "medio"


class CentroCustosChartItem(BaseModel):
    label: str
    value: float = 0
    metric: str


class CentroCustosResumoResponse(BaseModel):
    periodo: str
    veiculo: str
    metric: str
    chart_title: str
    resumo: str
    kpis: CentroCustosKpis
    rows: list[CentroCustosRow]
    chart: list[CentroCustosChartItem]


class CentroCustosFinanceiroKpis(BaseModel):
    rotas: int = 0
    compra_total: float = 0
    venda_total: float = 0
    despesas_total: float = 0
    despesas_veiculo: float = 0
    despesas_rota: float = 0
    diarias_total: float = 0
    lucro_bruto: float = 0
    lucro_liquido: float = 0
    margem_liquida: float = 0
    km_total: float = 0
    kg_total: float = 0
    lucro_km: float = 0
    lucro_kg: float = 0
    custo_km: float = 0
    custo_kg: float = 0


class CentroCustosDespesaComposicao(BaseModel):
    grupo: str
    valor: float = 0
    percentual: float = 0


class CentroCustosFinanceiroRow(BaseModel):
    codigo_programacao: str
    data: str = ""
    veiculo: str = ""
    motorista: str = ""
    rota: str = ""
    tipo_estimativa: str = "KG"
    operacao_tipo: str = "VENDA"
    transbordo_modalidade: str = ""
    transbordo_grupo: str = ""
    km_rodado: float = 0
    kg: float = 0
    compra: float = 0
    venda: float = 0
    despesas_veiculo: float = 0
    despesas_rota: float = 0
    diarias: float = 0
    despesas_total: float = 0
    lucro_bruto: float = 0
    lucro_liquido: float = 0
    margem_liquida: float = 0
    custo_km: float = 0
    lucro_km: float = 0
    venda_confirmada: float = 0
    venda_prevista: float = 0
    fonte_venda: str = "SEM_VENDA"
    confianca: str = "BAIXA"
    alertas: list[str] = Field(default_factory=list)
    nivel: int = 0
    parent_codigo: str = ""
    has_children: bool = False
    filhos: list[dict[str, Any]] = Field(default_factory=list)


class CentroCustosFinanceiroResponse(BaseModel):
    periodo: str
    veiculo: str
    kpis: CentroCustosFinanceiroKpis
    composicao: list[CentroCustosDespesaComposicao]
    rows: list[CentroCustosFinanceiroRow]


class CentroCustosDespesaRotaKpis(BaseModel):
    rotas: int = 0
    total: float = 0
    diarias: float = 0
    banhos: float = 0
    guardas: float = 0
    outras: float = 0
    media_rota: float = 0


class CentroCustosDespesaRotaItem(BaseModel):
    id: int
    descricao: str = ""
    tipo: str = "OUTRAS"
    categoria: str = ""
    valor: float = 0
    data_registro: str = ""
    observacao: str = ""


class CentroCustosDespesaRotaRow(BaseModel):
    codigo_programacao: str
    data: str = ""
    veiculo: str = ""
    motorista: str = ""
    rota: str = ""
    diarias: float = 0
    banhos: float = 0
    guardas: float = 0
    outras: float = 0
    total: float = 0
    tipo_totais: dict[str, float] = Field(default_factory=dict)
    qtd_despesas: int = 0
    maior_despesa: str = ""
    despesas: list[CentroCustosDespesaRotaItem] = Field(default_factory=list)


class CentroCustosDespesasRotaResponse(BaseModel):
    periodo: str
    veiculo: str
    kpis: CentroCustosDespesaRotaKpis
    rows: list[CentroCustosDespesaRotaRow]


class CentroCustosDespesaVeiculoPayload(BaseModel):
    veiculo: str = Field(min_length=1, max_length=40)
    valor: float = Field(gt=0)
    descricao: str = Field(min_length=1, max_length=220)
    documento_tipo: str = Field(default="MANUAL", max_length=40)
    documento_numero: str | None = Field(default=None, max_length=80)
    fornecedor: str | None = Field(default=None, max_length=120)
    data_registro: str | None = Field(default=None, max_length=30)
    observacao: str | None = Field(default=None, max_length=300)
    codigo_programacao: str | None = Field(default=None, max_length=80)
    perfil: str | None = Field(default="OUTROS", max_length=40)
    controle_tipo: str | None = Field(default="SEM_CONTROLE", max_length=40)
    data_vencimento: str | None = Field(default=None, max_length=30)
    km_vencimento: float | None = Field(default=None, ge=0)
    odometro: float | None = Field(default=None, ge=0)
    prioridade: str | None = Field(default="NORMAL", max_length=20)


class CentroCustosDespesaVeiculoResponse(BaseModel):
    id: int
    codigo_programacao: str
    veiculo: str
    descricao: str
    valor: float
    categoria: str = ""
    documento: str = ""
    estabelecimento: str = ""
    data_registro: str = ""
    perfil: str = "OUTROS"
    controle_tipo: str = "SEM_CONTROLE"
    data_vencimento: str = ""
    km_vencimento: float = 0
    odometro: float = 0
    prioridade: str = "NORMAL"
    status_controle: str = ""


class CentroCustosVeiculoCard(BaseModel):
    placa: str
    modelo: str = ""
    capacidade_cx: int = 0
    compra_total: float = 0
    venda_total: float = 0
    despesas_total: float = 0
    despesas_manutencao: float = 0
    lucro_liquido: float = 0
    rotas: int = 0
    km_rodado: float = 0
    litros: float = 0
    media_consumo: float = 0
    motoristas: int = 0
    ultima_data: str = ""
    status: str = "SEM_MOVIMENTO"


class CentroCustosVeiculoProgramacao(BaseModel):
    codigo_programacao: str
    data: str = ""
    motorista: str = ""
    rota: str = ""
    km_rodado: float = 0
    litros: float = 0
    media_km_l: float = 0
    despesas_total: float = 0
    status: str = ""


class CentroCustosVeiculoDespesa(BaseModel):
    id: int
    data_registro: str = ""
    descricao: str = ""
    categoria: str = ""
    grupo: str = ""
    valor: float = 0
    documento: str = ""
    estabelecimento: str = ""
    codigo_programacao: str = ""
    odometro: float = 0
    litros: float = 0
    valor_litro: float = 0
    perfil: str = "OUTROS"
    controle_tipo: str = "SEM_CONTROLE"
    data_vencimento: str = ""
    km_vencimento: float = 0
    prioridade: str = "NORMAL"
    status_controle: str = ""


class CentroCustosVeiculoPeca(BaseModel):
    grupo: str
    valor: float = 0
    eventos: int = 0


class CentroCustosVeiculosResponse(BaseModel):
    periodo: str
    veiculo: str
    resumo: CentroCustosFinanceiroKpis
    veiculos: list[CentroCustosVeiculoCard]


class CentroCustosVeiculoDetalheResponse(BaseModel):
    periodo: str
    veiculo: CentroCustosVeiculoCard
    motoristas: list[str]
    programacoes: list[CentroCustosVeiculoProgramacao]
    despesas: list[CentroCustosVeiculoDespesa]
    pecas: list[CentroCustosVeiculoPeca]
    alertas: list[str]


def parse_data_programacao(raw: Any) -> datetime | None:
    txt = str(raw or "").strip()
    if not txt:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d/%m/%y %H:%M:%S", "%d/%m/%y"):
        try:
            return datetime.strptime(txt[:19], fmt) if "H" in fmt else datetime.strptime(txt[:10], fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(txt.replace(" ", "T"))
    except Exception:
        return None


def data_ref(programacao: ProgramacaoDB) -> str:
    return str(programacao.data_saida or programacao.data_criacao or programacao.data or "").strip()


def iso_today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def kg_carregado_ref(programacao: ProgramacaoDB) -> float:
    return (
        safe_float(programacao.nf_kg_carregado, 0.0)
        or safe_float(programacao.kg_carregado, 0.0)
        or safe_float(programacao.nf_kg_vendido, 0.0)
        or safe_float(programacao.kg_estimado, 0.0)
        or safe_float(programacao.nf_kg, 0.0)
        or safe_float(programacao.quilos, 0.0)
    )


def compra_ref(programacao: ProgramacaoDB) -> float:
    kg = safe_float(programacao.nf_kg, 0.0) or safe_float(programacao.kg_nf, 0.0) or kg_carregado_ref(programacao)
    preco = safe_float(programacao.nf_preco, 0.0) or safe_float(programacao.preco_nf, 0.0)
    return money(kg * preco)


def is_transbordo_programacao(programacao: ProgramacaoDB) -> bool:
    tipo = normalize_ascii(getattr(programacao, "operacao_tipo", "")).replace("-", "_").replace(" ", "_")
    return tipo == "TRANSBORDO" or upper_text(getattr(programacao, "tipo_estimativa", "")) == "CX"


def carga_raiz_from_snapshot(value: Any, fallback: Any = "") -> str:
    try:
        data = json.loads(str(value or "{}"))
        if not isinstance(data, dict):
            data = {}
    except Exception:
        data = {}
    return upper_text(data.get("carga_raiz_programacao") or data.get("carga_origem_programacao") or fallback)


def transferencia_qtd_total(row: Any) -> int:
    return max(safe_int(row.get("qtd_caixas"), 0), safe_int(row.get("qtd_convertida"), 0), 0)


def transferencia_qtd_convertida(row: Any) -> int:
    qtd_convertida = max(safe_int(row.get("qtd_convertida"), 0), 0)
    if qtd_convertida > 0:
        return qtd_convertida
    if normalize_ascii(row.get("status")) == "CONVERTIDA":
        return transferencia_qtd_total(row)
    return 0


async def transferencias_compra_por_programacao(db: AsyncSession, codigos: list[str]) -> tuple[dict[str, float], dict[str, float]]:
    codigos_norm = [upper_text(codigo) for codigo in codigos if upper_text(codigo)]
    if not codigos_norm:
        return {}, {}
    try:
        result = await db.execute(
            text(
                """
                SELECT
                    t.codigo_origem,
                    t.codigo_destino,
                    t.qtd_caixas,
                    t.qtd_convertida,
                    t.status,
                    t.snapshot,
                    p.nf_kg,
                    p.kg_nf,
                    p.nf_kg_carregado,
                    p.kg_carregado,
                    p.nf_caixas,
                    p.total_caixas,
                    p.caixas_carregadas,
                    p.nf_preco,
                    p.preco_nf
                  FROM transferencias t
                  LEFT JOIN programacoes p
                    ON UPPER(COALESCE(p.codigo_programacao, ''))=UPPER(COALESCE(t.codigo_origem, ''))
                 WHERE TRIM(COALESCE(t.codigo_origem, '')) <> ''
                    OR TRIM(COALESCE(t.codigo_destino, '')) <> ''
                """
            ),
        )
    except Exception:
        return {}, {}
    rows = list(result.mappings().all())
    roots = {
        carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
        for row in rows
        if carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
    }
    root_map: dict[str, ProgramacaoDB] = {}
    if roots:
        try:
            root_result = await db.execute(select(ProgramacaoDB).where(func.upper(ProgramacaoDB.codigo_programacao).in_(roots)))
            root_map = {upper_text(item.codigo_programacao): item for item in root_result.scalars().all()}
        except Exception:
            root_map = {}
    saida: dict[str, float] = {}
    entrada: dict[str, float] = {}
    for row in rows:
        status_value = normalize_ascii(row.get("status"))
        if status_value in {"CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"}:
            continue
        qtd = transferencia_qtd_convertida(row)
        if qtd <= 0:
            continue
        raiz = carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
        root = root_map.get(raiz)
        kg_origem = (
            safe_float(getattr(root, "nf_kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "nf_kg", 0), 0.0)
            or safe_float(getattr(root, "kg_nf", 0), 0.0)
            or safe_float(row.get("nf_kg_carregado"), 0.0)
            or safe_float(row.get("kg_carregado"), 0.0)
            or safe_float(row.get("nf_kg"), 0.0)
            or safe_float(row.get("kg_nf"), 0.0)
        )
        caixas_origem = (
            safe_int(getattr(root, "nf_caixas", 0), 0)
            or safe_int(getattr(root, "total_caixas", 0), 0)
            or safe_int(getattr(root, "caixas_carregadas", 0), 0)
            or safe_int(row.get("nf_caixas"), 0)
            or safe_int(row.get("total_caixas"), 0)
            or safe_int(row.get("caixas_carregadas"), 0)
        )
        preco = (
            safe_float(getattr(root, "nf_preco", 0), 0.0)
            or safe_float(getattr(root, "preco_nf", 0), 0.0)
            or safe_float(row.get("nf_preco"), 0.0)
            or safe_float(row.get("preco_nf"), 0.0)
        )
        kg_por_caixa = (kg_origem / caixas_origem) if kg_origem > 0 and caixas_origem > 0 else 0.0
        valor = money(qtd * kg_por_caixa * preco) if kg_por_caixa > 0 and preco > 0 else 0.0
        origem = upper_text(row.get("codigo_origem"))
        destino = upper_text(row.get("codigo_destino"))
        if origem not in codigos_norm and destino not in codigos_norm:
            continue
        if origem in codigos_norm:
            saida[origem] = money(saida.get(origem, 0.0) + valor)
        if destino in codigos_norm:
            entrada[destino] = money(entrada.get(destino, 0.0) + valor)
    return saida, entrada


async def transferencias_destinos_por_origem(db: AsyncSession, codigos: list[str]) -> dict[str, list[str]]:
    codigos_norm = [upper_text(codigo) for codigo in codigos if upper_text(codigo)]
    if not codigos_norm:
        return {}
    try:
        result = await db.execute(
            text(
                """
                SELECT codigo_origem, codigo_destino, qtd_caixas, qtd_convertida, status, snapshot
                  FROM transferencias
                 WHERE TRIM(COALESCE(codigo_origem, '')) <> ''
                   AND TRIM(COALESCE(codigo_destino, '')) <> ''
                """
            )
        )
    except Exception:
        return {}
    codigos_set = set(codigos_norm)
    out: dict[str, list[str]] = {}
    for row in result.mappings().all():
        status_value = normalize_ascii(row.get("status"))
        if status_value in {"CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"}:
            continue
        qtd = transferencia_qtd_convertida(row)
        if qtd <= 0:
            continue
        origem = upper_text(row.get("codigo_origem"))
        destino = upper_text(row.get("codigo_destino"))
        raiz = carga_raiz_from_snapshot(row.get("snapshot"), origem) or origem
        origem_ref = raiz if raiz in codigos_set else origem
        if origem_ref not in codigos_set or destino not in codigos_set or destino == origem_ref:
            continue
        destinos = out.setdefault(origem_ref, [])
        if destino not in destinos:
            destinos.append(destino)
    return out


def item_key(cod_cliente: Any, pedido: Any) -> tuple[str, str]:
    return (upper_text(cod_cliente), upper_text(pedido))


def item_vendavel(item: ProgramacaoItemDB, controle: ProgramacaoItemControleDB | None = None) -> bool:
    status = normalize_ascii((controle.status_pedido if controle else item.status_pedido) or "")
    texto = normalize_ascii(f"{item.cod_cliente or ''} {item.nome_cliente or ''} {item.produto or ''} {item.observacao or ''}")
    if status in {"CANCELADO", "CANCELADA"}:
        return False
    if "TRANSB" in texto or "TRANSBORDO" in texto or "RESERVA" in texto:
        return False
    return safe_int(item.qnt_caixas, 0) > 0 or safe_float(item.kg, 0.0) > 0 or safe_float(item.preco, 0.0) > 0


def item_kg_venda(
    item: ProgramacaoItemDB,
    controle: ProgramacaoItemControleDB | None = None,
    *,
    media_kg_caixa: float = 0.0,
) -> float:
    kg_ref = safe_float(
        controle.peso_previsto if controle and safe_float(controle.peso_previsto, 0.0) > 0 else item.kg,
        0.0,
    )
    if kg_ref <= 0:
        caixas = safe_int(controle.caixas_atual if controle and safe_int(controle.caixas_atual, 0) > 0 else item.qnt_caixas, 0)
        media = safe_float(media_kg_caixa, 0.0)
        if caixas > 0 and media > 0:
            kg_ref = caixas * media
    return max(kg_ref, 0.0)


def item_valor_venda(
    item: ProgramacaoItemDB,
    controle: ProgramacaoItemControleDB | None = None,
    *,
    media_kg_caixa: float = 0.0,
) -> float:
    if not item_vendavel(item, controle):
        return 0.0
    kg_ref = item_kg_venda(item, controle, media_kg_caixa=media_kg_caixa)
    preco = (
        safe_float(controle.preco_atual, 0.0)
        if controle and safe_float(controle.preco_atual, 0.0) > 0
        else safe_float(item.preco_atual, 0.0) or safe_float(item.preco, 0.0)
    )
    if kg_ref > 0 and 0 < preco < 1000:
        return money(kg_ref * preco)
    if preco >= 1000:
        return money(preco)
    return 0.0


def melhor_venda_rota(
    *,
    recebimentos: float = 0.0,
    controles: float = 0.0,
    itens: float = 0.0,
    vendas_importadas: float = 0.0,
) -> float:
    for value in (itens, vendas_importadas, controles, recebimentos):
        valor = money(value)
        if valor > 0:
            return valor
    return 0.0


def venda_inteligente_rota(
    *,
    recebimentos: float = 0.0,
    controles: float = 0.0,
    itens: float = 0.0,
    vendas_importadas: float = 0.0,
    itens_vendaveis: int = 0,
    itens_com_recebimento: int = 0,
) -> dict[str, Any]:
    confirmada = money(max(recebimentos, controles, 0.0))
    prevista = money(max(itens, vendas_importadas, 0.0))
    cobertura_recebimento = (safe_int(itens_com_recebimento, 0) / max(safe_int(itens_vendaveis, 0), 1)) if itens_vendaveis > 0 else 0.0
    alertas: list[str] = []
    if confirmada > 0 and prevista > 0:
        diferenca_pct = abs(prevista - confirmada) / max(prevista, confirmada, 1.0)
        if diferenca_pct > 0.15:
            alertas.append("Venda confirmada difere mais de 15% da venda prevista.")
    if confirmada > 0 and (prevista <= 0 or confirmada >= prevista * 0.85 or cobertura_recebimento >= 0.85):
        fonte = "RECEBIMENTOS" if recebimentos >= controles else "CONTROLES"
        confianca = "ALTA"
        valor = confirmada
    elif prevista > 0:
        fonte_base = "ITENS" if itens >= vendas_importadas else "IMPORTACAO"
        if confirmada > 0:
            fonte = f"{fonte_base}_COM_RECEBIMENTO_PARCIAL"
            alertas.append(
                f"Recebimento parcial ({itens_com_recebimento}/{itens_vendaveis} item(ns)); usando venda prevista para nao distorcer o resultado."
            )
        else:
            fonte = fonte_base
            alertas.append("Venda sem recebimento confirmado; usando valor previsto dos itens.")
        confianca = "MEDIA" if itens_vendaveis > 0 else "BAIXA"
        valor = prevista
    else:
        fonte = "SEM_VENDA"
        confianca = "BAIXA"
        valor = 0.0
        alertas.append("Sem venda, recebimento ou item vendavel vinculado.")
    return {
        "valor": money(valor),
        "confirmada": confirmada,
        "prevista": prevista,
        "fonte": fonte,
        "confianca": confianca,
        "alertas": alertas,
    }


def row_dict_public(row: CentroCustosFinanceiroRow) -> dict[str, Any]:
    data = row.model_dump() if hasattr(row, "model_dump") else row.dict()
    data["filhos"] = []
    data["has_children"] = False
    return data


def consolidar_linha_transbordo(
    parent: CentroCustosFinanceiroRow,
    children: list[CentroCustosFinanceiroRow],
) -> CentroCustosFinanceiroRow:
    if not children:
        return parent
    parent_dict = parent.model_dump() if hasattr(parent, "model_dump") else parent.dict()
    numeric_fields = (
        "km_rodado",
        "kg",
        "compra",
        "venda",
        "despesas_veiculo",
        "despesas_rota",
        "diarias",
        "despesas_total",
        "venda_confirmada",
        "venda_prevista",
    )
    for field in numeric_fields:
        parent_dict[field] = safe_float(parent_dict.get(field), 0.0) + sum(safe_float(getattr(child, field), 0.0) for child in children)
    parent_dict["km_rodado"] = round(safe_float(parent_dict.get("km_rodado"), 0.0), 1)
    parent_dict["kg"] = round(safe_float(parent_dict.get("kg"), 0.0), 2)
    for field in ("compra", "venda", "despesas_veiculo", "despesas_rota", "diarias", "despesas_total", "venda_confirmada", "venda_prevista"):
        parent_dict[field] = money(parent_dict.get(field))
    parent_dict["lucro_bruto"] = money(safe_float(parent_dict.get("venda"), 0.0) - safe_float(parent_dict.get("compra"), 0.0))
    parent_dict["lucro_liquido"] = money(safe_float(parent_dict.get("lucro_bruto"), 0.0) - safe_float(parent_dict.get("despesas_total"), 0.0))
    venda = safe_float(parent_dict.get("venda"), 0.0)
    km = safe_float(parent_dict.get("km_rodado"), 0.0)
    parent_dict["margem_liquida"] = round((safe_float(parent_dict.get("lucro_liquido"), 0.0) / venda * 100.0) if venda > 0 else 0.0, 2)
    parent_dict["custo_km"] = three((safe_float(parent_dict.get("despesas_total"), 0.0) / km) if km > 0 else 0.0)
    parent_dict["lucro_km"] = three((safe_float(parent_dict.get("lucro_liquido"), 0.0) / km) if km > 0 else 0.0)
    if venda > 0:
        parent_dict["fonte_venda"] = "TRANSBORDO_CONSOLIDADO"
        child_conf = {upper_text(getattr(child, "confianca", "")) for child in children}
        parent_dict["confianca"] = "ALTA" if child_conf and child_conf <= {"ALTA"} else "MEDIA"
    parent_dict["has_children"] = True
    parent_dict["filhos"] = []
    alertas = [
        alerta
        for alerta in list(parent_dict.get("alertas") or [])
        if not (venda > 0 and "Sem venda" in str(alerta))
    ]
    alertas.append(f"Resultado consolidado com {len(children)} programacao(oes) destino do transbordo.")
    parent_dict["alertas"] = alertas
    return CentroCustosFinanceiroRow(**parent_dict)


def organizar_transbordos(rows: list[CentroCustosFinanceiroRow], links: dict[str, list[str]]) -> list[CentroCustosFinanceiroRow]:
    if not rows or not links:
        return rows
    rows_by_codigo = {upper_text(row.codigo_programacao): row for row in rows if upper_text(row.codigo_programacao)}
    child_codes: set[str] = set()
    result: list[CentroCustosFinanceiroRow] = []
    for row in rows:
        codigo = upper_text(row.codigo_programacao)
        destinos = [dest for dest in links.get(codigo, []) if dest in rows_by_codigo]
        if not destinos:
            continue
        children = [rows_by_codigo[dest] for dest in destinos]
        for child in children:
            child.nivel = 1
            child.parent_codigo = codigo
            child_codes.add(upper_text(child.codigo_programacao))
        parent = consolidar_linha_transbordo(row, children)
        parent.filhos = [row_dict_public(child) for child in children]
        result.append(parent)
    grouped_parents = {upper_text(row.codigo_programacao) for row in result}
    for row in rows:
        codigo = upper_text(row.codigo_programacao)
        if codigo in child_codes or codigo in grouped_parents:
            continue
        result.append(row)
    result.sort(key=lambda item: (item.lucro_liquido, item.codigo_programacao))
    return result


def despesa_grupo(despesa: DespesaDB) -> str:
    text_value = normalize_ascii(f"{despesa.tipo_despesa or ''} {despesa.categoria or ''} {despesa.descricao or ''}")
    if "DIARIA" in text_value:
        return "DIARIAS"
    if any(token in text_value for token in ("CAMINHAO", "CAMINHÃO", "VEICULO", "VEÍCULO", "COMBUSTIVEL", "COMBUSTÍVEL", "MANUTENCAO", "MANUTENÇÃO", "PNEU", "OLEO", "ÓLEO")):
        return "VEICULO"
    return "ROTA"


def despesa_rota_tipo(despesa: DespesaDB) -> str:
    tipo_raw = upper_text(despesa.tipo_despesa)
    if tipo_raw and tipo_raw not in {"ROTA", "OUTRA", "OUTRAS", "DESPESA", "DESPESAS"}:
        return tipo_raw
    text_value = normalize_ascii(
        f"{despesa.tipo_despesa or ''} {despesa.categoria or ''} {despesa.descricao or ''} {despesa.observacao or ''}"
    )
    if "DIARIA" in text_value:
        return "DIARIAS"
    if any(token in text_value for token in ("BANHO", "LAVAGEM", "LAVA JATO", "HIGIENIZACAO")):
        return "BANHOS"
    if any(token in text_value for token in ("GUARDA", "ESTACIONAMENTO", "PERNOITE", "SEGURANCA", "VIGIA")):
        return "GUARDAS"
    categoria = upper_text(despesa.categoria)
    if categoria and categoria not in {"ROTA", "OUTRA", "OUTRAS", "DESPESA", "DESPESAS"}:
        return categoria
    return "OUTRAS"


def peca_grupo(despesa: DespesaDB) -> str:
    meta = despesa_meta(despesa)
    perfil = upper_text(meta.get("perfil"))
    perfil_grupos = {
        "COMBUSTIVEL": "COMBUSTIVEL",
        "OLEO_FILTRO": "OLEOS E FILTROS",
        "PNEUS": "PNEUS",
        "CORREIAS": "CORREIAS",
        "FREIOS": "FREIOS",
        "BATERIA": "BATERIA",
        "REVISAO": "REVISAO",
        "MANUTENCAO": "MANUTENCAO",
        "DOCUMENTACAO": "DOCUMENTACAO",
        "MULTA": "MULTAS",
        "SEGURO": "DOCUMENTACAO",
        "LAVAGEM": "LAVAGEM",
    }
    if perfil in perfil_grupos:
        return perfil_grupos[perfil]
    text_value = normalize_ascii(f"{despesa.categoria or ''} {despesa.descricao or ''} {despesa.observacao or ''}")
    if any(token in text_value for token in ("PNEU", "PNEUS", "BORRACHARIA")):
        return "PNEUS"
    if any(token in text_value for token in ("OLEO", "LUBRIFICANTE", "FILTRO")):
        return "OLEOS E FILTROS"
    if "CORREIA" in text_value:
        return "CORREIAS"
    if any(token in text_value for token in ("COMBUSTIVEL", "DIESEL", "GASOLINA", "ETANOL")):
        return "COMBUSTIVEL"
    if any(token in text_value for token in ("MANUTENCAO", "OFICINA", "MECANICA", "PECAS", "PECA", "REPARO", "FREIO", "BATERIA")):
        return "MANUTENCAO"
    return "OUTROS"


def despesa_meta(despesa: DespesaDB) -> dict[str, Any]:
    raw = str(getattr(despesa, "desktop_web_json", None) or "").strip()
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def normalize_despesa_perfil(value: Any, descricao: Any = "") -> str:
    perfil = upper_text(value).replace("-", "_").replace(" ", "_")
    if perfil in DESPESA_VEICULO_PERFIS:
        return perfil
    text_value = normalize_ascii(descricao)
    if any(token in text_value for token in ("DIESEL", "GASOLINA", "ETANOL", "COMBUSTIVEL")):
        return "COMBUSTIVEL"
    if any(token in text_value for token in ("OLEO", "FILTRO", "LUBRIFICANTE")):
        return "OLEO_FILTRO"
    if "PNEU" in text_value:
        return "PNEUS"
    if "CORREIA" in text_value:
        return "CORREIAS"
    if "FREIO" in text_value:
        return "FREIOS"
    if "BATERIA" in text_value:
        return "BATERIA"
    if any(token in text_value for token in ("REVISAO", "REVISAO", "REVISAO")):
        return "REVISAO"
    if any(token in text_value for token in ("DOCUMENTO", "LICENCIAMENTO", "IPVA", "CRLV")):
        return "DOCUMENTACAO"
    if "MULTA" in text_value:
        return "MULTA"
    return "OUTROS"


def normalize_controle_tipo(value: Any, data_vencimento: Any = "", km_vencimento: Any = 0) -> str:
    controle = upper_text(value).replace("-", "_").replace(" ", "_")
    if controle not in DESPESA_CONTROLES:
        has_data = bool(str(data_vencimento or "").strip())
        has_km = safe_float(km_vencimento, 0.0) > 0
        if has_data and has_km:
            return "DATA_KM"
        if has_data:
            return "DATA"
        if has_km:
            return "KM"
        return "SEM_CONTROLE"
    return controle


def normalize_prioridade(value: Any) -> str:
    prioridade = upper_text(value or "NORMAL").replace("-", "_").replace(" ", "_")
    return prioridade if prioridade in DESPESA_PRIORIDADES else "NORMAL"


def controle_status(meta: dict[str, Any], *, km_atual: float = 0.0, hoje: datetime | None = None) -> str:
    controle = normalize_controle_tipo(meta.get("controle_tipo"), meta.get("data_vencimento"), meta.get("km_vencimento"))
    if controle == "SEM_CONTROLE":
        return ""
    hoje = hoje or datetime.now()
    vencido = False
    proximo = False
    data_venc = parse_data_programacao(meta.get("data_vencimento"))
    if data_venc is not None and controle in {"DATA", "DATA_KM"}:
        dias = (data_venc.date() - hoje.date()).days
        if dias < 0:
            vencido = True
        elif dias <= 15:
            proximo = True
    km_venc = safe_float(meta.get("km_vencimento"), 0.0)
    if km_venc > 0 and controle in {"KM", "DATA_KM"} and km_atual > 0:
        restante = km_venc - km_atual
        if restante < 0:
            vencido = True
        elif restante <= 500:
            proximo = True
    if vencido:
        return "VENCIDO"
    if proximo:
        return "PROXIMO"
    return "EM_DIA"


def money(value: Any) -> float:
    return round(safe_float(value, 0.0), 2)


def three(value: Any) -> float:
    return round(safe_float(value, 0.0), 3)


def normalize_periodo(value: str) -> str:
    periodo = upper_text(value or "30")
    return periodo if periodo in VALID_PERIODOS else "30"


def normalize_metric(value: str) -> str:
    metric = upper_text(value or "CUSTO_KM")
    return metric if metric in VALID_METRICS else "CUSTO_KM"


def cutoff_from_periodo(periodo: str) -> datetime | None:
    if periodo == "TODAS":
        return None
    try:
        return datetime.now() - timedelta(days=int(periodo))
    except Exception:
        return None


def chart_title(metric: str) -> str:
    if metric == "CUSTO_KG":
        return "Custo por Veiculo (Custo/KG)"
    if metric == "DESPESA_TOTAL":
        return "Custo por Veiculo (Despesa Total)"
    return "Custo por Veiculo (Custo/KM)"


def build_resumo_text(kpis: CentroCustosKpis) -> str:
    return (
        f"Veiculos: {kpis.veiculos} | Rotas: {kpis.rotas} | "
        f"KM: {kpis.km_total:.1f} | KG carregado: {kpis.kg_carregado:.2f} | "
        f"Despesas: R$ {kpis.despesas_total:,.2f} | "
        f"Custo/KM global: {kpis.custo_km_global:.3f} | Custo/KG global: {kpis.custo_kg_global:.3f}"
    )


async def centro_custos_base_rows(db: AsyncSession, limit: int) -> list[tuple[ProgramacaoDB, float]]:
    despesas_sum = (
        select(
            DespesaDB.codigo_programacao.label("codigo_programacao"),
            func.coalesce(func.sum(DespesaDB.valor), 0).label("total_desp"),
        )
        .group_by(DespesaDB.codigo_programacao)
        .subquery()
    )
    result = await db.execute(
        select(ProgramacaoDB, func.coalesce(despesas_sum.c.total_desp, 0))
        .outerjoin(despesas_sum, despesas_sum.c.codigo_programacao == ProgramacaoDB.codigo_programacao)
        .where(func.trim(func.coalesce(ProgramacaoDB.veiculo, "")) != "")
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    return [(row[0], safe_float(row[1], 0.0)) for row in result.all()]


def programacao_passes_filters(programacao: ProgramacaoDB, periodo: str, veiculo: str, cutoff: datetime | None = None) -> bool:
    veiculo_row = upper_text(programacao.veiculo)
    if not veiculo_row:
        return False
    if cutoff is not None:
        dt_ref = parse_data_programacao(data_ref(programacao))
        if dt_ref is None or dt_ref < cutoff:
            return False
    if veiculo and veiculo != "TODOS" and veiculo != veiculo_row:
        return False
    return True


def despesa_passes_filters(despesa: DespesaDB, periodo: str, veiculo: str, cutoff: datetime | None = None) -> bool:
    veiculo_row = upper_text(despesa.veiculo)
    if not veiculo_row:
        return False
    if cutoff is not None:
        dt_ref = parse_data_programacao(despesa.data_registro or despesa.registrado_em)
        if dt_ref is None or dt_ref < cutoff:
            return False
    if veiculo and veiculo != "TODOS" and veiculo != veiculo_row:
        return False
    return True


async def centro_custos_financeiro_data(
    db: AsyncSession,
    periodo: str,
    veiculo: str,
    limit: int,
    *,
    agrupar_transbordo: bool = True,
) -> CentroCustosFinanceiroResponse:
    cutoff = None
    if periodo != "TODAS":
        try:
            cutoff = datetime.now() - timedelta(days=int(periodo))
        except Exception:
            cutoff = None

    prog_result = await db.execute(
        select(ProgramacaoDB)
        .where(func.trim(func.coalesce(ProgramacaoDB.veiculo, "")) != "")
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    programacoes_periodo = [
        item
        for item in prog_result.scalars().all()
        if programacao_passes_filters(item, periodo, "TODOS", cutoff)
    ]
    programacoes = [
        item
        for item in programacoes_periodo
        if programacao_passes_filters(item, periodo, veiculo, cutoff)
    ]
    transferencias_links: dict[str, list[str]] = {}
    if agrupar_transbordo:
        codigos_periodo = [upper_text(item.codigo_programacao) for item in programacoes_periodo if upper_text(item.codigo_programacao)]
        transferencias_links = await transferencias_destinos_por_origem(db, codigos_periodo)
        if veiculo and veiculo != "TODOS":
            por_codigo_periodo = {upper_text(item.codigo_programacao): item for item in programacoes_periodo if upper_text(item.codigo_programacao)}
            codigos_visiveis = {upper_text(item.codigo_programacao) for item in programacoes if upper_text(item.codigo_programacao)}
            for codigo_origem in list(codigos_visiveis):
                for codigo_destino in transferencias_links.get(codigo_origem, []):
                    destino = por_codigo_periodo.get(codigo_destino)
                    if destino is not None and codigo_destino not in codigos_visiveis:
                        programacoes.append(destino)
                        codigos_visiveis.add(codigo_destino)
    codigos = [upper_text(item.codigo_programacao) for item in programacoes if upper_text(item.codigo_programacao)]
    programacao_por_codigo = {upper_text(item.codigo_programacao): item for item in programacoes if upper_text(item.codigo_programacao)}

    venda_itens_por_prog: dict[str, float] = {}
    kg_itens_por_prog: dict[str, float] = {}
    itens_vendaveis_por_prog: dict[str, int] = {}
    itens_com_recebimento_por_prog: dict[str, int] = {}
    controles_por_chave: dict[tuple[str, str, str], ProgramacaoItemControleDB] = {}
    if codigos:
        controle_result = await db.execute(select(ProgramacaoItemControleDB).where(func.upper(ProgramacaoItemControleDB.codigo_programacao).in_(codigos)))
        for controle in controle_result.scalars().all():
            codigo = upper_text(controle.codigo_programacao)
            controles_por_chave[(codigo, *item_key(controle.cod_cliente, controle.pedido))] = controle

        item_result = await db.execute(select(ProgramacaoItemDB).where(func.upper(ProgramacaoItemDB.codigo_programacao).in_(codigos)))
        for item in item_result.scalars().all():
            codigo = upper_text(item.codigo_programacao)
            controle = controles_por_chave.get((codigo, *item_key(item.cod_cliente, item.pedido)))
            programacao_ref = programacao_por_codigo.get(codigo)
            caixas_total = safe_int(getattr(programacao_ref, "total_caixas", 0), 0) if programacao_ref else 0
            kg_rota = kg_carregado_ref(programacao_ref) if programacao_ref else 0.0
            media_kg_caixa = (kg_rota / caixas_total) if caixas_total > 0 and kg_rota > 0 else 0.0
            kg = item_kg_venda(item, controle, media_kg_caixa=media_kg_caixa)
            valor_item = item_valor_venda(item, controle, media_kg_caixa=media_kg_caixa)
            if item_vendavel(item, controle):
                itens_vendaveis_por_prog[codigo] = safe_int(itens_vendaveis_por_prog.get(codigo), 0) + 1
                if controle and safe_float(controle.valor_recebido, 0.0) > 0:
                    itens_com_recebimento_por_prog[codigo] = safe_int(itens_com_recebimento_por_prog.get(codigo), 0) + 1
            venda_itens_por_prog[codigo] = safe_float(venda_itens_por_prog.get(codigo), 0.0) + valor_item
            kg_itens_por_prog[codigo] = safe_float(kg_itens_por_prog.get(codigo), 0.0) + max(kg, 0.0)

    venda_controles_por_prog: dict[str, float] = {}
    for (codigo, _cod_cliente, _pedido), controle in controles_por_chave.items():
        valor_recebido = safe_float(controle.valor_recebido, 0.0)
        if valor_recebido > 0:
            venda_controles_por_prog[codigo] = safe_float(venda_controles_por_prog.get(codigo), 0.0) + valor_recebido

    venda_importada_por_prog: dict[str, float] = {}
    if codigos:
        vendas_result = await db.execute(select(VendaImportadaDB).where(func.upper(VendaImportadaDB.codigo_programacao).in_(codigos)))
        for venda_importada in vendas_result.scalars().all():
            codigo = upper_text(venda_importada.codigo_programacao)
            venda_importada_por_prog[codigo] = safe_float(venda_importada_por_prog.get(codigo), 0.0) + max(safe_float(venda_importada.vr_total, 0.0), 0.0)

    recebimentos_por_prog: dict[str, float] = {}
    if codigos:
        receb_result = await db.execute(select(RecebimentoDB).where(func.upper(RecebimentoDB.codigo_programacao).in_(codigos)))
        for recebimento in receb_result.scalars().all():
            codigo = upper_text(recebimento.codigo_programacao)
            recebimentos_por_prog[codigo] = safe_float(recebimentos_por_prog.get(codigo), 0.0) + max(safe_float(recebimento.valor, 0.0), 0.0)

    despesas_por_prog: dict[str, dict[str, float]] = {}
    if codigos:
        desp_result = await db.execute(select(DespesaDB).where(func.upper(DespesaDB.codigo_programacao).in_(codigos)))
        for despesa in desp_result.scalars().all():
            codigo = upper_text(despesa.codigo_programacao)
            grupos = despesas_por_prog.setdefault(codigo, {"VEICULO": 0.0, "ROTA": 0.0, "DIARIAS": 0.0})
            grupo = despesa_grupo(despesa)
            grupos[grupo] = safe_float(grupos.get(grupo), 0.0) + safe_float(despesa.valor, 0.0)

    transferencia_saida_compra, transferencia_entrada_compra = await transferencias_compra_por_programacao(db, codigos)

    rows: list[CentroCustosFinanceiroRow] = []
    codigos_set = set(codigos)
    for programacao in programacoes:
        codigo = upper_text(programacao.codigo_programacao)
        despesas = despesas_por_prog.get(codigo, {"VEICULO": 0.0, "ROTA": 0.0, "DIARIAS": 0.0})
        desp_veic = money(despesas.get("VEICULO", 0.0))
        desp_rota = money(despesas.get("ROTA", 0.0))
        diarias = money(despesas.get("DIARIAS", 0.0))
        desp_total = money(desp_veic + desp_rota + diarias)
        venda_info = venda_inteligente_rota(
            recebimentos=recebimentos_por_prog.get(codigo, 0.0),
            controles=venda_controles_por_prog.get(codigo, 0.0),
            itens=venda_itens_por_prog.get(codigo, 0.0),
            vendas_importadas=venda_importada_por_prog.get(codigo, 0.0),
            itens_vendaveis=safe_int(itens_vendaveis_por_prog.get(codigo), 0),
            itens_com_recebimento=safe_int(itens_com_recebimento_por_prog.get(codigo), 0),
        )
        venda = safe_float(venda_info.get("valor"), 0.0)
        compra = money(max(compra_ref(programacao) - transferencia_saida_compra.get(codigo, 0.0), 0.0) + transferencia_entrada_compra.get(codigo, 0.0))
        if venda <= 0 and safe_int(itens_vendaveis_por_prog.get(codigo), 0) <= 0 and not transferencia_entrada_compra.get(codigo, 0.0):
            compra = 0.0
        lucro_bruto = money(venda - compra)
        lucro_liquido = money(lucro_bruto - desp_total)
        km = safe_float(programacao.km_rodado, 0.0)
        kg = safe_float(kg_itens_por_prog.get(codigo), 0.0) or kg_carregado_ref(programacao)
        margem = (lucro_liquido / venda * 100.0) if venda > 0 else 0.0
        alertas = list(venda_info.get("alertas") or [])
        if km <= 0 and desp_total > 0:
            alertas.append("Despesa registrada sem KM rodado para custo/KM.")
        if kg <= 0 and (venda > 0 or compra > 0):
            alertas.append("Movimento financeiro sem KG vinculado.")
        if desp_total <= 0 and (venda > 0 or compra > 0):
            alertas.append("Rota com compra/venda sem despesas registradas.")
        rows.append(
            CentroCustosFinanceiroRow(
                codigo_programacao=codigo,
                data=str(data_ref(programacao) or "")[:10],
                veiculo=upper_text(programacao.veiculo),
                motorista=upper_text(programacao.motorista),
                rota=upper_text(programacao.local_rota or programacao.tipo_rota),
                tipo_estimativa=upper_text(programacao.tipo_estimativa or "KG"),
                operacao_tipo="TRANSBORDO" if is_transbordo_programacao(programacao) else "VENDA",
                transbordo_modalidade=upper_text(getattr(programacao, "transbordo_modalidade", "") or ""),
                transbordo_grupo=upper_text(getattr(programacao, "transbordo_grupo", "") or ""),
                km_rodado=round(km, 1),
                kg=round(kg, 2),
                compra=compra,
                venda=venda,
                despesas_veiculo=desp_veic,
                despesas_rota=desp_rota,
                diarias=diarias,
                despesas_total=desp_total,
                lucro_bruto=lucro_bruto,
                lucro_liquido=lucro_liquido,
                margem_liquida=round(margem, 2),
                custo_km=three((desp_total / km) if km > 0 else 0.0),
                lucro_km=three((lucro_liquido / km) if km > 0 else 0.0),
                venda_confirmada=money(venda_info.get("confirmada", 0.0)),
                venda_prevista=money(venda_info.get("prevista", 0.0)),
                fonte_venda=str(venda_info.get("fonte") or "SEM_VENDA"),
                confianca=str(venda_info.get("confianca") or "BAIXA"),
                alertas=alertas,
            )
        )

    extras_result = await db.execute(
        select(DespesaDB)
        .where(func.trim(func.coalesce(DespesaDB.veiculo, "")) != "")
        .order_by(DespesaDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    extras_por_veiculo: dict[str, dict[str, Any]] = {}
    for despesa in extras_result.scalars().all():
        codigo_despesa = upper_text(despesa.codigo_programacao)
        if codigo_despesa in codigos_set:
            continue
        if not despesa_passes_filters(despesa, periodo, veiculo, cutoff):
            continue
        veiculo_extra = upper_text(despesa.veiculo)
        grupo = despesa_grupo(despesa)
        bucket = extras_por_veiculo.setdefault(
            veiculo_extra,
            {"data": "", "VEICULO": 0.0, "ROTA": 0.0, "DIARIAS": 0.0},
        )
        bucket[grupo] = safe_float(bucket.get(grupo), 0.0) + safe_float(despesa.valor, 0.0)
        data_despesa = str(despesa.data_registro or despesa.registrado_em or "").strip()[:10]
        if data_despesa and data_despesa > str(bucket.get("data") or ""):
            bucket["data"] = data_despesa

    for veiculo_extra, valores in extras_por_veiculo.items():
        desp_veic = money(valores.get("VEICULO", 0.0))
        desp_rota = money(valores.get("ROTA", 0.0))
        diarias = money(valores.get("DIARIAS", 0.0))
        desp_total = money(desp_veic + desp_rota + diarias)
        rows.append(
            CentroCustosFinanceiroRow(
                codigo_programacao=f"DESPESA-{veiculo_extra}",
                data=str(valores.get("data") or ""),
                veiculo=veiculo_extra,
                rota="DESPESA AVULSA",
                despesas_veiculo=desp_veic,
                despesas_rota=desp_rota,
                diarias=diarias,
                despesas_total=desp_total,
                lucro_bruto=0.0,
                lucro_liquido=money(-desp_total),
                fonte_venda="DESPESA_AVULSA",
                confianca="ALTA",
            )
        )

    if agrupar_transbordo:
        rows = organizar_transbordos(rows, transferencias_links)
    rows.sort(key=lambda item: (item.lucro_liquido, item.codigo_programacao))
    compra_total = sum(row.compra for row in rows)
    venda_total = sum(row.venda for row in rows)
    desp_veiculo_total = sum(row.despesas_veiculo for row in rows)
    desp_rota_total = sum(row.despesas_rota for row in rows)
    diarias_total = sum(row.diarias for row in rows)
    desp_total = desp_veiculo_total + desp_rota_total + diarias_total
    lucro_bruto = venda_total - compra_total
    lucro_liquido = lucro_bruto - desp_total
    km_total = sum(row.km_rodado for row in rows)
    kg_total = sum(row.kg for row in rows)
    rotas_count = sum(1 + len(row.filhos or []) for row in rows)
    kpis = CentroCustosFinanceiroKpis(
        rotas=rotas_count,
        compra_total=money(compra_total),
        venda_total=money(venda_total),
        despesas_total=money(desp_total),
        despesas_veiculo=money(desp_veiculo_total),
        despesas_rota=money(desp_rota_total),
        diarias_total=money(diarias_total),
        lucro_bruto=money(lucro_bruto),
        lucro_liquido=money(lucro_liquido),
        margem_liquida=round((lucro_liquido / venda_total * 100.0) if venda_total > 0 else 0.0, 2),
        km_total=round(km_total, 1),
        kg_total=round(kg_total, 2),
        lucro_km=three((lucro_liquido / km_total) if km_total > 0 else 0.0),
        lucro_kg=three((lucro_liquido / kg_total) if kg_total > 0 else 0.0),
        custo_km=three((desp_total / km_total) if km_total > 0 else 0.0),
        custo_kg=three((desp_total / kg_total) if kg_total > 0 else 0.0),
    )
    composicao = [
        CentroCustosDespesaComposicao(grupo="Veiculo/Caminhao", valor=money(desp_veiculo_total), percentual=round((desp_veiculo_total / desp_total * 100.0) if desp_total > 0 else 0.0, 2)),
        CentroCustosDespesaComposicao(grupo="Rota", valor=money(desp_rota_total), percentual=round((desp_rota_total / desp_total * 100.0) if desp_total > 0 else 0.0, 2)),
        CentroCustosDespesaComposicao(grupo="Diarias", valor=money(diarias_total), percentual=round((diarias_total / desp_total * 100.0) if desp_total > 0 else 0.0, 2)),
    ]
    return CentroCustosFinanceiroResponse(
        periodo=periodo,
        veiculo=veiculo or "TODOS",
        kpis=kpis,
        composicao=composicao,
        rows=rows,
    )


def aggregate_rows(
    base_rows: list[tuple[ProgramacaoDB, float]],
    periodo: str,
    veiculo: str,
    metric: str,
) -> CentroCustosResumoResponse:
    cutoff = None
    if periodo != "TODAS":
        try:
            cutoff = datetime.now() - timedelta(days=int(periodo))
        except Exception:
            cutoff = None

    agg: dict[str, dict[str, float | int]] = {}
    for programacao, total_desp in base_rows:
        veiculo_row = upper_text(programacao.veiculo)
        if not veiculo_row:
            continue
        dt_ref = parse_data_programacao(data_ref(programacao))
        if cutoff is not None and (dt_ref is None or dt_ref < cutoff):
            continue
        if veiculo and veiculo != "TODOS" and veiculo != veiculo_row:
            continue

        item = agg.setdefault(veiculo_row, {"rotas": 0, "km": 0.0, "kg": 0.0, "desp": 0.0})
        item["rotas"] = safe_int(item["rotas"], 0) + 1
        item["km"] = safe_float(item["km"], 0.0) + safe_float(programacao.km_rodado, 0.0)
        item["kg"] = safe_float(item["kg"], 0.0) + kg_carregado_ref(programacao)
        item["desp"] = safe_float(item["desp"], 0.0) + safe_float(total_desp, 0.0)

    rows_tmp = []
    for veic, data in sorted(agg.items(), key=lambda kv: (-safe_float(kv[1].get("desp", 0), 0.0), kv[0])):
        rotas = safe_int(data["rotas"], 0)
        km = safe_float(data["km"], 0.0)
        kg = safe_float(data["kg"], 0.0)
        desp = safe_float(data["desp"], 0.0)
        custo_km = (desp / km) if km > 0 else 0.0
        custo_kg = (desp / kg) if kg > 0 else 0.0
        ticket = (desp / rotas) if rotas > 0 else 0.0
        rows_tmp.append((veic, rotas, km, kg, desp, custo_km, custo_kg, ticket))

    media_custo_km = (sum(row[5] for row in rows_tmp) / float(len(rows_tmp))) if rows_tmp else 0.0
    rows = []
    for veic, rotas, km, kg, desp, custo_km, custo_kg, ticket in rows_tmp:
        if media_custo_km <= 0:
            classificacao = "medio"
        elif custo_km > media_custo_km * 1.2:
            classificacao = "alto"
        elif custo_km < media_custo_km * 0.9:
            classificacao = "bom"
        else:
            classificacao = "medio"
        rows.append(
            CentroCustosRow(
                veiculo=veic,
                rotas=rotas,
                km_rodado=round(km, 1),
                kg_carregado=round(kg, 2),
                despesas=money(desp),
                custo_km=three(custo_km),
                custo_kg=three(custo_kg),
                ticket_rota=money(ticket),
                classificacao=classificacao,
            )
        )

    total_rotas = sum(row.rotas for row in rows)
    total_km = sum(row.km_rodado for row in rows)
    total_kg = sum(row.kg_carregado for row in rows)
    total_desp = sum(row.despesas for row in rows)
    kpis = CentroCustosKpis(
        veiculos=len(rows),
        rotas=total_rotas,
        km_total=round(total_km, 1),
        kg_carregado=round(total_kg, 2),
        despesas_total=money(total_desp),
        custo_km_global=three((total_desp / total_km) if total_km > 0 else 0.0),
        custo_kg_global=three((total_desp / total_kg) if total_kg > 0 else 0.0),
    )

    metric_attr = {
        "CUSTO_KG": "custo_kg",
        "DESPESA_TOTAL": "despesas",
        "CUSTO_KM": "custo_km",
    }[metric]
    chart_rows = sorted(rows, key=lambda row: safe_float(getattr(row, metric_attr), 0.0), reverse=True)
    chart = [
        CentroCustosChartItem(label=row.veiculo, value=safe_float(getattr(row, metric_attr), 0.0), metric=metric)
        for row in chart_rows[:10]
    ]
    return CentroCustosResumoResponse(
        periodo=periodo,
        veiculo=veiculo or "TODOS",
        metric=metric,
        chart_title=chart_title(metric),
        resumo=build_resumo_text(kpis),
        kpis=kpis,
        rows=rows,
        chart=chart,
    )


def aggregate_financeiro_rows(
    financeiro: CentroCustosFinanceiroResponse,
    metric: str,
) -> CentroCustosResumoResponse:
    agg: dict[str, dict[str, float | int]] = {}
    for row in financeiro.rows:
        veiculo_row = upper_text(row.veiculo)
        if not veiculo_row:
            continue
        item = agg.setdefault(
            veiculo_row,
            {
                "rotas": 0,
                "km": 0.0,
                "kg": 0.0,
                "desp": 0.0,
                "compra": 0.0,
                "venda": 0.0,
                "lucro": 0.0,
            },
        )
        if not str(row.codigo_programacao or "").startswith("DESPESA-"):
            item["rotas"] = safe_int(item["rotas"], 0) + 1
        item["km"] = safe_float(item["km"], 0.0) + safe_float(row.km_rodado, 0.0)
        item["kg"] = safe_float(item["kg"], 0.0) + safe_float(row.kg, 0.0)
        item["desp"] = safe_float(item["desp"], 0.0) + safe_float(row.despesas_total, 0.0)
        item["compra"] = safe_float(item["compra"], 0.0) + safe_float(row.compra, 0.0)
        item["venda"] = safe_float(item["venda"], 0.0) + safe_float(row.venda, 0.0)
        item["lucro"] = safe_float(item["lucro"], 0.0) + safe_float(row.lucro_liquido, 0.0)

    rows_tmp = []
    for veic, data in sorted(agg.items(), key=lambda kv: (-safe_float(kv[1].get("desp", 0), 0.0), kv[0])):
        rotas = safe_int(data["rotas"], 0)
        km = safe_float(data["km"], 0.0)
        kg = safe_float(data["kg"], 0.0)
        desp = safe_float(data["desp"], 0.0)
        compra = safe_float(data["compra"], 0.0)
        venda = safe_float(data["venda"], 0.0)
        lucro = safe_float(data["lucro"], 0.0)
        custo_km = (desp / km) if km > 0 else 0.0
        custo_kg = (desp / kg) if kg > 0 else 0.0
        ticket = (desp / rotas) if rotas > 0 else desp
        rows_tmp.append((veic, rotas, km, kg, desp, compra, venda, lucro, custo_km, custo_kg, ticket))

    media_custo_km = (sum(row[8] for row in rows_tmp) / float(len(rows_tmp))) if rows_tmp else 0.0
    rows = []
    for veic, rotas, km, kg, desp, compra, venda, lucro, custo_km, custo_kg, ticket in rows_tmp:
        if lucro < 0:
            classificacao = "alto"
        elif media_custo_km > 0 and custo_km < media_custo_km * 0.9:
            classificacao = "bom"
        elif media_custo_km > 0 and custo_km > media_custo_km * 1.2:
            classificacao = "alto"
        else:
            classificacao = "medio"
        rows.append(
            CentroCustosRow(
                veiculo=veic,
                rotas=rotas,
                km_rodado=round(km, 1),
                kg_carregado=round(kg, 2),
                despesas=money(desp),
                compra=money(compra),
                venda=money(venda),
                lucro_liquido=money(lucro),
                custo_km=three(custo_km),
                custo_kg=three(custo_kg),
                ticket_rota=money(ticket),
                classificacao=classificacao,
            )
        )

    total_rotas = sum(row.rotas for row in rows)
    total_km = sum(row.km_rodado for row in rows)
    total_kg = sum(row.kg_carregado for row in rows)
    total_desp = sum(row.despesas for row in rows)
    kpis = CentroCustosKpis(
        veiculos=len(rows),
        rotas=total_rotas,
        km_total=round(total_km, 1),
        kg_carregado=round(total_kg, 2),
        despesas_total=money(total_desp),
        custo_km_global=three((total_desp / total_km) if total_km > 0 else 0.0),
        custo_kg_global=three((total_desp / total_kg) if total_kg > 0 else 0.0),
    )
    metric_attr = {
        "CUSTO_KG": "custo_kg",
        "DESPESA_TOTAL": "despesas",
        "CUSTO_KM": "custo_km",
    }[metric]
    chart_rows = sorted(rows, key=lambda row: safe_float(getattr(row, metric_attr), 0.0), reverse=True)
    chart = [
        CentroCustosChartItem(label=row.veiculo, value=safe_float(getattr(row, metric_attr), 0.0), metric=metric)
        for row in chart_rows[:10]
    ]
    return CentroCustosResumoResponse(
        periodo=financeiro.periodo,
        veiculo=financeiro.veiculo or "TODOS",
        metric=metric,
        chart_title=chart_title(metric),
        resumo=build_resumo_text(kpis),
        kpis=kpis,
        rows=rows,
        chart=chart,
    )


async def centro_custos_veiculos_data(
    db: AsyncSession,
    periodo: str,
    veiculo: str,
    limit: int,
) -> CentroCustosVeiculosResponse:
    cutoff = cutoff_from_periodo(periodo)
    cad_result = await db.execute(select(VeiculoDB).order_by(VeiculoDB.placa.asc()))
    veiculos_cadastro = {upper_text(item.placa): item for item in cad_result.scalars().all() if upper_text(item.placa)}

    prog_result = await db.execute(
        select(ProgramacaoDB)
        .where(func.trim(func.coalesce(ProgramacaoDB.veiculo, "")) != "")
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    programacoes = [
        item
        for item in prog_result.scalars().all()
        if programacao_passes_filters(item, periodo, veiculo, cutoff)
    ]

    desp_result = await db.execute(
        select(DespesaDB)
        .where(func.trim(func.coalesce(DespesaDB.veiculo, "")) != "")
        .order_by(DespesaDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    despesas = [
        item
        for item in desp_result.scalars().all()
        if despesa_passes_filters(item, periodo, veiculo, cutoff)
    ]

    placas = set(veiculos_cadastro.keys())
    placas.update(upper_text(item.veiculo) for item in programacoes if upper_text(item.veiculo))
    placas.update(upper_text(item.veiculo) for item in despesas if upper_text(item.veiculo))
    if veiculo and veiculo != "TODOS":
        placas = {item for item in placas if item == veiculo}

    prog_por_placa: dict[str, list[ProgramacaoDB]] = {}
    for item in programacoes:
        prog_por_placa.setdefault(upper_text(item.veiculo), []).append(item)

    desp_por_placa: dict[str, list[DespesaDB]] = {}
    for item in despesas:
        desp_por_placa.setdefault(upper_text(item.veiculo), []).append(item)

    financeiro = await centro_custos_financeiro_data(db, periodo, veiculo, limit, agrupar_transbordo=False)
    financeiro_por_placa: dict[str, dict[str, float]] = {}
    for row in financeiro.rows:
        placa_row = upper_text(row.veiculo)
        if not placa_row:
            continue
        bucket = financeiro_por_placa.setdefault(
            placa_row,
            {"compra": 0.0, "venda": 0.0, "despesas": 0.0, "lucro": 0.0},
        )
        bucket["compra"] = safe_float(bucket.get("compra"), 0.0) + safe_float(row.compra, 0.0)
        bucket["venda"] = safe_float(bucket.get("venda"), 0.0) + safe_float(row.venda, 0.0)
        bucket["despesas"] = safe_float(bucket.get("despesas"), 0.0) + safe_float(row.despesas_total, 0.0)
        bucket["lucro"] = safe_float(bucket.get("lucro"), 0.0) + safe_float(row.lucro_liquido, 0.0)

    cards: list[CentroCustosVeiculoCard] = []
    for placa in sorted(placas):
        cadastro = veiculos_cadastro.get(placa)
        progs = prog_por_placa.get(placa, [])
        desps = desp_por_placa.get(placa, [])
        financeiro_placa = financeiro_por_placa.get(placa, {})
        despesas_total = safe_float(financeiro_placa.get("despesas"), 0.0)
        if despesas_total <= 0:
            despesas_total = sum(safe_float(item.valor, 0.0) for item in desps)
        despesas_manut = sum(safe_float(item.valor, 0.0) for item in desps if peca_grupo(item) != "COMBUSTIVEL")
        km_total = sum(safe_float(item.km_rodado, 0.0) for item in progs)
        litros_total = sum(safe_float(item.litros, 0.0) for item in progs) + sum(safe_float(item.litros, 0.0) for item in desps)
        motoristas = {upper_text(item.motorista) for item in progs if upper_text(item.motorista)}
        datas = [str(data_ref(item) or "")[:10] for item in progs if data_ref(item)]
        datas.extend(str(item.data_registro or item.registrado_em or "")[:10] for item in desps if item.data_registro or item.registrado_em)
        media_consumo = (km_total / litros_total) if litros_total > 0 else 0.0
        if despesas_total > 0 or progs:
            status_value = "EM_CONTROLE"
        else:
            status_value = "SEM_MOVIMENTO"
        cards.append(
            CentroCustosVeiculoCard(
                placa=placa,
                modelo=str(getattr(cadastro, "modelo", "") or ""),
                capacidade_cx=safe_int(getattr(cadastro, "capacidade_cx", 0), 0),
                compra_total=money(financeiro_placa.get("compra", 0.0)),
                venda_total=money(financeiro_placa.get("venda", 0.0)),
                despesas_total=money(despesas_total),
                despesas_manutencao=money(despesas_manut),
                lucro_liquido=money(financeiro_placa.get("lucro", 0.0)),
                rotas=len(progs),
                km_rodado=round(km_total, 1),
                litros=round(litros_total, 2),
                media_consumo=three(media_consumo),
                motoristas=len(motoristas),
                ultima_data=max(datas) if datas else "",
                status=status_value,
            )
        )

    return CentroCustosVeiculosResponse(
        periodo=periodo,
        veiculo=veiculo or "TODOS",
        resumo=financeiro.kpis,
        veiculos=cards,
    )


async def centro_custos_veiculo_detalhe_data(
    db: AsyncSession,
    placa: str,
    periodo: str,
    limit: int,
) -> CentroCustosVeiculoDetalheResponse:
    placa_norm = upper_text(placa)
    base = await centro_custos_veiculos_data(db, periodo, placa_norm, limit)
    card = next((item for item in base.veiculos if item.placa == placa_norm), None)
    if card is None:
        raise HTTPException(status_code=404, detail="Veiculo nao encontrado no centro de custos.")

    cutoff = cutoff_from_periodo(periodo)
    prog_result = await db.execute(
        select(ProgramacaoDB)
        .where(func.upper(ProgramacaoDB.veiculo) == placa_norm)
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    programacoes_db = [
        item
        for item in prog_result.scalars().all()
        if programacao_passes_filters(item, periodo, placa_norm, cutoff)
    ]
    codigos = [upper_text(item.codigo_programacao) for item in programacoes_db if upper_text(item.codigo_programacao)]

    despesas_result = await db.execute(
        select(DespesaDB)
        .where(func.upper(DespesaDB.veiculo) == placa_norm)
        .order_by(DespesaDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    despesas_db = [
        item
        for item in despesas_result.scalars().all()
        if despesa_passes_filters(item, periodo, placa_norm, cutoff)
    ]

    despesas_por_codigo: dict[str, float] = {}
    for despesa in despesas_db:
        codigo = upper_text(despesa.codigo_programacao)
        if codigo:
            despesas_por_codigo[codigo] = safe_float(despesas_por_codigo.get(codigo), 0.0) + safe_float(despesa.valor, 0.0)

    km_atual = max(
        [
            safe_float(item.km_final, 0.0)
            for item in programacoes_db
            if safe_float(item.km_final, 0.0) > 0
        ]
        + [
            safe_float(item.odometro, 0.0)
            for item in despesas_db
            if safe_float(item.odometro, 0.0) > 0
        ]
        + [0.0]
    )

    programacoes = [
        CentroCustosVeiculoProgramacao(
            codigo_programacao=upper_text(item.codigo_programacao),
            data=str(data_ref(item) or "")[:10],
            motorista=upper_text(item.motorista),
            rota=upper_text(item.local_rota or item.tipo_rota),
            km_rodado=round(safe_float(item.km_rodado, 0.0), 1),
            litros=round(safe_float(item.litros, 0.0), 2),
            media_km_l=three(safe_float(item.media_km_l, 0.0) or ((safe_float(item.km_rodado, 0.0) / safe_float(item.litros, 0.0)) if safe_float(item.litros, 0.0) > 0 else 0.0)),
            despesas_total=money(despesas_por_codigo.get(upper_text(item.codigo_programacao), 0.0)),
            status=upper_text(item.status),
        )
        for item in programacoes_db
    ]

    despesas = []
    for item in despesas_db:
        meta = despesa_meta(item)
        perfil = normalize_despesa_perfil(meta.get("perfil") or item.categoria, item.descricao)
        controle = normalize_controle_tipo(meta.get("controle_tipo"), meta.get("data_vencimento"), meta.get("km_vencimento"))
        despesas.append(
            CentroCustosVeiculoDespesa(
                id=item.id,
                data_registro=str(item.data_registro or item.registrado_em or "")[:10],
                descricao=upper_text(item.descricao),
                categoria=upper_text(item.categoria),
                grupo=peca_grupo(item),
                valor=money(item.valor),
                documento=str(item.documento or "").strip(),
                estabelecimento=upper_text(item.estabelecimento),
                codigo_programacao=upper_text(item.codigo_programacao),
                odometro=round(safe_float(item.odometro, 0.0), 1),
                litros=round(safe_float(item.litros, 0.0), 2),
                valor_litro=money(item.valor_litro),
                perfil=perfil,
                controle_tipo=controle,
                data_vencimento=str(meta.get("data_vencimento") or "")[:10],
                km_vencimento=round(safe_float(meta.get("km_vencimento"), 0.0), 1),
                prioridade=normalize_prioridade(meta.get("prioridade")),
                status_controle=controle_status(meta, km_atual=km_atual),
            )
        )

    pecas_map: dict[str, dict[str, float | int]] = {}
    for item in despesas_db:
        grupo = peca_grupo(item)
        bucket = pecas_map.setdefault(grupo, {"valor": 0.0, "eventos": 0})
        bucket["valor"] = safe_float(bucket["valor"], 0.0) + safe_float(item.valor, 0.0)
        bucket["eventos"] = safe_int(bucket["eventos"], 0) + 1
    pecas = [
        CentroCustosVeiculoPeca(grupo=grupo, valor=money(data["valor"]), eventos=safe_int(data["eventos"], 0))
        for grupo, data in sorted(pecas_map.items(), key=lambda pair: (-safe_float(pair[1]["valor"], 0.0), pair[0]))
    ]

    motoristas = sorted({upper_text(item.motorista) for item in programacoes_db if upper_text(item.motorista)})
    alertas: list[str] = []
    if len(motoristas) > 1:
        alertas.append(f"Veiculo rodou com {len(motoristas)} motoristas no periodo.")
    if card.media_consumo <= 0 and card.km_rodado > 0:
        alertas.append("KM informado sem litros para calcular consumo.")
    elif card.media_consumo > 0 and card.media_consumo < 2:
        alertas.append("Media de consumo baixa para o periodo.")
    if card.despesas_manutencao > card.despesas_total * 0.7 and card.despesas_total > 0:
        alertas.append("Manutencao representa mais de 70% dos gastos do veiculo.")
    controles_vencidos = [item for item in despesas if item.status_controle == "VENCIDO"]
    controles_proximos = [item for item in despesas if item.status_controle == "PROXIMO"]
    if controles_vencidos:
        principais = ", ".join(f"{item.grupo}: {item.descricao}" for item in controles_vencidos[:3])
        alertas.append(f"Controle vencido: {principais}.")
    if controles_proximos:
        principais = ", ".join(f"{item.grupo}: {item.descricao}" for item in controles_proximos[:3])
        alertas.append(f"Controle proximo do prazo: {principais}.")
    if not alertas:
        alertas.append("Sem alertas criticos para o filtro atual.")

    return CentroCustosVeiculoDetalheResponse(
        periodo=periodo,
        veiculo=card,
        motoristas=motoristas,
        programacoes=programacoes,
        despesas=despesas,
        pecas=pecas,
        alertas=alertas,
    )


@router.get("/options", response_model=CentroCustosOptions)
async def centro_custos_options(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacoes_result = await db.execute(
        select(func.upper(func.trim(ProgramacaoDB.veiculo)))
        .where(func.trim(func.coalesce(ProgramacaoDB.veiculo, "")) != "")
        .distinct()
        .order_by(func.upper(func.trim(ProgramacaoDB.veiculo)).asc())
    )
    cadastros_result = await db.execute(
        select(func.upper(func.trim(VeiculoDB.placa)))
        .where(func.trim(func.coalesce(VeiculoDB.placa, "")) != "")
        .distinct()
    )
    despesas_result = await db.execute(
        select(func.upper(func.trim(DespesaDB.veiculo)))
        .where(func.trim(func.coalesce(DespesaDB.veiculo, "")) != "")
        .distinct()
    )
    veiculos_set = {
        upper_text(item)
        for item in [
            *programacoes_result.scalars().all(),
            *cadastros_result.scalars().all(),
            *despesas_result.scalars().all(),
        ]
        if upper_text(item)
    }
    veiculos = sorted(veiculos_set)
    return CentroCustosOptions(
        periodos=["7", "15", "30", "60", "90", "180", "TODAS"],
        metricas=["CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"],
        veiculos=["TODOS", *veiculos],
        despesa_veiculo_perfis=[
            {"codigo": codigo, "nome": nome}
            for codigo, nome in DESPESA_VEICULO_PERFIS.items()
        ],
        despesa_controles=[
            {"codigo": "SEM_CONTROLE", "nome": "Sem controle"},
            {"codigo": "DATA", "nome": "Vencimento por data"},
            {"codigo": "KM", "nome": "Vencimento por KM"},
            {"codigo": "DATA_KM", "nome": "Data e KM"},
        ],
        prioridades=["BAIXA", "NORMAL", "ALTA", "CRITICA"],
    )


@router.get("/resumo", response_model=CentroCustosResumoResponse)
async def centro_custos_resumo(
    periodo: str = Query(default="30"),
    veiculo: str = Query(default="TODOS"),
    metric: str = Query(default="CUSTO_KM"),
    limit: int = Query(default=5000, ge=1, le=20000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_norm = normalize_periodo(periodo)
    veiculo_norm = upper_text(veiculo or "TODOS") or "TODOS"
    metric_norm = normalize_metric(metric)
    financeiro = await centro_custos_financeiro_data(db, periodo_norm, veiculo_norm, limit, agrupar_transbordo=False)
    return aggregate_financeiro_rows(financeiro, metric_norm)


@router.get("/financeiro", response_model=CentroCustosFinanceiroResponse)
async def centro_custos_financeiro(
    periodo: str = Query(default="30"),
    veiculo: str = Query(default="TODOS"),
    limit: int = Query(default=5000, ge=1, le=20000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_norm = normalize_periodo(periodo)
    veiculo_norm = upper_text(veiculo or "TODOS") or "TODOS"
    return await centro_custos_financeiro_data(db, periodo_norm, veiculo_norm, limit, agrupar_transbordo=True)


@router.get("/veiculos", response_model=CentroCustosVeiculosResponse)
async def centro_custos_veiculos(
    periodo: str = Query(default="30"),
    veiculo: str = Query(default="TODOS"),
    limit: int = Query(default=5000, ge=1, le=20000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_norm = normalize_periodo(periodo)
    veiculo_norm = upper_text(veiculo or "TODOS") or "TODOS"
    return await centro_custos_veiculos_data(db, periodo_norm, veiculo_norm, limit)


@router.get("/despesas-rota", response_model=CentroCustosDespesasRotaResponse)
async def centro_custos_despesas_rota(
    periodo: str = Query(default="30"),
    veiculo: str = Query(default="TODOS"),
    limit: int = Query(default=10000, ge=1, le=20000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    periodo_norm = normalize_periodo(periodo)
    veiculo_norm = upper_text(veiculo or "TODOS") or "TODOS"
    cutoff = cutoff_from_periodo(periodo_norm)

    prog_result = await db.execute(
        select(ProgramacaoDB)
        .where(func.trim(func.coalesce(ProgramacaoDB.codigo_programacao, "")) != "")
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 20000), 1))
    )
    programacoes = [
        item
        for item in prog_result.scalars().all()
        if programacao_passes_filters(item, periodo_norm, veiculo_norm, cutoff)
    ]
    codigos = [upper_text(item.codigo_programacao) for item in programacoes if upper_text(item.codigo_programacao)]
    if not codigos:
        return CentroCustosDespesasRotaResponse(
            periodo=periodo_norm,
            veiculo=veiculo_norm,
            kpis=CentroCustosDespesaRotaKpis(),
            rows=[],
        )

    despesas_result = await db.execute(select(DespesaDB).where(func.upper(DespesaDB.codigo_programacao).in_(codigos)))
    despesas_por_codigo: dict[str, list[DespesaDB]] = {}
    for despesa in despesas_result.scalars().all():
        codigo = upper_text(despesa.codigo_programacao)
        if codigo:
            despesas_por_codigo.setdefault(codigo, []).append(despesa)

    rows: list[CentroCustosDespesaRotaRow] = []
    for programacao in programacoes:
        codigo = upper_text(programacao.codigo_programacao)
        despesas = despesas_por_codigo.get(codigo, [])
        totais = {"DIARIAS": 0.0, "BANHOS": 0.0, "GUARDAS": 0.0, "OUTRAS": 0.0}
        tipo_totais: dict[str, float] = {}
        maior = ("", 0.0)
        for despesa in despesas:
            tipo = despesa_rota_tipo(despesa)
            valor = money(despesa.valor)
            totais[tipo] = safe_float(totais.get(tipo), 0.0) + valor
            tipo_totais[tipo] = money(safe_float(tipo_totais.get(tipo), 0.0) + valor)
            if valor > maior[1]:
                maior = (upper_text(despesa.descricao or despesa.categoria or tipo), valor)
        total = money(sum(tipo_totais.values()))
        if total <= 0:
            continue
        rows.append(
            CentroCustosDespesaRotaRow(
                codigo_programacao=codigo,
                data=str(data_ref(programacao) or "")[:10],
                veiculo=upper_text(programacao.veiculo),
                motorista=upper_text(programacao.motorista),
                rota=upper_text(programacao.local_rota or programacao.tipo_rota),
                diarias=money(totais["DIARIAS"]),
                banhos=money(totais["BANHOS"]),
                guardas=money(totais["GUARDAS"]),
                outras=money(totais["OUTRAS"]),
                total=total,
                tipo_totais=tipo_totais,
                qtd_despesas=len(despesas),
                maior_despesa=f"{maior[0]} - R$ {maior[1]:.2f}" if maior[0] else "",
                despesas=[
                    CentroCustosDespesaRotaItem(
                        id=despesa.id,
                        descricao=upper_text(despesa.descricao),
                        tipo=despesa_rota_tipo(despesa),
                        categoria=upper_text(despesa.categoria),
                        valor=money(despesa.valor),
                        data_registro=str(despesa.data_registro or ""),
                        observacao=str(despesa.observacao or "").strip(),
                    )
                    for despesa in despesas
                ],
            )
        )

    rows.sort(key=lambda item: item.total, reverse=True)
    total = money(sum(row.total for row in rows))
    kpis = CentroCustosDespesaRotaKpis(
        rotas=len(rows),
        total=total,
        diarias=money(sum(row.diarias for row in rows)),
        banhos=money(sum(row.banhos for row in rows)),
        guardas=money(sum(row.guardas for row in rows)),
        outras=money(sum(row.outras for row in rows)),
        media_rota=money(total / len(rows)) if rows else 0.0,
    )
    return CentroCustosDespesasRotaResponse(periodo=periodo_norm, veiculo=veiculo_norm, kpis=kpis, rows=rows)


@router.get("/veiculos/{placa}", response_model=CentroCustosVeiculoDetalheResponse)
async def centro_custos_veiculo_detalhe(
    placa: str,
    periodo: str = Query(default="30"),
    limit: int = Query(default=5000, ge=1, le=20000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_norm = normalize_periodo(periodo)
    return await centro_custos_veiculo_detalhe_data(db, placa, periodo_norm, limit)


@router.post("/despesas-veiculo", response_model=CentroCustosDespesaVeiculoResponse, status_code=status.HTTP_201_CREATED)
async def criar_despesa_veiculo_centro_custos(
    payload: CentroCustosDespesaVeiculoPayload,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    veiculo = upper_text(payload.veiculo)
    codigo_programacao = upper_text(payload.codigo_programacao)
    motorista = ""
    if codigo_programacao:
        result = await db.execute(select(ProgramacaoDB).where(func.upper(ProgramacaoDB.codigo_programacao) == codigo_programacao))
        programacao = result.scalar_one_or_none()
        if not programacao:
            raise HTTPException(status_code=404, detail="Programacao informada nao encontrada.")
        veiculo = upper_text(programacao.veiculo or veiculo)
        motorista = upper_text(programacao.motorista)
    else:
        codigo_programacao = f"VEICULO:{veiculo}"

    documento_tipo = upper_text(payload.documento_tipo or "MANUAL")
    if documento_tipo not in {"CUPOM FISCAL", "NOTA FISCAL", "MANUAL"}:
        documento_tipo = "MANUAL"
    documento = upper_text(payload.documento_numero)
    if documento and documento_tipo != "MANUAL":
        documento = f"{documento_tipo}: {documento}"
    elif documento_tipo != "MANUAL":
        documento = documento_tipo
    perfil = normalize_despesa_perfil(payload.perfil, payload.descricao)
    controle_tipo = normalize_controle_tipo(payload.controle_tipo, payload.data_vencimento, payload.km_vencimento)
    prioridade = normalize_prioridade(payload.prioridade)
    meta = {
        "perfil": perfil,
        "perfil_label": DESPESA_VEICULO_PERFIS.get(perfil, perfil),
        "controle_tipo": controle_tipo,
        "data_vencimento": str(payload.data_vencimento or "").strip()[:10],
        "km_vencimento": safe_float(payload.km_vencimento, 0.0),
        "prioridade": prioridade,
        "registrado_por": str(current_user.nome or current_user.username or "ADMIN").strip(),
    }

    despesa = DespesaDB(
        codigo_programacao=codigo_programacao,
        descricao=upper_text(payload.descricao),
        valor=money(payload.valor),
        data_registro=str(payload.data_registro or "").strip()[:30] or iso_today(),
        tipo_despesa="VEICULO",
        categoria=perfil,
        motorista=motorista,
        veiculo=veiculo,
        observacao=str(payload.observacao or "").strip(),
        estabelecimento=upper_text(payload.fornecedor),
        documento=documento,
        odometro=safe_float(payload.odometro, 0.0),
        origem="WEB_CENTRO_CUSTOS",
        registrado_em=datetime.now().isoformat(timespec="seconds"),
        desktop_web_json=json.dumps(meta, ensure_ascii=False, sort_keys=True),
    )
    db.add(despesa)
    await db.commit()
    await db.refresh(despesa)
    return CentroCustosDespesaVeiculoResponse(
        id=despesa.id,
        codigo_programacao=upper_text(despesa.codigo_programacao),
        veiculo=upper_text(despesa.veiculo),
        descricao=upper_text(despesa.descricao),
        valor=money(despesa.valor),
        categoria=upper_text(despesa.categoria),
        documento=str(despesa.documento or "").strip(),
        estabelecimento=upper_text(despesa.estabelecimento),
        data_registro=str(despesa.data_registro or ""),
        perfil=perfil,
        controle_tipo=controle_tipo,
        data_vencimento=str(meta.get("data_vencimento") or ""),
        km_vencimento=round(safe_float(meta.get("km_vencimento"), 0.0), 1),
        odometro=round(safe_float(despesa.odometro, 0.0), 1),
        prioridade=prioridade,
        status_controle=controle_status(meta, km_atual=safe_float(despesa.odometro, 0.0)),
    )
