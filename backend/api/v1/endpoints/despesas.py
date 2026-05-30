# backend/api/v1/endpoints/despesas.py
"""
Despesas endpoints mirroring the desktop DespesasPage core flow.
"""
from __future__ import annotations

from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import re
from typing import Any
import json
from uuid import uuid4

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, Request, UploadFile, status
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel, Field
from sqlalchemy import func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import safe_float, safe_int
from backend.api.v1.endpoints.programacao import get_programacao_by_codigo, upper_text
from backend.api.v1.endpoints.recebimentos import (
    ajudante_map as recebimentos_ajudante_map,
    diarias_for_programacao,
)
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import AjudanteDB
from backend.models.despesa import DespesaDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemControleDB, ProgramacaoItemDB
from backend.models.recebimento import RecebimentoDB
from backend.models.user import User
from backend.services.audit import client_ip_from_request, record_audit_log
from backend.services.roteiro_operacional import registrar_roteiro_operacional

router = APIRouter()

CANCELLED_STATUSES = {"CANCELADA", "CANCELADO"}
CEDULAS = (200, 100, 50, 20, 10, 5, 2)
DIARIA_DESCRICAO_MOTORISTA = "DIARIAS MOTORISTA"
DIARIA_DESCRICAO_AJUDANTES = "DIARIAS AJUDANTES"
DIARIA_DESCRICOES_AUTO = {DIARIA_DESCRICAO_MOTORISTA, DIARIA_DESCRICAO_AJUDANTES}


class DespesaProgramacaoOption(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    veiculo: str = ""
    status: str = ""
    prestacao_status: str = "PENDENTE"
    fechada: bool = False


class DespesasCabecalho(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    veiculo: str = ""
    equipe: str = ""
    rota: str = ""
    local_carregamento: str = ""
    data_saida: str = ""
    hora_saida: str = ""
    data_chegada: str = ""
    hora_chegada: str = ""
    status: str = ""
    prestacao_status: str = "PENDENTE"
    adiantamento_origem: str = ""
    operacao_tipo: str = "VENDA"
    transbordo_modalidade: str = ""
    transbordo_observacao: str = ""
    fechada: bool = False


class DespesaItem(BaseModel):
    id: int
    codigo_programacao: str
    descricao: str = ""
    valor: float = 0
    data_registro: str = ""
    tipo_despesa: str = "ROTA"
    categoria: str = ""
    motorista: str = ""
    veiculo: str = ""
    observacao: str = ""
    id_local: str = ""
    forma_pagamento: str = ""
    comprovante_path: str = ""
    estabelecimento: str = ""
    documento: str = ""
    litros: float = 0
    valor_litro: float = 0
    desconto: float = 0
    combustivel: str = ""
    odometro: float = 0
    lat: float | None = None
    lon: float | None = None
    accuracy: float | None = None
    origem: str = ""
    foto: dict[str, Any] = Field(default_factory=dict)


class RotaResumo(BaseModel):
    km_inicial: float = 0
    km_final: float = 0
    litros: float = 0
    km_rodado: float = 0
    media_km_l: float = 0
    custo_km: float = 0
    rota_observacao: str = ""


class NfResumo(BaseModel):
    nf_numero: str = ""
    nf_kg: float = 0
    nf_preco: float = 0
    nf_caixas: int = 0
    nf_kg_carregado: float = 0
    nf_kg_vendido: float = 0
    nf_saldo: float = 0
    nf_media_carregada: float = 0
    nf_caixa_final: int = 0
    mortalidade_transbordo_aves: int = 0
    mortalidade_transbordo_kg: float = 0
    obs_transbordo: str = ""
    kg_nf_util: float = 0
    nf_saldo_valor: float = 0
    desconto_fornecedor: float = 0
    total_compra_bruta: float = 0
    total_compra_liquida: float = 0
    preco_medio_venda: float = 0
    total_compra: float = 0
    receita_estimada: float = 0
    despesas_rota: float = 0
    lucro_bruto: float = 0
    lucro_liquido: float = 0
    margem_liquida: float = 0


class FinanceiroResumo(BaseModel):
    total_recebido: float = 0
    total_despesas: float = 0
    adiantamento: float = 0
    valor_dinheiro: float = 0
    pix_motorista: float = 0
    total_entradas: float = 0
    total_saidas: float = 0
    valor_final_caixa: float = 0
    total_devolvido: float = 0
    diferenca: float = 0
    resultado_liquido: float = 0
    cedulas: dict[str, int] = Field(default_factory=dict)


class OperacionalResumo(BaseModel):
    clientes_total: int = 0
    pedidos_entregues: int = 0
    pedidos_pendentes: int = 0
    pedidos_cancelados: int = 0
    pedidos_alterados: int = 0
    caixas_programadas: int = 0
    caixas_carregadas: int = 0
    caixas_entregues: int = 0
    caixas_canceladas: int = 0
    caixas_transferidas_saida: int = 0
    caixas_transferidas_entrada: int = 0
    caixas_transferidas_pendentes: int = 0
    kg_programado: float = 0
    kg_carregado: float = 0
    kg_entregue: float = 0
    kg_cancelado: float = 0
    kg_saldo: float = 0
    media_carregada: float = 0
    media_entregue: float = 0
    mortalidade_cliente_aves: int = 0
    mortalidade_cliente_kg: float = 0
    mortalidade_transbordo_aves: int = 0
    mortalidade_transbordo_kg: float = 0
    mortalidade_total_aves: int = 0
    mortalidade_total_kg: float = 0
    valor_previsto: float = 0
    valor_entregue: float = 0
    valor_recebido_app: float = 0
    valor_recebido_manual: float = 0
    valor_recebido_total: float = 0
    gps_entregas: int = 0
    gps_pendentes: int = 0
    distancia_estimativa_km: float = 0


class DespesasBundleResponse(BaseModel):
    cabecalho: DespesasCabecalho
    rota: RotaResumo
    nf: NfResumo
    financeiro: FinanceiroResumo
    operacional: OperacionalResumo = Field(default_factory=OperacionalResumo)
    despesas: list[DespesaItem]
    entregas: list[dict[str, Any]] = Field(default_factory=list)
    fotos: list[dict[str, Any]] = Field(default_factory=list)
    ajudantes_historico: list[dict[str, Any]] = Field(default_factory=list)
    transbordo_foto: dict[str, Any] = Field(default_factory=dict)


class DespesaPayload(BaseModel):
    descricao: str = Field(min_length=1, max_length=220)
    valor: float = Field(gt=0)
    categoria: str | None = Field(default=None, max_length=80)
    tipo_despesa: str | None = Field(default="ROTA", max_length=40)
    observacao: str | None = Field(default=None, max_length=300)
    data_registro: str | None = Field(default=None, max_length=30)


class DespesaPatchPayload(BaseModel):
    descricao: str | None = Field(default=None, max_length=220)
    valor: float | None = Field(default=None, ge=0)
    categoria: str | None = Field(default=None, max_length=80)
    tipo_despesa: str | None = Field(default=None, max_length=40)
    observacao: str | None = Field(default=None, max_length=300)
    data_registro: str | None = Field(default=None, max_length=30)


class RotaPayload(BaseModel):
    km_inicial: float = Field(default=0, ge=0)
    km_final: float = Field(default=0, ge=0)
    litros: float = Field(default=0, ge=0)
    rota_observacao: str | None = Field(default=None, max_length=1000)


class NfPayload(BaseModel):
    nf_numero: str | None = Field(default=None, max_length=80)
    nf_kg: float = Field(default=0, ge=0)
    nf_preco: float = Field(default=0, ge=0)
    nf_caixas: int = Field(default=0, ge=0)
    nf_kg_carregado: float | None = Field(default=None, ge=0)
    nf_kg_vendido: float | None = Field(default=None, ge=0)
    nf_saldo: float | None = Field(default=None, ge=0)
    nf_media_carregada: float = Field(default=0, ge=0)
    nf_caixa_final: int = Field(default=0, ge=0)
    mortalidade_transbordo_aves: int = Field(default=0, ge=0)
    mortalidade_transbordo_kg: float = Field(default=0, ge=0)
    obs_transbordo: str | None = Field(default=None, max_length=500)


class FinanceiroPayload(BaseModel):
    adiantamento: float = Field(default=0, ge=0)
    pix_motorista: float = Field(default=0, ge=0)
    cedulas: dict[str, int] = Field(default_factory=dict)


class MortalidadeManualResponse(BaseModel):
    ok: bool = True
    codigo_programacao: str
    pedido: str = ""
    cliente: str = ""
    mortalidade_aves: int = 0
    mortalidade_kg: float = 0
    valor_desconto: float = 0
    foto: dict[str, Any] = Field(default_factory=dict)


def pdf_clean(value: Any) -> str:
    return (
        str(value or "")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u00a0", " ")
        .strip()
    )


def pdf_money(value: Any) -> str:
    return f"R$ {safe_float(value, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pdf_number(value: Any, places: int = 2) -> str:
    return f"{safe_float(value, 0.0):,.{places}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def safe_file_segment(value: Any, fallback: str = "arquivo") -> str:
    text_value = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip())[:80].strip("._-")
    return text_value or fallback


async def save_mortalidade_manual_photo(
    *,
    codigo: str,
    pedido: str,
    foto: UploadFile | None,
) -> dict[str, Any]:
    if foto is None or not foto.filename:
        return {}
    content = await foto.read()
    if not content:
        return {}
    max_size = 10 * 1024 * 1024
    if len(content) > max_size:
        raise HTTPException(status_code=413, detail="Foto maior que 10MB.")
    suffix = Path(foto.filename).suffix.lower()
    if suffix not in {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}:
        suffix = ".jpg"
    root = Path(".rotahub_runtime") / "fotos_rotas" / "manual_mortalidade"
    root.mkdir(parents=True, exist_ok=True)
    id_foto = f"{safe_file_segment(codigo, 'programacao')}_manual_mortalidade_{uuid4().hex[:12]}"
    file_name = f"{id_foto}{suffix}"
    file_path = root / file_name
    file_path.write_bytes(content)
    return {
        "id_foto": id_foto,
        "arquivo_nome": file_name,
        "path_local": str(file_path),
        "storage_path": str(file_path),
        "mime_type": foto.content_type or "",
        "tamanho_bytes": len(content),
        "pedido": pedido,
        "registrado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def pdf_clip_width(pdf: Any, text: Any, max_width: float, font_name: str = "Helvetica", font_size: int = 9) -> str:
    out = pdf_clean(text)
    if pdf.stringWidth(out, font_name, font_size) <= max_width:
        return out
    ell = "..."
    while out and pdf.stringWidth(out + ell, font_name, font_size) > max_width:
        out = out[:-1]
    return (out + ell) if out else ell


def pdf_wrap_lines(pdf: Any, text: Any, max_width: float, font_name: str = "Helvetica", font_size: int = 8) -> list[str]:
    lines: list[str] = []
    chunks = str(text or "").splitlines() or [""]
    for raw in chunks:
        words = pdf_clean(raw).split()
        if not words:
            lines.append("")
            continue
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if pdf.stringWidth(candidate, font_name, font_size) <= max_width:
                current = candidate
                continue
            if current:
                lines.append(current)
            current = word
            while pdf.stringWidth(current, font_name, font_size) > max_width and len(current) > 1:
                cut = len(current) - 1
                while cut > 1 and pdf.stringWidth(current[:cut] + "-", font_name, font_size) > max_width:
                    cut -= 1
                lines.append(current[:cut] + "-")
                current = current[cut:]
        if current:
            lines.append(current)
    return lines or [""]


def pdf_datetime(value: Any) -> str:
    raw = pdf_clean(value)
    if not raw:
        return ""
    normalized = raw.replace("T", " ").replace("Z", "").strip()
    normalized = re.sub(r"([+-]\d\d:\d\d)$", "", normalized).strip()
    for fmt, size in (
        ("%Y-%m-%d %H:%M:%S", 19),
        ("%Y-%m-%d %H:%M", 16),
        ("%d/%m/%Y %H:%M:%S", 19),
        ("%d/%m/%Y %H:%M", 16),
        ("%d/%m/%y %H:%M:%S", 17),
        ("%d/%m/%y %H:%M", 14),
    ):
        try:
            return datetime.strptime(normalized[:size], fmt).strftime("%d/%m/%y %H:%M:%S")
        except Exception:
            pass
    date_part = normalized.split(" ")[0]
    time_part = normalized.split(" ", 1)[1] if " " in normalized else ""
    return f"{pdf_date(date_part)} {pdf_time(time_part)}".strip() or raw


def pdf_date(value: Any) -> str:
    raw = pdf_clean(value)
    if not raw:
        return ""
    raw = raw.replace("T", " ").split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d/%m/%y")
        except Exception:
            pass
    return raw


def pdf_time(value: Any) -> str:
    raw = pdf_clean(value)
    if not raw:
        return ""
    if "T" in raw or " " in raw:
        raw = raw.replace("T", " ").split(" ", 1)[1]
    raw = raw.split(".")[0].strip()
    raw = re.sub(r"([+-]\d\d:\d\d)$", "", raw).strip()
    parts = raw.split(":")
    if len(parts) >= 3:
        return ":".join(parts[:3])
    if len(parts) == 2:
        return f"{parts[0]}:{parts[1]}:00"
    return raw


def pdf_date_time(date_value: Any, time_value: Any) -> str:
    joined = f"{pdf_date(date_value)} {pdf_time(time_value)}".strip()
    return joined


def qtd_diarias_from_obs(value: Any) -> str:
    text = upper_text(value)
    for pattern in (
        r"QTD\s*DIARIAS?\s*[:=-]\s*([0-9]+)",
        r"\b([0-9]+)\s*DIARIAS?\b",
        r"\b([0-9]+)\s*AJUDANTES?\b",
    ):
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return "-"


def normalize_media_kg(value: Any) -> float:
    media = safe_float(value, 0.0)
    if media > 20:
        media = media / 1000.0
    return media


def draw_despesas_pdf_page(
    pdf: Any,
    bundle: DespesasBundleResponse,
    *,
    programacao: ProgramacaoDB,
    recebimentos: list[RecebimentoDB],
    itens: list[ProgramacaoItemDB],
    controles: list[ProgramacaoItemControleDB],
    transferencias_operacionais: list[dict[str, Any]],
    equipe_txt: str,
    reimpressao: bool = False,
) -> None:
    from reportlab.lib.units import mm

    width, height = pdf._pagesize
    left, right, top, bottom = 12 * mm, 12 * mm, 12 * mm, 12 * mm
    y = height - top

    cab = bundle.cabecalho
    rota = bundle.rota
    nf = bundle.nf
    financeiro = bundle.financeiro
    operacional = bundle.operacional
    codigo = cab.codigo_programacao
    table_x = left
    table_w = width - left - right
    col_total_w = table_w
    col1_w = col_total_w * 0.50
    col_gap = 6 * mm
    col1_x = left
    col2_x = left + col1_w + col_gap
    col2_w = (width - right) - col2_x
    line_h = 5.2 * mm
    reimpressao_label = (
        f"REIMPRESSAO - gerada em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        if reimpressao
        else ""
    )

    motorista = cab.motorista or "-"
    veiculo = cab.veiculo or "-"
    equipe_label = equipe_txt or cab.equipe or "-"
    nf_numero = nf.nf_numero or getattr(programacao, "num_nf", "") or getattr(programacao, "nf_numero", "") or ""
    local_rota = cab.rota or "-"
    local_carregamento = cab.local_carregamento or "-"
    saida = pdf_date_time(cab.data_saida, cab.hora_saida)
    chegada = pdf_date_time(cab.data_chegada, cab.hora_chegada)
    is_transbordo = is_transbordo_programacao(programacao)
    transf_saida = [row for row in transferencias_operacionais if row.get("direcao") == "SAIDA"]
    transf_entrada = [row for row in transferencias_operacionais if row.get("direcao") == "ENTRADA"]
    somente_distribuicao = bool(is_transbordo and transf_saida and not recebimentos)

    def new_page(title: str | None = None) -> None:
        nonlocal y
        pdf.showPage()
        y = height - top
        if title:
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(left, y, title)
            y -= 8 * mm
            if reimpressao_label:
                pdf.setFont("Helvetica-Bold", 8)
                pdf.drawString(left, y, reimpressao_label)
                y -= 5 * mm

    def draw_kv(x: float, y0: float, key: Any, value: Any, col_w: float) -> None:
        key_text = f"{pdf_clean(key)}:"
        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawString(x, y0, key_text)
        pdf.setFont("Helvetica", 9)
        key_w = pdf.stringWidth(key_text, "Helvetica-Bold", 9)
        offset = min(max(22 * mm, key_w + 2.5 * mm), col_w * 0.58)
        avail = max(col_w - offset - 1.5 * mm, 10 * mm)
        pdf.drawString(x + offset, y0, pdf_clip_width(pdf, value, avail, "Helvetica", 9))

    def signature_blocks(anchor_footer: bool = False) -> None:
        nonlocal y
        block_w = (width - left - right - 10 * mm) / 2.0
        block_h = 18 * mm
        gap_x = 10 * mm
        gap_y = 10 * mm
        title_y = bottom + (block_h * 2) + gap_y + (11 * mm) if anchor_footer else y

        def one_block(x: float, y_top: float, title: str) -> None:
            pdf.setLineWidth(0.8)
            pdf.rect(x, y_top - block_h, block_w, block_h, stroke=1, fill=0)
            pdf.setFont("Helvetica-Bold", 9)
            pdf.drawString(x + 3 * mm, y_top - 5 * mm, title)
            pdf.setFont("Helvetica", 9)
            pdf.line(x + 3 * mm, y_top - 13 * mm, x + block_w - 3 * mm, y_top - 13 * mm)
            pdf.drawString(x + 3 * mm, y_top - 16.5 * mm, "Assinatura / Carimbo")

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, title_y, "ASSINATURAS / CONFERENCIA")
        first_row_top = title_y - (10 * mm)
        second_row_top = first_row_top - (block_h + gap_y)
        one_block(left, first_row_top, "SETOR FATURAMENTO")
        one_block(left + block_w + gap_x, first_row_top, "SETOR FINANCEIRO")
        one_block(left, second_row_top, "SETOR DE CAIXA")
        one_block(left + block_w + gap_x, second_row_top, "SETOR DE CONFERENCIA")
        if not anchor_footer:
            y = second_row_top - block_h

    def draw_recebimentos_sheet() -> None:
        nonlocal y
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(left, y, f"FOLHA DE RECEBIMENTOS - PROGRAMACAO {codigo}")
        y -= 8 * mm
        if reimpressao_label:
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(left, y, reimpressao_label)
            y -= 5 * mm

        draw_kv(col1_x, y, "Motorista", motorista, col1_w)
        draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Equipe", equipe_label, col1_w)
        draw_kv(col2_x, y, "NF", nf_numero, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Saida", saida, col1_w)
        draw_kv(col2_x, y, "Chegada", chegada, col2_w)
        y -= 8 * mm

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "RECEBIMENTOS REGISTRADOS")
        y -= 6 * mm

        col_cod = table_w * 0.12
        col_nome = table_w * 0.28
        col_forma = table_w * 0.14
        col_valor = table_w * 0.12
        col_obs = table_w * 0.18
        col_data = table_w - (col_cod + col_nome + col_forma + col_valor + col_obs)
        x_cod = table_x
        x_nome = x_cod + col_cod
        x_forma = x_nome + col_nome
        x_valor = x_forma + col_forma
        x_obs = x_valor + col_valor
        x_data = x_obs + col_obs
        row_h = 6.2 * mm

        def header() -> None:
            nonlocal y
            pdf.setFont("Helvetica-Bold", 8)
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(x_cod + 2, y - row_h + 3, "COD")
            pdf.drawString(x_nome + 2, y - row_h + 3, "CLIENTE")
            pdf.drawString(x_forma + 2, y - row_h + 3, "FORMA")
            pdf.drawRightString(x_valor + col_valor - 2, y - row_h + 3, "VALOR")
            pdf.drawString(x_obs + 2, y - row_h + 3, "OBS")
            pdf.drawRightString(x_data + col_data - 2, y - row_h + 3, "DATA")
            y -= row_h
            pdf.setFont("Helvetica", 8)

        header()
        total_receb = 0.0
        signatures_drawn = False
        if not recebimentos:
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(x_cod + 2, y - row_h + 3, "SEM RECEBIMENTOS REGISTRADOS")
            y -= row_h
        for recebimento in recebimentos:
            if (y - row_h) < bottom + ((70 if not signatures_drawn else 16) * mm):
                if not signatures_drawn:
                    signature_blocks(anchor_footer=True)
                    signatures_drawn = True
                new_page(f"FOLHA DE RECEBIMENTOS - PROGRAMACAO {codigo} (CONT.)")
                header()
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(x_cod + 2, y - row_h + 3, pdf_clip_width(pdf, recebimento.cod_cliente, col_cod - 4, "Helvetica", 8))
            pdf.drawString(x_nome + 2, y - row_h + 3, pdf_clip_width(pdf, recebimento.nome_cliente, col_nome - 4, "Helvetica", 8))
            pdf.drawString(x_forma + 2, y - row_h + 3, pdf_clip_width(pdf, recebimento.forma_pagamento, col_forma - 4, "Helvetica", 8))
            pdf.drawRightString(x_valor + col_valor - 2, y - row_h + 3, pdf_number(recebimento.valor, 2))
            pdf.drawString(x_obs + 2, y - row_h + 3, pdf_clip_width(pdf, recebimento.observacao, col_obs - 4, "Helvetica", 8))
            pdf.drawRightString(x_data + col_data - 2, y - row_h + 3, pdf_clip_width(pdf, pdf_datetime(recebimento.data_registro), col_data - 4, "Helvetica", 7))
            y -= row_h
            total_receb += safe_float(recebimento.valor, 0.0)

        y -= 4 * mm
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawRightString(width - right, y, f"TOTAL RECEBIMENTOS: {pdf_money(total_receb)}")
        y -= 8 * mm
        if not signatures_drawn:
            signature_blocks(anchor_footer=True)

    def draw_assinaturas_sheet() -> None:
        nonlocal y
        new_page(f"ASSINATURAS / CONFERENCIA - PROGRAMACAO {codigo}")
        pdf.setFont("Helvetica", 9)
        pdf.drawString(left, y, "Prestacao sem folha de recebimentos: operacao de transbordo/distribuicao de carga.")
        y -= 10 * mm
        signature_blocks(anchor_footer=False)

    def draw_retorno_sheet(start_new: bool = True) -> None:
        nonlocal y
        control_by_key: dict[tuple[str, str], ProgramacaoItemControleDB] = {}
        control_by_cod: dict[str, ProgramacaoItemControleDB] = {}
        for controle in controles:
            cod = upper_text(controle.cod_cliente)
            pedido = upper_text(controle.pedido)
            control_by_key[(cod, pedido)] = controle
            control_by_cod.setdefault(cod, controle)

        status_entregue = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLU\u00cdDO"}
        status_cancelado = {"CANCELADO", "CANCELADA"}
        media = normalize_media_kg(getattr(programacao, "media", 0) or nf.nf_media_carregada)
        if media <= 0:
            media = normalize_media_kg(nf.nf_media_carregada)
        aves_cx = safe_int(nf.nf_caixa_final, 0) or safe_int(getattr(programacao, "aves_caixa_final", 0), 0) or 6
        kg_carregado = safe_float(nf.nf_kg_carregado, 0.0) or safe_float(nf.nf_kg, 0.0)
        caixas_carregadas = safe_int(nf.nf_caixas, 0) or sum(max(safe_int(item.qnt_caixas, 0), 0) for item in itens)

        linhas: list[tuple[str, str, str, str, str, str, str, str, str, str]] = []
        ocorrencias: list[tuple[str, str, str, str, str, str, str]] = []
        tot_entregues = 0
        tot_cancelados = 0
        tot_cx_ent = 0
        tot_kg_ent = 0.0
        tot_valor = 0.0
        tot_mort_aves = safe_int(nf.mortalidade_transbordo_aves, 0)
        tot_mort_kg = safe_float(nf.mortalidade_transbordo_kg, 0.0)
        divergencias = 0

        for item in itens:
            controle = control_by_key.get((upper_text(item.cod_cliente), upper_text(item.pedido))) or control_by_cod.get(upper_text(item.cod_cliente))
            status_item = upper_text((controle.status_pedido if controle else "") or item.status_pedido or "PENDENTE") or "PENDENTE"
            cx_prog = max(safe_int(item.qnt_caixas, 0), 0)
            cx_alt = max(safe_int((controle.caixas_atual if controle else None) or item.caixas_atual, 0), 0)
            cx_ent = 0 if status_item in status_cancelado else (cx_alt if cx_alt > 0 else (cx_prog if status_item in status_entregue else 0))
            preco_orig = safe_float(item.preco, 0.0)
            preco_final = safe_float((controle.preco_atual if controle else None) or item.preco_atual, 0.0) or preco_orig
            kg_orig = safe_float(item.kg, 0.0)
            kg_base = safe_float((controle.peso_previsto if controle else None), 0.0) or kg_orig
            if kg_base <= 0 and media > 0:
                kg_base = float(cx_ent if cx_ent > 0 else cx_prog) * float(max(aves_cx, 1)) * media
            aves_total = max((cx_ent if cx_ent > 0 else cx_prog) * max(aves_cx, 1), 0)
            media_cliente = (kg_base / aves_total) if aves_total > 0 and kg_base > 0 else media
            mort_aves = max(safe_int((controle.mortalidade_aves if controle else None), 0), 0)
            mort_kg = mort_aves * media_cliente if media_cliente > 0 else 0.0
            kg_ent = 0.0 if status_item in status_cancelado else kg_base
            valor_total = max(kg_ent * preco_final, 0.0)
            delta_kg = kg_ent - kg_orig if kg_orig > 0 else 0.0
            alterado = bool(
                (controle and (controle.alteracao_tipo or controle.alteracao_detalhe))
                or (cx_alt > 0 and cx_alt != cx_prog)
                or (preco_final > 0 and abs(preco_final - preco_orig) >= 0.001)
            )

            if status_item in status_entregue:
                tot_entregues += 1
            if status_item in status_cancelado:
                tot_cancelados += 1
            if abs(delta_kg) >= 0.01:
                divergencias += 1
            tot_cx_ent += cx_ent
            tot_kg_ent += kg_ent
            tot_valor += valor_total
            tot_mort_aves += mort_aves
            tot_mort_kg += mort_kg

            linhas.append(
                (
                    pdf_clean(item.cod_cliente)[:6],
                    upper_text(item.nome_cliente)[:24],
                    status_item[:10],
                    f"{cx_prog}/{cx_ent}",
                    pdf_number(media_cliente, 3),
                    pdf_number(kg_ent, 2),
                    pdf_number(preco_final, 2),
                    pdf_number(valor_total, 2),
                    f"{mort_aves}/{pdf_number(mort_kg, 2)}",
                    (f"{delta_kg:+.2f}".replace(".", ",") if abs(delta_kg) >= 0.01 else "-"),
                )
            )
            if status_item in status_cancelado or alterado:
                ocorrencias.append(
                    (
                        upper_text(item.nome_cliente)[:20],
                        status_item[:10],
                        "SIM" if status_item in status_cancelado else "NAO",
                        "SIM" if alterado else "NAO",
                        upper_text((controle.alteracao_tipo if controle else "") or "")[:14],
                        pdf_clean((controle.alteracao_detalhe if controle else "") or item.observacao or "")[:28],
                        upper_text((controle.alterado_por if controle else "") or item.alterado_por or "")[:14],
                    )
                )

        if start_new:
            pdf.showPage()
            y = height - top
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(left, y, f"FOLHA DE RETORNO OPERACIONAL - PROGRAMACAO {codigo}")
        y -= 8 * mm
        if reimpressao_label:
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(left, y, reimpressao_label)
            y -= 5 * mm
        draw_kv(col1_x, y, "Motorista", motorista, col1_w)
        draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Equipe", equipe_label, col1_w)
        draw_kv(col2_x, y, "NF", nf_numero, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Local Rota", local_rota, col1_w)
        draw_kv(col2_x, y, "Carregou em", local_carregamento, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Saida", saida, col1_w)
        draw_kv(col2_x, y, "Chegada", chegada, col2_w)
        y -= 6 * mm

        pdf.setFont("Helvetica", 8)
        saldo_nf = safe_float(nf.nf_saldo, 0.0)
        if abs(saldo_nf) < 0.0001 and (operacional.kg_carregado > 0 or operacional.kg_entregue > 0):
            saldo_nf = operacional.kg_saldo
        pdf.drawString(
            left,
            y,
            pdf_clip_width(
                pdf,
                f"KG NF {pdf_number(nf.nf_kg, 2)} | KG carregado {pdf_number(operacional.kg_carregado or kg_carregado, 2)} | "
                f"Saldo NF {pdf_number(saldo_nf, 2)} | Aves/CX {aves_cx} | CX final {nf.nf_caixa_final}",
                table_w,
                "Helvetica",
                8,
            ),
        )
        y -= 4.2 * mm
        medias = [
            safe_float(getattr(programacao, "media_1", 0), 0.0),
            safe_float(getattr(programacao, "media_2", 0), 0.0),
            safe_float(getattr(programacao, "media_3", 0), 0.0),
        ]
        medias_txt = " / ".join(pdf_number(normalize_media_kg(v), 3) for v in medias if v > 0) or pdf_number(media, 3)
        pdf.drawString(left, y, pdf_clip_width(pdf, f"Medias lancadas: {medias_txt}", table_w, "Helvetica", 8))
        y -= 7 * mm

        box_gap = 4 * mm
        box_w = (table_w - (box_gap * 3)) / 4.0
        box_h = 13 * mm

        def summary_box(x: float, y_top: float, title: str, value: str) -> None:
            pdf.rect(x, y_top - box_h, box_w, box_h, stroke=1, fill=0)
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(x + 2 * mm, y_top - 4.5 * mm, title)
            pdf.setFont("Helvetica", 11)
            pdf.drawRightString(x + box_w - 2 * mm, y_top - 9.5 * mm, pdf_clip_width(pdf, value, box_w - 4 * mm, "Helvetica", 11))

        summary_box(left, y, "Clientes", str(operacional.clientes_total or len(itens)))
        summary_box(left + box_w + box_gap, y, "Entregues/Cancel.", f"{operacional.pedidos_entregues}/{operacional.pedidos_cancelados}")
        summary_box(left + ((box_w + box_gap) * 2), y, "CX Carreg./Entreg.", f"{operacional.caixas_carregadas or caixas_carregadas}/{operacional.caixas_entregues or tot_cx_ent}")
        summary_box(left + ((box_w + box_gap) * 3), y, "KG Carreg./Entreg.", f"{pdf_number(operacional.kg_carregado or kg_carregado, 2)}/{pdf_number(operacional.kg_entregue or tot_kg_ent, 2)}")
        y -= (box_h + 6 * mm)
        summary_box(left, y, "Media Oper.", pdf_number(operacional.media_entregue or media, 3))
        summary_box(left + box_w + box_gap, y, "Ocorrencias", f"{operacional.mortalidade_total_aves or tot_mort_aves} / {pdf_number(operacional.mortalidade_total_kg or tot_mort_kg, 2)}kg")
        summary_box(left + ((box_w + box_gap) * 2), y, "Valor Entregue", pdf_money(operacional.valor_entregue or tot_valor))
        summary_box(left + ((box_w + box_gap) * 3), y, "Alterados", str(operacional.pedidos_alterados or divergencias))
        y -= (box_h + 7 * mm)

        if transferencias_operacionais:
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(left, y, "TRANSBORDO / TRANSFERENCIAS DE CARGA")
            y -= 5.5 * mm
            pdf.setFont("Helvetica", 8)
            if is_transbordo:
                modal = upper_text(getattr(programacao, "transbordo_modalidade", "")) or "TRANSBORDO"
                obs_trans = pdf_clean(getattr(programacao, "transbordo_observacao", "") or "")
                pdf.drawString(
                    left,
                    y,
                    pdf_clip_width(
                        pdf,
                        f"Operacao: {modal} | Raiz da carga preservada para fechamento web/app motorista. {obs_trans}",
                        table_w,
                        "Helvetica",
                        8,
                    ),
                )
                y -= 4.5 * mm

            row_h_trans = 6.0 * mm
            trans_ws = [table_w * v for v in [0.09, 0.11, 0.14, 0.12, 0.07, 0.07, 0.07, 0.09, 0.09, 0.07]]
            trans_ws.append(table_w - sum(trans_ws))
            trans_xs = [table_x]
            for col_w in trans_ws[:-1]:
                trans_xs.append(trans_xs[-1] + col_w)

            def header_transbordo() -> None:
                nonlocal y
                pdf.setFont("Helvetica-Bold", 7)
                pdf.rect(table_x, y - row_h_trans + 1, table_w, row_h_trans, stroke=1, fill=0)
                for idx, title in enumerate(["TIPO", "CODIGO", "MOTORISTA", "VEICULO", "CX", "CONV", "SALDO", "KG", "KG CONV", "MEDIA", "STATUS"]):
                    if idx in {4, 5, 6, 7, 8, 9}:
                        pdf.drawRightString(trans_xs[idx] + trans_ws[idx] - 2, y - row_h_trans + 3, title)
                    else:
                        pdf.drawString(trans_xs[idx] + 2, y - row_h_trans + 3, title)
                y -= row_h_trans
                pdf.setFont("Helvetica", 7)

            header_transbordo()
            total_saida_cx = total_entrada_cx = 0
            total_saida_kg = total_entrada_kg = 0.0
            for row in transferencias_operacionais:
                if y < bottom + 42 * mm:
                    new_page(f"FOLHA DE RETORNO OPERACIONAL - PROGRAMACAO {codigo} (CONT.)")
                    header_transbordo()
                direcao = row.get("direcao") or "-"
                caixas = safe_int(row.get("caixas"), 0)
                caixas_convertidas = safe_int(row.get("caixas_convertidas"), 0)
                caixas_saldo = safe_int(row.get("caixas_saldo"), 0)
                kg = safe_float(row.get("kg"), 0.0)
                kg_convertido = safe_float(row.get("kg_convertido"), 0.0)
                if direcao == "SAIDA":
                    total_saida_cx += caixas_convertidas
                    total_saida_kg += kg_convertido
                elif direcao == "ENTRADA":
                    total_entrada_cx += caixas_convertidas
                    total_entrada_kg += kg_convertido
                values = [
                    "ENVIO" if direcao == "SAIDA" else "RECEB.",
                    row.get("codigo_vinculado") or "-",
                    row.get("motorista") or "-",
                    row.get("veiculo") or "-",
                    str(caixas),
                    str(caixas_convertidas),
                    str(caixas_saldo),
                    pdf_number(kg, 2),
                    pdf_number(kg_convertido, 2),
                    pdf_number(row.get("media"), 3),
                    row.get("status") or "-",
                ]
                pdf.rect(table_x, y - row_h_trans + 1, table_w, row_h_trans, stroke=1, fill=0)
                for idx, value in enumerate(values):
                    text = pdf_clip_width(pdf, value, trans_ws[idx] - 4, "Helvetica", 7)
                    if idx in {4, 5, 6, 7, 8, 9}:
                        pdf.drawRightString(trans_xs[idx] + trans_ws[idx] - 2, y - row_h_trans + 3, text)
                    else:
                        pdf.drawString(trans_xs[idx] + 2, y - row_h_trans + 3, text)
                y -= row_h_trans
            y -= 3.5 * mm
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(
                left,
                y,
                pdf_clip_width(
                    pdf,
                    f"Total recebido convertido: {total_entrada_cx} cx / {pdf_number(total_entrada_kg, 2)} kg | "
                    f"Total distribuido convertido: {total_saida_cx} cx / {pdf_number(total_saida_kg, 2)} kg | "
                    f"Divergencia kg entregue/carregado: {pdf_number((operacional.kg_entregue or tot_kg_ent) - (operacional.kg_carregado or kg_carregado), 2)}",
                    table_w,
                    "Helvetica-Bold",
                    8,
                ),
            )
            y -= 7 * mm

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "ENTREGAS POR CLIENTE")
        y -= 5.5 * mm

        row_h = 5.8 * mm
        widths = [0.08, 0.21, 0.11, 0.08, 0.08, 0.10, 0.10, 0.12, 0.08]
        col_ws = [table_w * v for v in widths]
        col_ws.append(table_w - sum(col_ws))
        xs = [table_x]
        for col_w in col_ws[:-1]:
            xs.append(xs[-1] + col_w)

        def header_entregas() -> None:
            nonlocal y
            headers = ["COD", "CLIENTE", "STATUS", "CX", "MEDIA", "KG", "PRECO", "VALOR", "MORT", "DIV"]
            pdf.setFont("Helvetica-Bold", 7)
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, header in enumerate(headers):
                if idx >= 3:
                    pdf.drawRightString(xs[idx] + col_ws[idx] - 2, y - row_h + 3, header)
                else:
                    pdf.drawString(xs[idx] + 2, y - row_h + 3, header)
            y -= row_h
            pdf.setFont("Helvetica", 7)

        header_entregas()
        if not linhas:
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(table_x + 2, y - row_h + 3, "SEM ENTREGAS REGISTRADAS")
            y -= row_h
        for row in linhas:
            if y < bottom + 58 * mm:
                new_page(f"FOLHA DE RETORNO OPERACIONAL - PROGRAMACAO {codigo} (CONT.)")
                header_entregas()
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, value in enumerate(row):
                text = pdf_clip_width(pdf, value, col_ws[idx] - 4, "Helvetica", 7)
                if idx >= 3:
                    pdf.drawRightString(xs[idx] + col_ws[idx] - 2, y - row_h + 3, text)
                else:
                    pdf.drawString(xs[idx] + 2, y - row_h + 3, text)
            y -= row_h

        y -= 6 * mm
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "OCORRENCIAS / AJUSTES")
        y -= 5.5 * mm
        occ_ws = [table_w * v for v in [0.21, 0.10, 0.06, 0.06, 0.15, 0.26]]
        occ_ws.append(table_w - sum(occ_ws))
        occ_xs = [table_x]
        for col_w in occ_ws[:-1]:
            occ_xs.append(occ_xs[-1] + col_w)

        def header_ocorrencias() -> None:
            nonlocal y
            pdf.setFont("Helvetica-Bold", 7)
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, header in enumerate(["CLIENTE", "STATUS", "CANC", "ALT", "PARA", "MOTIVO", "AUTORIZOU"]):
                pdf.drawString(occ_xs[idx] + 2, y - row_h + 3, header)
            y -= row_h
            pdf.setFont("Helvetica", 7)

        header_ocorrencias()
        if not ocorrencias:
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(table_x + 2, y - row_h + 3, "SEM OCORRENCIAS DE CANCELAMENTO/ALTERACAO")
            y -= row_h
        for row in ocorrencias:
            if y < bottom + 20 * mm:
                new_page(f"OCORRENCIAS - RETORNO OPERACIONAL {codigo} (CONT.)")
                header_ocorrencias()
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, value in enumerate(row):
                text = pdf_clip_width(pdf, value, occ_ws[idx] - 4, "Helvetica", 7)
                if idx in {2, 3}:
                    pdf.drawCentredString(occ_xs[idx] + (occ_ws[idx] / 2.0), y - row_h + 3, text)
                else:
                    pdf.drawString(occ_xs[idx] + 2, y - row_h + 3, text)
            y -= row_h

    def draw_mortalidade_sheet() -> None:
        nonlocal y
        fotos_mort = [
            foto for foto in (bundle.fotos or [])
            if "MORTALIDADE" in upper_text(foto.get("categoria") or foto.get("tipo_registro"))
            or "DOA" in upper_text(foto.get("categoria") or foto.get("tipo_registro"))
        ]
        controles_mort = [controle for controle in controles if safe_int(controle.mortalidade_aves, 0) > 0]
        total_cliente_aves = sum(max(safe_int(controle.mortalidade_aves, 0), 0) for controle in controles_mort)

        media_ref = normalize_media_kg(getattr(programacao, "media", 0) or nf.nf_media_carregada)
        if media_ref <= 0:
            medias_validas = [
                normalize_media_kg(value)
                for value in (
                    getattr(programacao, "media_1", 0),
                    getattr(programacao, "media_2", 0),
                    getattr(programacao, "media_3", 0),
                )
                if safe_float(value, 0.0) > 0
            ]
            media_ref = sum(medias_validas) / len(medias_validas) if medias_validas else 0.0

        total_cliente_kg = 0.0
        linhas_cliente: list[tuple[str, str, str, str, str, str, str, str]] = []
        itens_por_chave = {(upper_text(item.cod_cliente), upper_text(item.pedido)): item for item in itens}
        itens_por_cod: dict[str, ProgramacaoItemDB] = {}
        for item in itens:
            itens_por_cod.setdefault(upper_text(item.cod_cliente), item)

        for controle in controles_mort:
            cod = upper_text(controle.cod_cliente)
            pedido = upper_text(controle.pedido)
            item = itens_por_chave.get((cod, pedido)) or itens_por_cod.get(cod)
            cx_base = max(safe_int((controle.caixas_atual if controle.caixas_atual is not None else None) or (item.qnt_caixas if item else 0), 0), 0)
            aves_cx = safe_int(nf.nf_caixa_final, 0) or safe_int(getattr(programacao, "aves_caixa_final", 0), 0) or 6
            peso_previsto = safe_float(controle.peso_previsto, 0.0)
            if peso_previsto <= 0 and item:
                peso_previsto = safe_float(item.kg, 0.0)
            aves_base = max(cx_base * max(aves_cx, 1), 0)
            media_cliente = (peso_previsto / aves_base) if peso_previsto > 0 and aves_base > 0 else media_ref
            mort_aves = max(safe_int(controle.mortalidade_aves, 0), 0)
            mort_kg = max(mort_aves * media_cliente, 0.0)
            total_cliente_kg += mort_kg
            foto_ref = parse_json_dict(getattr(controle, "foto_mortalidade_ref_json", ""))
            foto_label = (
                foto_ref.get("arquivo_nome")
                or foto_ref.get("storage_path")
                or getattr(controle, "foto_mortalidade_path", "")
                or getattr(controle, "mortalidade_foto_path", "")
                or "-"
            )
            linhas_cliente.append(
                (
                    cod[:8],
                    upper_text(getattr(item, "nome_cliente", "") if item else "")[:22],
                    pedido[:12],
                    upper_text(controle.status_pedido)[:10],
                    str(mort_aves),
                    pdf_number(mort_kg, 2),
                    pdf_number(media_cliente, 3),
                    pdf_clean(foto_label)[:24],
                )
            )

        doa_aves = safe_int(nf.mortalidade_transbordo_aves, 0)
        doa_kg = safe_float(nf.mortalidade_transbordo_kg, 0.0)
        total_aves = total_cliente_aves + doa_aves
        total_kg = total_cliente_kg + doa_kg
        kg_nf = safe_float(nf.nf_kg, 0.0)
        kg_carregado = safe_float(nf.nf_kg_carregado, 0.0) or safe_float(nf.nf_kg, 0.0)
        kg_util_nf = max(kg_nf - doa_kg, 0.0) if kg_nf > 0 else 0.0
        kg_util_rota = max(kg_carregado - total_kg, 0.0) if kg_carregado > 0 else 0.0

        new_page(f"OCORRENCIAS / IMPACTO NA CARGA - PLANEJAMENTO {codigo}")
        draw_kv(col1_x, y, "Motorista", motorista, col1_w)
        draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "NF", nf_numero, col1_w)
        draw_kv(col2_x, y, "Local rota", local_rota, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Media carregamento", pdf_number(media_ref, 3), col1_w)
        draw_kv(col2_x, y, "Fotos ocorrencias", str(len(fotos_mort)), col2_w)
        y -= 8 * mm

        box_gap = 4 * mm
        box_w = (table_w - (box_gap * 3)) / 4.0
        box_h = 13 * mm

        def box(x: float, y_top: float, title: str, value: str) -> None:
            pdf.rect(x, y_top - box_h, box_w, box_h, stroke=1, fill=0)
            pdf.setFont("Helvetica-Bold", 8)
            pdf.drawString(x + 2 * mm, y_top - 4.5 * mm, title)
            pdf.setFont("Helvetica", 10)
            pdf.drawRightString(x + box_w - 2 * mm, y_top - 9.5 * mm, pdf_clip_width(pdf, value, box_w - 4 * mm, "Helvetica", 10))

        box(left, y, "Operacao / Transbordo", f"{doa_aves} unid. / {pdf_number(doa_kg, 2)} kg")
        box(left + box_w + box_gap, y, "Ocorr. cliente", f"{total_cliente_aves} unid. / {pdf_number(total_cliente_kg, 2)} kg")
        box(left + ((box_w + box_gap) * 2), y, "Total ocorr.", f"{total_aves} unid. / {pdf_number(total_kg, 2)} kg")
        box(left + ((box_w + box_gap) * 3), y, "KG util rota", pdf_number(kg_util_rota, 2))
        y -= box_h + 6 * mm

        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawString(left, y, "EFEITO NO FECHAMENTO DA CARGA")
        y -= 5 * mm
        pdf.setFont("Helvetica", 8)
        efeitos = [
            f"KG NF: {pdf_number(kg_nf, 2)} | KG carregado: {pdf_number(kg_carregado, 2)}",
            f"Operacao/transbordo reduz NF util para: {pdf_number(kg_util_nf, 2)} kg",
            f"Ocorrencias totais apuradas na rota: {total_aves} unid. / {pdf_number(total_kg, 2)} kg",
            f"Obs transbordo: {pdf_clean(nf.obs_transbordo or '-')}",
        ]
        for line in efeitos:
            pdf.drawString(left, y, pdf_clip_width(pdf, line, table_w, "Helvetica", 8))
            y -= 4.5 * mm
        y -= 4 * mm

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "OCORRENCIAS POR CLIENTE")
        y -= 5.5 * mm
        row_h = 6.0 * mm
        col_ws = [table_w * v for v in [0.09, 0.24, 0.12, 0.10, 0.08, 0.10, 0.10]]
        col_ws.append(table_w - sum(col_ws))
        xs = [table_x]
        for col_w in col_ws[:-1]:
            xs.append(xs[-1] + col_w)

        def header_cliente() -> None:
            nonlocal y
            pdf.setFont("Helvetica-Bold", 7)
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, header in enumerate(["COD", "CLIENTE", "PEDIDO", "STATUS", "AVES", "KG", "MEDIA", "FOTO"]):
                pdf.drawString(xs[idx] + 2, y - row_h + 3, header)
            y -= row_h
            pdf.setFont("Helvetica", 7)

        header_cliente()
        if not linhas_cliente:
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(table_x + 2, y - row_h + 3, "SEM MORTALIDADE POR CLIENTE REGISTRADA")
            y -= row_h
        for row in linhas_cliente:
            if y < bottom + 20 * mm:
                new_page(f"MORTALIDADE POR CLIENTE - PROGRAMACAO {codigo} (CONT.)")
                header_cliente()
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            for idx, value in enumerate(row):
                pdf.drawString(xs[idx] + 2, y - row_h + 3, pdf_clip_width(pdf, value, col_ws[idx] - 4, "Helvetica", 7))
            y -= row_h

    def draw_despesas_sheet() -> None:
        nonlocal y
        new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {codigo}")

        def section(title: str) -> None:
            nonlocal y
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(left, y, title)
            y -= 6 * mm
            pdf.setFont("Helvetica", 9)

        section("DADOS DA ROTA")
        draw_kv(col1_x, y, "Motorista", motorista, col1_w)
        draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Equipe", equipe_label, col1_w)
        draw_kv(col2_x, y, "NF", nf_numero, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Local da Rota", local_rota, col1_w)
        draw_kv(col2_x, y, "Local Carregamento", local_carregamento, col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Saida", saida, col1_w)
        draw_kv(col2_x, y, "Chegada", chegada, col2_w)
        y -= 8 * mm

        section("CARREGAMENTO / NOTA FISCAL")
        draw_kv(col1_x, y, "NF KG", pdf_number(nf.nf_kg, 2), col1_w)
        draw_kv(col2_x, y, "NF Caixas", str(nf.nf_caixas), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "KG Carregado", pdf_number(nf.nf_kg_carregado, 2), col1_w)
        draw_kv(col2_x, y, "KG Vendido", pdf_number(nf.nf_kg_vendido, 2), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Saldo (KG)", pdf_number(nf.nf_saldo, 2), col1_w)
        draw_kv(col2_x, y, "Preco NF", pdf_money(nf.nf_preco), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Media carregada", pdf_number(normalize_media_kg(getattr(programacao, "media", 0) or nf.nf_media_carregada), 3), col1_w)
        draw_kv(col2_x, y, "Caixa final", str(nf.nf_caixa_final), col2_w)
        y -= 8 * mm

        section("CONSOLIDADO OPERACIONAL APP MOTORISTA")
        draw_kv(col1_x, y, "Pedidos", f"E {operacional.pedidos_entregues} / P {operacional.pedidos_pendentes} / C {operacional.pedidos_cancelados}", col1_w)
        draw_kv(col2_x, y, "Alterados", str(operacional.pedidos_alterados), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Caixas", f"Carreg {operacional.caixas_carregadas} / Entreg {operacional.caixas_entregues}", col1_w)
        draw_kv(col2_x, y, "Transferencias", f"Saiu {operacional.caixas_transferidas_saida} / Entrou {operacional.caixas_transferidas_entrada}", col2_w)
        y -= line_h
        draw_kv(col1_x, y, "KG real", f"Carreg {pdf_number(operacional.kg_carregado, 2)} / Entreg {pdf_number(operacional.kg_entregue, 2)}", col1_w)
        draw_kv(col2_x, y, "Saldo carga", pdf_number(operacional.kg_saldo, 2), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Medias", f"Carreg {pdf_number(operacional.media_carregada, 3)} / Entreg {pdf_number(operacional.media_entregue, 3)}", col1_w)
        draw_kv(col2_x, y, "GPS", f"{operacional.gps_entregas} com localizacao / {operacional.gps_pendentes} pendentes", col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Ocorrencias", f"Cliente {operacional.mortalidade_cliente_aves} unid. / Total {pdf_number(operacional.mortalidade_total_kg, 2)} kg", col1_w)
        draw_kv(col2_x, y, "Recebido app/manual", f"{pdf_money(operacional.valor_recebido_app)} / {pdf_money(operacional.valor_recebido_manual)}", col2_w)
        y -= 8 * mm

        section("DADOS DE ROTA (KM)")
        draw_kv(col1_x, y, "KM Inicial", pdf_number(rota.km_inicial, 2), col1_w)
        draw_kv(col2_x, y, "KM Final", pdf_number(rota.km_final, 2), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Litros", pdf_number(rota.litros, 2), col1_w)
        draw_kv(col2_x, y, "KM Rodado", pdf_number(rota.km_rodado, 2), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Media", pdf_number(rota.media_km_l, 2), col1_w)
        draw_kv(col2_x, y, "Custo KM", pdf_number(rota.custo_km, 2), col2_w)
        y -= 8 * mm

        section("CONTAGEM DE CEDULAS")
        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawString(left, y, "CEDULA")
        pdf.drawString(left + 25 * mm, y, "QTD")
        pdf.drawString(left + 45 * mm, y, "TOTAL")
        y -= 5 * mm
        pdf.setFont("Helvetica", 9)
        ced_total = 0.0
        for ced in CEDULAS:
            qtd = safe_int(financeiro.cedulas.get(str(ced), 0), 0)
            total_ced = ced * qtd
            ced_total += total_ced
            pdf.drawString(left, y, f"R$ {pdf_number(ced, 2)}")
            pdf.drawString(left + 25 * mm, y, str(qtd))
            pdf.drawString(left + 45 * mm, y, pdf_money(total_ced))
            y -= 5 * mm
        y -= 2 * mm
        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawString(left, y, f"TOTAL DINHEIRO: {pdf_money(financeiro.valor_dinheiro or ced_total)}")
        y -= 5 * mm
        pdf.drawString(left, y, f"PIX MOTORISTA: {pdf_money(financeiro.pix_motorista)}")
        y -= 5 * mm
        pdf.drawString(left, y, f"TOTAL DEVOLVIDO: {pdf_money(financeiro.total_devolvido)}")
        y -= 8 * mm

        section("DESPESAS")
        col_desc = table_w * 0.25
        col_cat = table_w * 0.13
        col_val = table_w * 0.12
        col_obs = table_w * 0.32
        col_data = table_w - (col_desc + col_cat + col_val + col_obs)
        x_desc = table_x
        x_cat = x_desc + col_desc
        x_val = x_cat + col_cat
        x_obs = x_val + col_val
        x_data = x_obs + col_obs
        row_h = 6.5 * mm
        row_line_h = 4.2 * mm

        def despesas_header() -> None:
            nonlocal y
            pdf.setFont("Helvetica-Bold", 8)
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(x_desc + 2, y - row_h + 3, "DESCRICAO")
            pdf.drawString(x_cat + 2, y - row_h + 3, "CATEGORIA")
            pdf.drawRightString(x_val + col_val - 2, y - row_h + 3, "VALOR")
            pdf.drawString(x_obs + 2, y - row_h + 3, "OBS")
            pdf.drawRightString(x_data + col_data - 2, y - row_h + 3, "DATA")
            y -= row_h
            pdf.setFont("Helvetica", 8)

        despesas_header()
        if not bundle.despesas:
            pdf.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            pdf.drawString(x_desc + 2, y - row_h + 3, "SEM DESPESAS REGISTRADAS")
            y -= row_h
        for idx, despesa in enumerate(bundle.despesas):
            desc = despesa.descricao
            obs = despesa.observacao
            desc_up = upper_text(desc)
            if desc_up in {"DIARIAS MOTORISTA", "DIARIA MOTORISTA"}:
                obs = f"QTD DIARIAS: {qtd_diarias_from_obs(obs)} | MOTORISTA: {motorista}"
            elif desc_up in {"DIARIAS AJUDANTES", "DIARIA AJUDANTE", "DIARIA AJUDANTES"}:
                obs = f"QTD DIARIAS: {qtd_diarias_from_obs(obs)} | AJUDANTES: {equipe_label}"
            desc_lines = pdf_wrap_lines(pdf, desc, col_desc - 4, "Helvetica", 8)
            cat_lines = pdf_wrap_lines(pdf, despesa.categoria, col_cat - 4, "Helvetica", 8)
            obs_lines = pdf_wrap_lines(pdf, obs, col_obs - 4, "Helvetica", 8)
            data_lines = pdf_wrap_lines(pdf, pdf_datetime(despesa.data_registro), col_data - 4, "Helvetica", 8)
            line_count = max(len(desc_lines), len(cat_lines), len(obs_lines), len(data_lines), 1)
            row_h_curr = max(row_h, (line_count * row_line_h) + 4)
            reserva_mm = 48 if idx == (len(bundle.despesas) - 1) else 24
            if y < bottom + row_h_curr + (reserva_mm * mm):
                new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {codigo} (CONT.)")
                pdf.setFont("Helvetica-Bold", 10)
                pdf.drawString(left, y, "DESPESAS")
                y -= 6 * mm
                despesas_header()
            pdf.rect(table_x, y - row_h_curr + 1, table_w, row_h_curr, stroke=1, fill=0)
            text_y = y - 10
            for line_idx, line in enumerate(desc_lines):
                pdf.drawString(x_desc + 2, text_y - (line_idx * row_line_h), line)
            for line_idx, line in enumerate(cat_lines):
                pdf.drawString(x_cat + 2, text_y - (line_idx * row_line_h), line)
            pdf.drawRightString(x_val + col_val - 2, text_y, pdf_number(despesa.valor, 2))
            for line_idx, line in enumerate(obs_lines):
                pdf.drawString(x_obs + 2, text_y - (line_idx * row_line_h), line)
            for line_idx, line in enumerate(data_lines):
                pdf.drawRightString(x_data + col_data - 2, text_y - (line_idx * row_line_h), line)
            y -= row_h_curr

        y -= 6 * mm
        if y < bottom + (48 * mm):
            new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {codigo} (CONT.)")
        section("RESUMO FINANCEIRO (CAIXA)")
        draw_kv(col1_x, y, "Recebimentos", pdf_money(financeiro.total_recebido), col1_w)
        draw_kv(col2_x, y, "Adiantamento", pdf_money(financeiro.adiantamento), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Despesas", pdf_money(financeiro.total_despesas), col1_w)
        draw_kv(col2_x, y, "Cedulas", pdf_money(ced_total), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "PIX Motorista", pdf_money(financeiro.pix_motorista), col1_w)
        draw_kv(col2_x, y, "Total Devolvido", pdf_money(financeiro.total_devolvido), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Total Entradas", pdf_money(financeiro.total_entradas), col1_w)
        draw_kv(col2_x, y, "Total Saidas", pdf_money(financeiro.total_saidas), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Caixa", pdf_money(financeiro.valor_final_caixa), col1_w)
        draw_kv(col2_x, y, "Diferenca", pdf_money(financeiro.diferenca), col2_w)
        y -= line_h
        draw_kv(col1_x, y, "Resultado", pdf_money(financeiro.resultado_liquido), col1_w)
        pdf.setFont("Helvetica", 8)
        pdf.drawRightString(width - right, bottom - 2 * mm, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    if somente_distribuicao:
        draw_retorno_sheet(start_new=False)
        draw_mortalidade_sheet()
        draw_despesas_sheet()
        draw_assinaturas_sheet()
    else:
        draw_recebimentos_sheet()
        draw_retorno_sheet()
        draw_mortalidade_sheet()
        draw_despesas_sheet()


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
        raise HTTPException(status_code=409, detail="Programacao cancelada nao aceita despesas.")


def assert_can_open_prestacao(programacao: ProgramacaoDB) -> None:
    if status_ref(programacao) in CANCELLED_STATUSES:
        raise HTTPException(status_code=409, detail="Programacao cancelada nao aceita prestacao de contas.")


def money(value: Any) -> float:
    return round(safe_float(value, 0.0), 2)


def one_decimal(value: Any) -> float:
    return round(safe_float(value, 0.0), 1)


def two_decimal(value: Any) -> float:
    return round(safe_float(value, 0.0), 2)


def three_decimal(value: Any) -> float:
    return round(safe_float(value, 0.0), 3)


def cedulas_from_programacao(programacao: ProgramacaoDB) -> dict[str, int]:
    return {str(ced): max(safe_int(getattr(programacao, f"ced_{ced}_qtd", 0), 0), 0) for ced in CEDULAS}


def cedulas_total(cedulas: dict[str, int]) -> float:
    return money(sum(float(ced) * safe_int(qtd, 0) for ced, qtd in cedulas.items()))


def despesa_to_response(item: DespesaDB) -> DespesaItem:
    try:
        foto = json.loads(str(item.foto_despesa_ref_json or "{}"))
        if not isinstance(foto, dict):
            foto = {}
    except Exception:
        foto = {}
    return DespesaItem(
        id=item.id,
        codigo_programacao=upper_text(item.codigo_programacao),
        descricao=upper_text(item.descricao),
        valor=money(item.valor),
        data_registro=str(item.data_registro or "")[:19],
        tipo_despesa=upper_text(item.tipo_despesa or "ROTA"),
        categoria=upper_text(item.categoria),
        motorista=upper_text(item.motorista),
        veiculo=upper_text(item.veiculo),
        observacao=str(item.observacao or ""),
        id_local=str(item.id_local or ""),
        forma_pagamento=upper_text(item.forma_pagamento),
        comprovante_path=str(item.comprovante_path or ""),
        estabelecimento=upper_text(item.estabelecimento),
        documento=str(item.documento or ""),
        litros=safe_float(item.litros, 0),
        valor_litro=money(item.valor_litro),
        desconto=money(item.desconto),
        combustivel=upper_text(item.combustivel),
        odometro=safe_float(item.odometro, 0),
        lat=item.lat,
        lon=item.lon,
        accuracy=item.accuracy,
        origem=upper_text(item.origem),
        foto=foto,
    )


async def sum_recebimentos(db: AsyncSession, codigo: str) -> float:
    result = await db.execute(
        select(func.coalesce(func.sum(RecebimentoDB.valor), 0)).where(func.upper(RecebimentoDB.codigo_programacao) == codigo)
    )
    return money(result.scalar() or 0)


async def sum_despesas(db: AsyncSession, codigo: str) -> float:
    result = await db.execute(
        select(func.coalesce(func.sum(DespesaDB.valor), 0)).where(func.upper(DespesaDB.codigo_programacao) == codigo)
    )
    return money(result.scalar() or 0)


async def despesa_rows(db: AsyncSession, codigo: str) -> list[DespesaDB]:
    result = await db.execute(
        select(DespesaDB)
        .where(func.upper(DespesaDB.codigo_programacao) == codigo)
        .order_by(DespesaDB.id.desc())
    )
    return list(result.scalars().all())


async def sync_diarias_despesas(db: AsyncSession, programacao: ProgramacaoDB) -> bool:
    codigo = upper_text(programacao.codigo_programacao)
    nomes_por_id = await recebimentos_ajudante_map(db)
    diarias = await diarias_for_programacao(db, programacao, nomes_por_id)
    result = await db.execute(select(DespesaDB).where(func.upper(DespesaDB.codigo_programacao) == codigo))
    despesas = list(result.scalars().all())
    existentes = [
        item
        for item in despesas
        if upper_text(item.categoria) == "DIARIAS"
        and upper_text(item.descricao) in DIARIA_DESCRICOES_AUTO
    ]

    changed = False
    if safe_float(diarias.qtd_diarias, 0.0) <= 0:
        for item in existentes:
            await db.delete(item)
            changed = True
        if changed:
            await db.flush()
        return changed

    now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    desired = {
        DIARIA_DESCRICAO_MOTORISTA: {
            "valor": money(diarias.total_motorista),
            "observacao": diarias.observacao_motorista,
            "id_local": "AUTO_DIARIA_MOTORISTA",
        },
        DIARIA_DESCRICAO_AJUDANTES: {
            "valor": money(diarias.total_ajudantes),
            "observacao": diarias.observacao_ajudantes,
            "id_local": "AUTO_DIARIA_AJUDANTES",
        },
    }
    by_desc: dict[str, list[DespesaDB]] = {desc: [] for desc in desired}
    for item in existentes:
        by_desc.setdefault(upper_text(item.descricao), []).append(item)

    for descricao, meta in desired.items():
        rows = by_desc.get(descricao) or []
        target = rows[0] if rows else None
        is_new = target is None
        if target is None:
            target = DespesaDB(codigo_programacao=codigo)
            db.add(target)
            changed = True

        updates = {
            "descricao": descricao,
            "valor": meta["valor"],
            "categoria": "DIARIAS",
            "tipo_despesa": "DIARIAS",
            "motorista": upper_text(programacao.motorista),
            "veiculo": upper_text(programacao.veiculo),
            "observacao": meta["observacao"],
            "id_local": meta["id_local"],
            "forma_pagamento": "PAGO",
            "origem": "RECEBIMENTOS_WEB",
        }
        if is_new or not target.data_registro:
            updates["data_registro"] = now_s
        for field, value in updates.items():
            if getattr(target, field, None) != value:
                setattr(target, field, value)
                changed = True

        for duplicate in rows[1:]:
            await db.delete(duplicate)
            changed = True

    if changed:
        await db.flush()
    return changed


async def recebimento_rows(db: AsyncSession, codigo: str) -> list[RecebimentoDB]:
    result = await db.execute(
        select(RecebimentoDB)
        .where(func.upper(RecebimentoDB.codigo_programacao) == codigo)
        .order_by(RecebimentoDB.data_registro.desc(), RecebimentoDB.id.desc())
    )
    return list(result.scalars().all())


async def item_rows(db: AsyncSession, codigo: str) -> list[ProgramacaoItemDB]:
    result = await db.execute(
        select(ProgramacaoItemDB)
        .where(func.upper(ProgramacaoItemDB.codigo_programacao) == codigo)
        .order_by(ProgramacaoItemDB.nome_cliente.asc(), ProgramacaoItemDB.cod_cliente.asc())
    )
    return list(result.scalars().all())


async def controle_rows(db: AsyncSession, codigo: str) -> list[ProgramacaoItemControleDB]:
    result = await db.execute(
        select(ProgramacaoItemControleDB)
        .where(func.upper(ProgramacaoItemControleDB.codigo_programacao) == codigo)
        .order_by(ProgramacaoItemControleDB.updated_at.desc(), ProgramacaoItemControleDB.id.desc())
    )
    return list(result.scalars().all())


def parse_json_dict(raw: Any) -> dict[str, Any]:
    try:
        data = json.loads(str(raw or "{}"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def parse_json_list(raw: Any) -> list[dict[str, Any]]:
    try:
        data = json.loads(str(raw or "[]"))
        if isinstance(data, list):
            return [item for item in data if isinstance(item, dict)]
    except Exception:
        pass
    return []


async def mobile_fotos_rows(db: AsyncSession, codigo: str) -> list[dict[str, Any]]:
    result = await db.execute(
        text(
            """
            SELECT id_foto, codigo_programacao, categoria, tipo_registro, cod_cliente, cliente_nome, pedido,
                   id_vinculo, path_local, storage_path, arquivo_nome, mime_type, tamanho_bytes,
                   motorista_codigo, motorista_nome, registrado_em, payload_json
              FROM rota_fotos
             WHERE UPPER(codigo_programacao)=:codigo
             ORDER BY id DESC
            """
        ),
        {"codigo": codigo},
    )
    out = []
    for row in result.mappings().all():
        item = dict(row)
        item["payload"] = parse_json_dict(item.pop("payload_json", ""))
        out.append(item)
    return out


async def entrega_rows(db: AsyncSession, codigo: str) -> list[dict[str, Any]]:
    result = await db.execute(
        text(
            """
            SELECT pc.cod_cliente, pc.pedido, pc.status_pedido, pc.mortalidade_aves,
                   pc.media_aplicada, pc.valor_recebido, pc.forma_recebimento,
                   pc.peso_previsto, pc.caixas_atual, pc.preco_atual,
                   pc.alteracao_tipo, pc.alteracao_detalhe, pc.alterado_por,
                   pc.lat_evento, pc.lon_evento,
                   COALESCE(pc.lat_entrega, pc.lat_evento) AS latitude,
                   COALESCE(pc.lon_entrega, pc.lon_evento) AS longitude,
                   pc.accuracy_entrega, pc.timestamp_entrega,
                   pc.endereco_evento, pc.cidade_evento, pc.bairro_evento,
                   pc.foto_mortalidade_path, pc.mortalidade_foto_path, pc.foto_mortalidade_ref_json,
                   pc.updated_at, pi.nome_cliente
              FROM programacao_itens_controle pc
              LEFT JOIN programacao_itens pi
                ON UPPER(pi.codigo_programacao)=UPPER(pc.codigo_programacao)
               AND UPPER(pi.cod_cliente)=UPPER(pc.cod_cliente)
               AND COALESCE(pi.pedido, '')=COALESCE(pc.pedido, '')
             WHERE UPPER(pc.codigo_programacao)=:codigo
             ORDER BY COALESCE(pc.timestamp_entrega, pc.updated_at, '') DESC, pc.id DESC
            """
        ),
        {"codigo": codigo},
    )
    out = []
    for row in result.mappings().all():
        item = dict(row)
        item["foto"] = parse_json_dict(item.pop("foto_mortalidade_ref_json", ""))
        out.append(item)
    return out


async def resolve_equipe_pdf(db: AsyncSession, equipe_raw: str | None) -> str:
    raw = str(equipe_raw or "").strip()
    if not raw:
        return ""
    result = await db.execute(select(AjudanteDB))
    nomes = {
        str(item.id): upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
        for item in result.scalars().all()
    }
    out = []
    seen = set()
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


async def itens_totais(db: AsyncSession, codigo: str) -> tuple[float, float, int]:
    result = await db.execute(select(ProgramacaoItemDB).where(func.upper(ProgramacaoItemDB.codigo_programacao) == codigo))
    kg_total = 0.0
    receita_total = 0.0
    caixas_total = 0
    for item in result.scalars().all():
        kg = safe_float(item.kg, 0.0)
        preco = safe_float(item.preco, 0.0)
        kg_total += kg
        receita_total += kg * preco
        caixas_total += safe_int(item.qnt_caixas, 0)
    return two_decimal(kg_total), money(receita_total), caixas_total


def build_rota(programacao: ProgramacaoDB, total_despesas: float) -> RotaResumo:
    km_inicial = safe_float(programacao.km_inicial, 0.0)
    km_final = safe_float(programacao.km_final, 0.0)
    litros = safe_float(programacao.litros, 0.0)
    km_rodado = safe_float(programacao.km_rodado, 0.0)
    if km_rodado <= 0 and km_final >= km_inicial:
        km_rodado = km_final - km_inicial
    media_km_l = (km_rodado / litros) if litros > 0 else safe_float(programacao.media_km_l, 0.0)
    custo_km = (total_despesas / km_rodado) if km_rodado > 0 else safe_float(programacao.custo_km, 0.0)
    return RotaResumo(
        km_inicial=one_decimal(km_inicial),
        km_final=one_decimal(km_final),
        litros=two_decimal(litros),
        km_rodado=one_decimal(max(km_rodado, 0.0)),
        media_km_l=two_decimal(media_km_l),
        custo_km=two_decimal(custo_km),
        rota_observacao=str(programacao.rota_observacao or ""),
    )


def aves_por_caixa_ref(programacao: ProgramacaoDB) -> int:
    return (
        safe_int(getattr(programacao, "qnt_aves_caixa_final", 0), 0)
        or safe_int(getattr(programacao, "aves_caixa_final", 0), 0)
        or safe_int(getattr(programacao, "qnt_aves_por_cx", 0), 0)
        or 6
    )


def media_carregada_ref(programacao: ProgramacaoDB) -> float:
    media = normalize_media_kg(
        safe_float(getattr(programacao, "media", 0), 0.0)
        or safe_float(getattr(programacao, "nf_media_carregada", 0), 0.0)
    )
    if media > 0:
        return media
    medias = [
        normalize_media_kg(getattr(programacao, field, 0))
        for field in ("media_1", "media_2", "media_3")
        if safe_float(getattr(programacao, field, 0), 0.0) > 0
    ]
    return (sum(medias) / len(medias)) if medias else 0.0


async def transferencias_totais(db: AsyncSession, codigo: str) -> tuple[int, int, int]:
    try:
        result = await db.execute(
            text(
                """
                SELECT codigo_origem, codigo_destino, qtd_caixas, qtd_convertida, status
                  FROM transferencias
                 WHERE UPPER(COALESCE(codigo_origem, ''))=:codigo
                    OR UPPER(COALESCE(codigo_destino, ''))=:codigo
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        return 0, 0, 0
    saida = 0
    entrada = 0
    pendente = 0
    for row in result.mappings().all():
        status_value = upper_text(row.get("status"))
        if status_value in CANCELLED_STATUSES:
            continue
        qtd = transferencia_qtd_total(row)
        qtd_convertida = transferencia_qtd_convertida(row)
        qtd_saldo = max(qtd - qtd_convertida, 0)
        if not qtd:
            continue
        is_pending = status_value in {"", "PENDENTE", "ABERTA", "AGUARDANDO", "SOLICITADA"}
        if is_pending:
            pendente += qtd
        if upper_text(row.get("codigo_origem")) == codigo:
            saida += qtd
        if upper_text(row.get("codigo_destino")) == codigo:
            entrada += qtd
    return saida, entrada, pendente


def transferencia_qtd_total(row: Any) -> int:
    return max(safe_int(row.get("qtd_caixas"), 0), safe_int(row.get("qtd_convertida"), 0), 0)


def transferencia_qtd_convertida(row: Any) -> int:
    qtd_convertida = max(safe_int(row.get("qtd_convertida"), 0), 0)
    if qtd_convertida > 0:
        return qtd_convertida
    if upper_text(row.get("status")) == "CONVERTIDA":
        return transferencia_qtd_total(row)
    return 0


async def transferencias_compra_totais(db: AsyncSession, codigo: str) -> tuple[float, float]:
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
                 WHERE UPPER(COALESCE(t.codigo_origem, ''))=:codigo
                    OR UPPER(COALESCE(t.codigo_destino, ''))=:codigo
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        return 0.0, 0.0
    saida_valor = 0.0
    entrada_valor = 0.0
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
    for row in rows:
        status_value = upper_text(row.get("status"))
        if status_value in CANCELLED_STATUSES:
            continue
        qtd = transferencia_qtd_total(row)
        qtd_convertida = transferencia_qtd_convertida(row)
        qtd_saldo = max(qtd - qtd_convertida, 0)
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
        if upper_text(row.get("codigo_origem")) == codigo:
            saida_valor += valor
        if upper_text(row.get("codigo_destino")) == codigo:
            entrada_valor += valor
    return money(saida_valor), money(entrada_valor)


async def transferencias_operacionais_pdf(db: AsyncSession, codigo: str) -> list[dict[str, Any]]:
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
                    t.criado_em AS created_at,
                    origem.motorista AS motorista_origem,
                    origem.veiculo AS veiculo_origem,
                    origem.nf_kg AS origem_nf_kg,
                    origem.kg_nf AS origem_kg_nf,
                    origem.nf_kg_carregado AS origem_nf_kg_carregado,
                    origem.kg_carregado AS origem_kg_carregado,
                    origem.nf_caixas AS origem_nf_caixas,
                    origem.total_caixas AS origem_total_caixas,
                    origem.caixas_carregadas AS origem_caixas_carregadas,
                    origem.media AS origem_media,
                    origem.media AS origem_nf_media,
                    destino.motorista AS motorista_destino,
                    destino.veiculo AS veiculo_destino
                  FROM transferencias t
                  LEFT JOIN programacoes origem
                    ON UPPER(COALESCE(origem.codigo_programacao, ''))=UPPER(COALESCE(t.codigo_origem, ''))
                  LEFT JOIN programacoes destino
                    ON UPPER(COALESCE(destino.codigo_programacao, ''))=UPPER(COALESCE(t.codigo_destino, ''))
                 WHERE UPPER(COALESCE(t.codigo_origem, ''))=:codigo
                    OR UPPER(COALESCE(t.codigo_destino, ''))=:codigo
                 ORDER BY COALESCE(t.criado_em, '') ASC, t.id ASC
                """
            ),
            {"codigo": codigo},
        )
    except Exception:
        return []

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

    out: list[dict[str, Any]] = []
    for row in rows:
        status_value = upper_text(row.get("status"))
        if status_value in CANCELLED_STATUSES:
            continue
        origem = upper_text(row.get("codigo_origem"))
        destino = upper_text(row.get("codigo_destino"))
        direcao = "SAIDA" if origem == codigo else "ENTRADA"
        qtd = transferencia_qtd_total(row)
        qtd_convertida = transferencia_qtd_convertida(row)
        qtd_saldo = max(qtd - qtd_convertida, 0)
        if qtd <= 0:
            continue
        raiz = carga_raiz_from_snapshot(row.get("snapshot"), origem)
        root = root_map.get(raiz)
        kg_base = (
            safe_float(getattr(root, "nf_kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "kg_carregado", 0), 0.0)
            or safe_float(getattr(root, "nf_kg", 0), 0.0)
            or safe_float(getattr(root, "kg_nf", 0), 0.0)
            or safe_float(row.get("origem_nf_kg_carregado"), 0.0)
            or safe_float(row.get("origem_kg_carregado"), 0.0)
            or safe_float(row.get("origem_nf_kg"), 0.0)
            or safe_float(row.get("origem_kg_nf"), 0.0)
        )
        caixas_base = (
            safe_int(getattr(root, "nf_caixas", 0), 0)
            or safe_int(getattr(root, "total_caixas", 0), 0)
            or safe_int(getattr(root, "caixas_carregadas", 0), 0)
            or safe_int(row.get("origem_nf_caixas"), 0)
            or safe_int(row.get("origem_total_caixas"), 0)
            or safe_int(row.get("origem_caixas_carregadas"), 0)
        )
        media = (
            media_carregada_ref(root)
            if root
            else normalize_media_kg(row.get("origem_media") or row.get("origem_nf_media"))
        )
        kg_por_caixa = (kg_base / caixas_base) if kg_base > 0 and caixas_base > 0 else 0.0
        kg_transferido = two_decimal(qtd * kg_por_caixa) if kg_por_caixa > 0 else 0.0
        kg_convertido = two_decimal(qtd_convertida * kg_por_caixa) if kg_por_caixa > 0 else 0.0
        out.append(
            {
                "direcao": direcao,
                "codigo_vinculado": destino if direcao == "SAIDA" else origem,
                "motorista": row.get("motorista_destino") if direcao == "SAIDA" else row.get("motorista_origem"),
                "veiculo": row.get("veiculo_destino") if direcao == "SAIDA" else row.get("veiculo_origem"),
                "caixas": qtd,
                "caixas_convertidas": qtd_convertida,
                "caixas_saldo": qtd_saldo,
                "kg": kg_transferido,
                "kg_convertido": kg_convertido,
                "media": media,
                "carga_raiz": raiz,
                "status": status_value or "-",
                "created_at": row.get("created_at"),
            }
        )
    return out


async def build_operacional(
    db: AsyncSession,
    programacao: ProgramacaoDB,
    itens: list[ProgramacaoItemDB],
    controles: list[ProgramacaoItemControleDB],
    recebimentos: list[RecebimentoDB],
) -> OperacionalResumo:
    codigo = upper_text(programacao.codigo_programacao)
    control_by_key: dict[tuple[str, str], ProgramacaoItemControleDB] = {}
    control_by_cod: dict[str, ProgramacaoItemControleDB] = {}
    for controle in controles:
        cod = upper_text(controle.cod_cliente)
        pedido = upper_text(controle.pedido)
        if (cod, pedido) not in control_by_key:
            control_by_key[(cod, pedido)] = controle
        control_by_cod.setdefault(cod, controle)

    delivered_statuses = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
    cancelled_statuses = CANCELLED_STATUSES
    media_base = media_carregada_ref(programacao)
    aves_cx = max(aves_por_caixa_ref(programacao), 1)

    caixas_programadas = 0
    caixas_entregues = 0
    caixas_canceladas = 0
    kg_programado = 0.0
    kg_entregue = 0.0
    kg_cancelado = 0.0
    valor_previsto = 0.0
    valor_entregue = 0.0
    pedidos_entregues = 0
    pedidos_cancelados = 0
    pedidos_alterados = 0
    gps_entregas = 0
    gps_pendentes = 0
    distancia_estimativa = 0.0
    mortalidade_cliente_aves = 0
    mortalidade_cliente_kg = 0.0

    for item in itens:
        cod = upper_text(item.cod_cliente)
        pedido = upper_text(item.pedido)
        controle = control_by_key.get((cod, pedido)) or control_by_cod.get(cod)
        status_item = upper_text((controle.status_pedido if controle else "") or item.status_pedido or "PENDENTE") or "PENDENTE"
        is_delivered = status_item in delivered_statuses
        is_cancelled = status_item in cancelled_statuses

        cx_prog = max(safe_int(item.qnt_caixas, 0), 0)
        cx_atual = max(safe_int((controle.caixas_atual if controle else None) or item.caixas_atual, 0), 0)
        cx_final = 0 if is_cancelled else (cx_atual if cx_atual > 0 else (cx_prog if is_delivered else 0))
        preco_original = safe_float(item.preco, 0.0)
        preco_final = safe_float((controle.preco_atual if controle else None) or item.preco_atual, 0.0) or preco_original
        kg_orig = safe_float(item.kg, 0.0)
        if kg_orig <= 0 and media_base > 0:
            kg_orig = cx_prog * aves_cx * media_base

        media_cliente = normalize_media_kg(controle.media_aplicada if controle else 0)
        if media_cliente <= 0 and cx_prog > 0 and kg_orig > 0:
            media_cliente = kg_orig / float(cx_prog * aves_cx)
        if media_cliente <= 0:
            media_cliente = media_base

        peso_app = safe_float(controle.peso_previsto if controle else 0, 0.0)
        if peso_app > 0:
            kg_final = peso_app
        elif media_cliente > 0 and cx_final > 0:
            kg_final = cx_final * aves_cx * media_cliente
        elif kg_orig > 0 and cx_prog > 0:
            kg_final = kg_orig * (cx_final / cx_prog)
        else:
            kg_final = 0.0
        if is_cancelled:
            kg_final = 0.0

        mort_aves = max(safe_int(controle.mortalidade_aves if controle else 0, 0), 0)
        mort_kg = mort_aves * media_cliente if media_cliente > 0 else 0.0
        lat = safe_float(getattr(controle, "lat_entrega", 0), 0.0) or safe_float(getattr(controle, "lat_evento", 0), 0.0) if controle else 0.0
        lon = safe_float(getattr(controle, "lon_entrega", 0), 0.0) or safe_float(getattr(controle, "lon_evento", 0), 0.0) if controle else 0.0
        has_gps = abs(lat) > 0.000001 and abs(lon) > 0.000001
        alterado = bool(
            (controle and (controle.alteracao_tipo or controle.alteracao_detalhe or controle.media_aplicada))
            or (cx_atual > 0 and cx_atual != cx_prog)
            or (preco_final > 0 and abs(preco_final - preco_original) >= 0.001)
            or is_cancelled
        )

        caixas_programadas += cx_prog
        caixas_entregues += cx_final
        caixas_canceladas += cx_prog if is_cancelled else 0
        kg_programado += kg_orig
        kg_entregue += kg_final
        kg_cancelado += kg_orig if is_cancelled else 0.0
        valor_previsto += kg_orig * preco_original
        valor_entregue += kg_final * preco_final
        mortalidade_cliente_aves += mort_aves
        mortalidade_cliente_kg += mort_kg
        pedidos_entregues += 1 if is_delivered else 0
        pedidos_cancelados += 1 if is_cancelled else 0
        pedidos_alterados += 1 if alterado else 0
        gps_entregas += 1 if has_gps else 0
        gps_pendentes += 0 if has_gps else 1
        distancia_estimativa += safe_float(getattr(item, "distancia", 0), 0.0) or safe_float(getattr(controle, "distancia", 0), 0.0) if controle else 0.0

    valor_recebido_app = money(sum(max(safe_float(controle.valor_recebido, 0.0), 0.0) for controle in controles))
    valor_recebido_manual = money(sum(max(safe_float(item.valor, 0.0), 0.0) for item in recebimentos))
    mort_trans_aves = max(safe_int(programacao.mortalidade_transbordo_aves, 0), 0)
    mort_trans_kg = max(safe_float(programacao.mortalidade_transbordo_kg, 0.0), 0.0)
    kg_carregado = (
        safe_float(programacao.nf_kg_carregado, 0.0)
        or safe_float(programacao.kg_carregado, 0.0)
        or safe_float(programacao.nf_kg, 0.0)
        or kg_programado
    )
    caixas_carregadas = (
        safe_int(programacao.nf_caixas, 0)
        or safe_int(programacao.total_caixas, 0)
        or safe_int(programacao.caixas_carregadas, 0)
        or caixas_programadas
    )
    total_mort_aves = mortalidade_cliente_aves + mort_trans_aves
    total_mort_kg = mortalidade_cliente_kg + mort_trans_kg
    saida, entrada, pendente = await transferencias_totais(db, codigo)
    media_entregue = (kg_entregue / (caixas_entregues * aves_cx)) if caixas_entregues > 0 and aves_cx > 0 else 0.0
    pedidos_pendentes = max(len(itens) - pedidos_entregues - pedidos_cancelados, 0)
    kg_saldo = max(kg_carregado + (entrada * aves_cx * media_base) - kg_entregue - total_mort_kg - (saida * aves_cx * media_base), 0.0)

    return OperacionalResumo(
        clientes_total=len(itens),
        pedidos_entregues=pedidos_entregues,
        pedidos_pendentes=pedidos_pendentes,
        pedidos_cancelados=pedidos_cancelados,
        pedidos_alterados=pedidos_alterados,
        caixas_programadas=caixas_programadas,
        caixas_carregadas=caixas_carregadas,
        caixas_entregues=caixas_entregues,
        caixas_canceladas=caixas_canceladas,
        caixas_transferidas_saida=saida,
        caixas_transferidas_entrada=entrada,
        caixas_transferidas_pendentes=pendente,
        kg_programado=two_decimal(kg_programado),
        kg_carregado=two_decimal(kg_carregado),
        kg_entregue=two_decimal(kg_entregue),
        kg_cancelado=two_decimal(kg_cancelado),
        kg_saldo=two_decimal(kg_saldo),
        media_carregada=three_decimal(media_base),
        media_entregue=three_decimal(media_entregue),
        mortalidade_cliente_aves=mortalidade_cliente_aves,
        mortalidade_cliente_kg=two_decimal(mortalidade_cliente_kg),
        mortalidade_transbordo_aves=mort_trans_aves,
        mortalidade_transbordo_kg=two_decimal(mort_trans_kg),
        mortalidade_total_aves=total_mort_aves,
        mortalidade_total_kg=two_decimal(total_mort_kg),
        valor_previsto=money(valor_previsto),
        valor_entregue=money(valor_entregue),
        valor_recebido_app=valor_recebido_app,
        valor_recebido_manual=valor_recebido_manual,
        valor_recebido_total=money(valor_recebido_app + valor_recebido_manual),
        gps_entregas=gps_entregas,
        gps_pendentes=gps_pendentes,
        distancia_estimativa_km=one_decimal(distancia_estimativa),
    )


def build_financeiro(programacao: ProgramacaoDB, total_recebido: float, total_despesas: float) -> FinanceiroResumo:
    cedulas = cedulas_from_programacao(programacao)
    valor_dinheiro = cedulas_total(cedulas)
    if valor_dinheiro <= 0:
        valor_dinheiro = money(programacao.valor_dinheiro)
    adiantamento = money(safe_float(programacao.adiantamento, 0.0) or safe_float(programacao.adiantamento_rota, 0.0))
    pix_motorista = money(programacao.pix_motorista)
    total_entradas = money(total_recebido + adiantamento)
    total_devolvido = money(valor_dinheiro + pix_motorista)
    total_saidas = money(total_despesas + total_devolvido)
    valor_final_caixa = money(total_entradas - total_despesas)
    diferenca = money(valor_final_caixa - total_devolvido)
    resultado_liquido = money(total_entradas - total_saidas)
    return FinanceiroResumo(
        total_recebido=total_recebido,
        total_despesas=total_despesas,
        adiantamento=adiantamento,
        valor_dinheiro=valor_dinheiro,
        pix_motorista=pix_motorista,
        total_entradas=total_entradas,
        total_saidas=total_saidas,
        valor_final_caixa=valor_final_caixa,
        total_devolvido=total_devolvido,
        diferenca=diferenca,
        resultado_liquido=resultado_liquido,
        cedulas=cedulas,
    )


def normalize_operacao_tipo(value: Any, tipo_estimativa: Any = "") -> str:
    raw = upper_text(value).replace("-", "_").replace(" ", "_")
    if raw in {"TRANSBORDO", "TRANSFERENCIA_CARGA", "REDISTRIBUICAO"}:
        return "TRANSBORDO"
    return "TRANSBORDO" if upper_text(tipo_estimativa) == "CX" else "VENDA"


def is_transbordo_programacao(programacao: ProgramacaoDB) -> bool:
    return normalize_operacao_tipo(getattr(programacao, "operacao_tipo", ""), getattr(programacao, "tipo_estimativa", "")) == "TRANSBORDO"


def carga_raiz_from_snapshot(value: Any, fallback: Any = "") -> str:
    data = parse_json_dict(value)
    return upper_text(data.get("carga_raiz_programacao") or data.get("carga_origem_programacao") or fallback)


def build_nf(
    programacao: ProgramacaoDB,
    kg_itens: float,
    receita_itens: float,
    caixas_itens: int,
    total_despesas: float,
    operacional: OperacionalResumo | None = None,
    transferencia_saida_valor: float = 0.0,
    transferencia_entrada_valor: float = 0.0,
) -> NfResumo:
    nf_kg = two_decimal(programacao.nf_kg or programacao.kg_carregado or 0)
    nf_preco = money(programacao.nf_preco)
    nf_caixas = safe_int(programacao.nf_caixas, 0) or safe_int(programacao.total_caixas, 0) or (
        operacional.caixas_carregadas if operacional else 0
    ) or caixas_itens
    kg_carregado = two_decimal(programacao.nf_kg_carregado or (operacional.kg_carregado if operacional else 0) or nf_kg)
    kg_vendido = two_decimal(programacao.nf_kg_vendido or (operacional.kg_entregue if operacional else 0) or kg_itens)
    saldo_nf_real = max(nf_kg - kg_carregado, 0.0) if nf_kg > 0 and kg_carregado > 0 else 0.0
    saldo = two_decimal(programacao.nf_saldo if safe_float(programacao.nf_saldo, 0.0) > 0 else saldo_nf_real)
    desconto_fornecedor = money(saldo * nf_preco)
    receita_ref = (operacional.valor_entregue if operacional else 0.0) or receita_itens
    kg_ref = (operacional.kg_entregue if operacional else 0.0) or kg_itens
    preco_medio_venda = money((receita_ref / kg_ref) if kg_ref > 0 else 0.0)
    total_compra_bruta = money(nf_kg * nf_preco)
    total_compra_liquida = money(max(total_compra_bruta - desconto_fornecedor, 0.0))
    if total_compra_liquida <= 0 and kg_carregado > 0 and nf_preco > 0:
        total_compra_liquida = money(kg_carregado * nf_preco)
    total_compra_liquida = money(max(total_compra_liquida - transferencia_saida_valor, 0.0) + transferencia_entrada_valor)
    receita_estimada = money(receita_ref if receita_ref > 0 else (kg_vendido * preco_medio_venda))
    lucro_bruto = money(receita_estimada - total_compra_liquida)
    lucro_liquido = money(lucro_bruto - total_despesas)
    margem = two_decimal((lucro_liquido / receita_estimada * 100.0) if receita_estimada > 0 else 0.0)
    transbordo_sem_venda = is_transbordo_programacao(programacao) and receita_ref <= 0 and kg_vendido <= 0
    if transbordo_sem_venda:
        desconto_fornecedor = 0.0
        total_compra_bruta = 0.0
        total_compra_liquida = 0.0
        receita_estimada = 0.0
        lucro_bruto = 0.0
        lucro_liquido = 0.0
        margem = 0.0
    mort_kg = two_decimal(programacao.mortalidade_transbordo_kg)
    return NfResumo(
        nf_numero=upper_text(programacao.nf_numero or programacao.num_nf),
        nf_kg=nf_kg,
        nf_preco=nf_preco,
        nf_caixas=nf_caixas,
        nf_kg_carregado=kg_carregado,
        nf_kg_vendido=kg_vendido,
        nf_saldo=saldo,
        nf_media_carregada=two_decimal((operacional.media_carregada if operacional else 0) or programacao.media),
        nf_caixa_final=safe_int(programacao.qnt_aves_caixa_final, 0) or safe_int(programacao.aves_caixa_final, 0),
        mortalidade_transbordo_aves=safe_int(programacao.mortalidade_transbordo_aves, 0),
        mortalidade_transbordo_kg=mort_kg,
        obs_transbordo=str(programacao.obs_transbordo or ""),
        kg_nf_util=two_decimal(max(nf_kg - mort_kg, 0.0)),
        nf_saldo_valor=desconto_fornecedor,
        desconto_fornecedor=desconto_fornecedor,
        total_compra_bruta=total_compra_bruta,
        total_compra_liquida=total_compra_liquida,
        preco_medio_venda=preco_medio_venda,
        total_compra=total_compra_bruta,
        receita_estimada=receita_estimada,
        despesas_rota=total_despesas,
        lucro_bruto=lucro_bruto,
        lucro_liquido=lucro_liquido,
        margem_liquida=margem,
    )


async def serialize_bundle(db: AsyncSession, programacao: ProgramacaoDB) -> DespesasBundleResponse:
    codigo = upper_text(programacao.codigo_programacao)
    total_recebido = await sum_recebimentos(db, codigo)
    total_despesas = await sum_despesas(db, codigo)
    kg_itens, receita_itens, caixas_itens = await itens_totais(db, codigo)
    itens = await item_rows(db, codigo)
    controles = await controle_rows(db, codigo)
    recebimentos = await recebimento_rows(db, codigo)
    operacional = await build_operacional(db, programacao, itens, controles, recebimentos)
    transferencia_saida_valor, transferencia_entrada_valor = await transferencias_compra_totais(db, codigo)
    despesas = [despesa_to_response(item) for item in await despesa_rows(db, codigo)]
    fotos = await mobile_fotos_rows(db, codigo)
    entregas = await entrega_rows(db, codigo)
    ajudantes_historico = parse_json_list(programacao.historico_ajudantes)
    transbordo_foto = parse_json_dict(programacao.foto_doa_ref_json)
    return DespesasBundleResponse(
        cabecalho=DespesasCabecalho(
            codigo_programacao=codigo,
            motorista=upper_text(programacao.motorista),
            veiculo=upper_text(programacao.veiculo),
            equipe=programacao.equipe or "",
            rota=upper_text(programacao.local_rota or programacao.tipo_rota),
            local_carregamento=upper_text(
                programacao.local_carregamento
                or programacao.granja_carregada
                or programacao.local_carregado
                or programacao.local_carreg
            ),
            data_saida=str(programacao.saida_data or programacao.data_saida or ""),
            hora_saida=str(programacao.saida_hora or programacao.hora_saida or ""),
            data_chegada=str(programacao.data_chegada or ""),
            hora_chegada=str(programacao.hora_chegada or ""),
            status=status_ref(programacao),
            prestacao_status=upper_text(programacao.prestacao_status or "PENDENTE"),
            adiantamento_origem=upper_text(programacao.adiantamento_origem),
            operacao_tipo=normalize_operacao_tipo(getattr(programacao, "operacao_tipo", ""), getattr(programacao, "tipo_estimativa", "")),
            transbordo_modalidade=upper_text(getattr(programacao, "transbordo_modalidade", "") or ""),
            transbordo_observacao=upper_text(getattr(programacao, "transbordo_observacao", "") or ""),
            fechada=is_closed(programacao),
        ),
        rota=build_rota(programacao, total_despesas),
        nf=build_nf(
            programacao,
            kg_itens,
            receita_itens,
            caixas_itens,
            total_despesas,
            operacional,
            transferencia_saida_valor,
            transferencia_entrada_valor,
        ),
        financeiro=build_financeiro(programacao, operacional.valor_recebido_total or total_recebido, total_despesas),
        operacional=operacional,
        despesas=despesas,
        entregas=entregas,
        fotos=fotos,
        ajudantes_historico=ajudantes_historico,
        transbordo_foto=transbordo_foto,
    )


async def refresh_km_metrics(db: AsyncSession, programacao: ProgramacaoDB) -> None:
    total_despesas = await sum_despesas(db, upper_text(programacao.codigo_programacao))
    km_inicial = safe_float(programacao.km_inicial, 0.0)
    km_final = safe_float(programacao.km_final, 0.0)
    litros = safe_float(programacao.litros, 0.0)
    km_rodado = km_final - km_inicial
    if km_rodado < 0:
        km_rodado = 0.0
    programacao.km_rodado = km_rodado
    programacao.media_km_l = (km_rodado / litros) if litros > 0 else 0.0
    programacao.custo_km = (total_despesas / km_rodado) if km_rodado > 0 else 0.0


@router.get("/programacoes", response_model=list[DespesaProgramacaoOption])
async def listar_programacoes_despesas(
    limit: int = 300,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    result = await db.execute(select(ProgramacaoDB).order_by(ProgramacaoDB.id.desc()).limit(max(min(limit, 500), 1)))
    out = []
    for programacao in result.scalars().all():
        status_value = status_ref(programacao)
        if status_value in CANCELLED_STATUSES:
            continue
        if is_closed(programacao):
            continue
        out.append(
            DespesaProgramacaoOption(
                codigo_programacao=upper_text(programacao.codigo_programacao),
                motorista=upper_text(programacao.motorista),
                veiculo=upper_text(programacao.veiculo),
                status=status_value,
                prestacao_status=upper_text(programacao.prestacao_status or "PENDENTE"),
                fechada=is_closed(programacao),
            )
        )
    return out


@router.get("/mortalidade/fotos")
async def listar_fotos_mortalidade(
    limit: int = 500,
    periodo: str = Query("TODAS", description="TODAS ou quantidade de dias: 7,15,30,60,90,180"),
    data_inicio: str = Query("", description="Data inicial YYYY-MM-DD"),
    data_fim: str = Query("", description="Data final YYYY-MM-DD"),
    codigo_programacao: str = Query("", description="Filtro por programacao"),
    motorista: str = Query("", description="Filtro por motorista"),
    nf: str = Query("", description="Filtro por nota fiscal"),
    escopo: str = Query("TODOS", description="TODOS|CLIENTE|DOA"),
    busca: str = Query("", description="Busca por cliente, pedido, observacao, foto ou codigo"),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    limit = max(1, min(limit, 1000))
    fetch_limit = max(limit * 4, 1000)
    fotos_result = await db.execute(
        text(
            """
            SELECT
                rf.id_foto, rf.codigo_programacao, rf.categoria, rf.tipo_registro, rf.cod_cliente,
                rf.cliente_nome, rf.pedido, rf.id_vinculo, rf.path_local, rf.storage_path,
                rf.arquivo_nome, rf.mime_type, rf.tamanho_bytes, rf.motorista_codigo,
                rf.motorista_nome, rf.registrado_em, rf.payload_json
              FROM rota_fotos rf
             WHERE UPPER(COALESCE(rf.categoria, '')) LIKE '%MORTALIDADE%'
                OR UPPER(COALESCE(rf.categoria, '')) LIKE '%DOA%'
                OR UPPER(COALESCE(rf.tipo_registro, '')) LIKE '%MORTALIDADE%'
                OR UPPER(COALESCE(rf.tipo_registro, '')) LIKE '%DOA%'
             ORDER BY rf.id DESC
             LIMIT :limit
            """
        ),
        {"limit": fetch_limit},
    )
    fotos = []
    fotos_por_cliente: dict[tuple[str, str, str], dict[str, Any]] = {}
    fotos_por_programacao: dict[str, dict[str, Any]] = {}
    for row in fotos_result.mappings().all():
        foto = dict(row)
        foto["payload"] = parse_json_dict(foto.pop("payload_json", ""))
        codigo = upper_text(foto.get("codigo_programacao"))
        cliente_key = (codigo, upper_text(foto.get("cod_cliente")), upper_text(foto.get("pedido")))
        categoria = upper_text(foto.get("categoria"))
        tipo = upper_text(foto.get("tipo_registro"))
        is_cliente = "CLIENTE" in categoria or "CLIENTE" in tipo or bool(foto.get("cod_cliente"))
        if is_cliente:
            fotos_por_cliente.setdefault(cliente_key, foto)
        else:
            fotos_por_programacao.setdefault(codigo, foto)
        fotos.append(foto)

    def data_hora(value: Any) -> tuple[str, str]:
        raw = str(value or "").strip()
        if not raw:
            return "", ""
        if "T" in raw:
            raw = raw.replace("T", " ")
        parts = raw.split()
        data = parts[0] if parts else raw
        hora = parts[1][:8] if len(parts) > 1 else ""
        return data, hora

    def clean_text(value: Any) -> str:
        return str(value or "").strip()

    def media_programacao(row: dict[str, Any]) -> float:
        return normalize_media_kg(row.get("media") or row.get("media_1") or row.get("media_2") or row.get("media_3"))

    def preco_compra(row: dict[str, Any]) -> float:
        return safe_float(row.get("nf_preco"), 0.0) or safe_float(row.get("preco_nf"), 0.0)

    def kg_media_from_item(row: dict[str, Any]) -> tuple[float, str]:
        media = normalize_media_kg(row.get("media_aplicada"))
        if media > 0:
            return media, "MEDIA_CLIENTE"
        peso_previsto = safe_float(row.get("peso_previsto"), 0.0)
        caixas = safe_float(row.get("caixas_atual"), 0.0) or safe_float(row.get("qnt_caixas"), 0.0) or safe_float(row.get("caixas"), 0.0)
        aves_por_caixa = safe_float(row.get("aves_por_caixa"), 0.0)
        if peso_previsto > 0 and caixas > 0 and aves_por_caixa > 0:
            return peso_previsto / (caixas * aves_por_caixa), "PESO_PREVISTO_CAIXAS"
        peso_item = safe_float(row.get("kg_item"), 0.0) or safe_float(row.get("kg_cliente"), 0.0)
        if peso_item > 0 and caixas > 0 and aves_por_caixa > 0:
            return peso_item / (caixas * aves_por_caixa), "ITEM_CAIXAS"
        media_carga = media_programacao(row)
        if media_carga > 0:
            return media_carga, "MEDIA_PROGRAMACAO"
        return 0.0, "SEM_MEDIA"

    def preco_cliente_from_item(row: dict[str, Any], detalhe: dict[str, Any]) -> tuple[float, str]:
        opcoes = (
            (detalhe.get("preco_cliente"), "DETALHE_CLIENTE"),
            (row.get("preco_atual"), "CONTROLE_PRECO_ATUAL"),
            (row.get("item_preco_atual"), "ITEM_PRECO_ATUAL"),
            (row.get("item_preco"), "ITEM_PRECO"),
        )
        for value, source in opcoes:
            preco = safe_float(value, 0.0)
            if preco > 0:
                return preco, source
        compra = preco_compra(row)
        if compra > 0:
            return compra, "PRECO_COMPRA_FALLBACK"
        return 0.0, "SEM_PRECO"

    def arquivo_foto(foto: dict[str, Any] | None) -> str:
        if not foto:
            return ""
        return clean_text(foto.get("storage_path") or foto.get("path_local") or foto.get("arquivo_nome"))

    records: list[dict[str, Any]] = []

    doa_result = await db.execute(
        text(
            """
            SELECT p.codigo_programacao, p.nf_numero, p.num_nf, p.motorista, p.veiculo,
                   p.local_rota, p.tipo_rota, p.local_carregamento, p.granja_carregada,
                   p.media, p.media_1, p.media_2, p.media_3, p.nf_preco, p.preco_nf,
                   p.nf_kg, p.kg_carregado, p.nf_kg_carregado,
                   p.mortalidade_transbordo_aves, p.mortalidade_transbordo_kg,
                   p.obs_transbordo, p.data_criacao, p.data, p.status
              FROM programacoes p
             WHERE COALESCE(p.mortalidade_transbordo_aves, 0) > 0
                OR COALESCE(p.mortalidade_transbordo_kg, 0) > 0
                OR TRIM(COALESCE(p.obs_transbordo, '')) <> ''
             ORDER BY COALESCE(p.data_criacao, p.data, '') DESC, p.id DESC
             LIMIT :limit
            """
        ),
        {"limit": fetch_limit},
    )
    for row in doa_result.mappings().all():
        item = dict(row)
        codigo = upper_text(item.get("codigo_programacao"))
        foto = fotos_por_programacao.get(codigo)
        media = media_programacao(item)
        aves = safe_int(item.get("mortalidade_transbordo_aves"), 0)
        kg = safe_float(item.get("mortalidade_transbordo_kg"), 0.0)
        if kg <= 0 and aves > 0 and media > 0:
            kg = aves * media
            fonte_kg = "UNIDADES_X_MEDIA_PROGRAMACAO"
        else:
            fonte_kg = "KG_TRANSBORDO"
        preco = preco_compra(item)
        alertas = []
        if kg <= 0 and aves > 0:
            alertas.append("KG_TRANSBORDO_NAO_INFORMADO")
        if preco <= 0 and kg > 0:
            alertas.append("PRECO_COMPRA_NAO_INFORMADO")
        data, hora = data_hora((foto or {}).get("registrado_em") or item.get("data_criacao") or item.get("data"))
        records.append(
            {
                "escopo": "DOA_TRANSBORDO",
                "codigo_programacao": codigo,
                "num_nf": clean_text(item.get("nf_numero") or item.get("num_nf")),
                "motorista": upper_text(item.get("motorista")),
                "veiculo": upper_text(item.get("veiculo")),
                "rota": upper_text(item.get("local_rota") or item.get("tipo_rota")),
                "local_rota": upper_text(item.get("local_rota") or item.get("tipo_rota")),
                "local_carregamento": upper_text(item.get("local_carregamento") or item.get("granja_carregada")),
                "data": data,
                "hora": hora,
                "mortalidade_doa_aves": aves,
                "mortalidade_doa_kg": two_decimal(kg),
                "mortalidade_cliente_aves": 0,
                "mortalidade_cliente_kg": 0.0,
                "media_carregamento": two_decimal(media),
                "kg_afeta_carga": two_decimal(kg),
                "preco_compra": money(preco),
                "fonte_kg": fonte_kg,
                "fonte_preco": "NF_PRECO" if preco > 0 else "SEM_PRECO",
                "alertas": alertas,
                "valor_afetado": money(kg * preco),
                "obs": clean_text(item.get("obs_transbordo")),
                "foto": arquivo_foto(foto),
                "id_foto": clean_text((foto or {}).get("id_foto")),
                **(foto or {}),
            }
        )

    controle_info = await db.execute(text("PRAGMA table_info(programacao_itens_controle)"))
    controle_cols = {str(row[1]) for row in controle_info.fetchall()}
    item_info = await db.execute(text("PRAGMA table_info(programacao_itens)"))
    item_cols = {str(row[1]) for row in item_info.fetchall()}
    aves_por_caixa_expr = "pc.aves_por_caixa" if "aves_por_caixa" in controle_cols else "0"
    item_caixas_expr = "pi.caixas" if "caixas" in item_cols else "0"
    kg_cliente_expr = "pi.kg_cliente" if "kg_cliente" in item_cols else "0"
    item_preco_atual_expr = "pi.preco_atual" if "preco_atual" in item_cols else "0"

    cliente_result = await db.execute(
        text(
            f"""
            SELECT pc.codigo_programacao, pc.cod_cliente, pc.pedido, pc.mortalidade_aves,
                   pc.media_aplicada, pc.peso_previsto, pc.caixas_atual, {aves_por_caixa_expr} AS aves_por_caixa, pc.status_pedido,
                   pc.alteracao_tipo, pc.alteracao_detalhe, pc.preco_atual,
                   pc.timestamp_entrega, pc.updated_at,
                   COALESCE(pc.lat_entrega, pc.lat_evento) AS latitude,
                   COALESCE(pc.lon_entrega, pc.lon_evento) AS longitude,
                   p.nf_numero, p.num_nf, p.motorista, p.veiculo,
                   p.media, p.media_1, p.media_2, p.media_3, p.nf_preco, p.preco_nf,
                   pi.nome_cliente, {item_caixas_expr} AS caixas, pi.qnt_caixas, pi.kg AS kg_item,
                   {kg_cliente_expr} AS kg_cliente, pi.preco AS item_preco, {item_preco_atual_expr} AS item_preco_atual
              FROM programacao_itens_controle pc
              LEFT JOIN programacoes p
                ON UPPER(p.codigo_programacao)=UPPER(pc.codigo_programacao)
              LEFT JOIN programacao_itens pi
                ON UPPER(pi.codigo_programacao)=UPPER(pc.codigo_programacao)
               AND UPPER(COALESCE(pi.cod_cliente, ''))=UPPER(COALESCE(pc.cod_cliente, ''))
               AND COALESCE(pi.pedido, '')=COALESCE(pc.pedido, '')
             WHERE COALESCE(pc.mortalidade_aves, 0) > 0
             ORDER BY COALESCE(pc.timestamp_entrega, pc.updated_at, '') DESC, pc.id DESC
             LIMIT :limit
            """
        ),
        {"limit": fetch_limit},
    )
    for row in cliente_result.mappings().all():
        item = dict(row)
        codigo = upper_text(item.get("codigo_programacao"))
        key = (codigo, upper_text(item.get("cod_cliente")), upper_text(item.get("pedido")))
        foto = fotos_por_cliente.get(key)
        detalhe = parse_json_dict(item.get("alteracao_detalhe"))
        media_carga = media_programacao(item)
        media_cliente, fonte_kg = kg_media_from_item(item)
        aves = safe_int(item.get("mortalidade_aves"), 0)
        kg = aves * media_cliente if aves > 0 and media_cliente > 0 else 0.0
        preco, fonte_preco = preco_cliente_from_item(item, detalhe)
        alertas = []
        if kg <= 0 and aves > 0:
            alertas.append("MEDIA_CLIENTE_NAO_IDENTIFICADA")
        if preco <= 0 and kg > 0:
            alertas.append("PRECO_CLIENTE_NAO_IDENTIFICADO")
        data, hora = data_hora((foto or {}).get("registrado_em") or item.get("timestamp_entrega") or item.get("updated_at"))
        cliente_nome = upper_text(detalhe.get("cliente") or item.get("nome_cliente") or item.get("cod_cliente"))
        records.append(
            {
                "escopo": "CLIENTE",
                "codigo_programacao": codigo,
                "num_nf": clean_text(item.get("nf_numero") or item.get("num_nf")),
                "motorista": upper_text(item.get("motorista")),
                "veiculo": upper_text(item.get("veiculo")),
                "cod_cliente": upper_text(item.get("cod_cliente")),
                "nome_cliente": cliente_nome,
                "cliente_nome": cliente_nome,
                "vendedor": upper_text(detalhe.get("vendedor")),
                "pedido": clean_text(item.get("pedido")),
                "status_pedido": clean_text(item.get("status_pedido")),
                "data": data,
                "hora": hora,
                "data_pedido": clean_text(detalhe.get("data_pedido")),
                "motivo": clean_text(detalhe.get("motivo")),
                "origem": upper_text(detalhe.get("origem") or item.get("alteracao_tipo")),
                "mortalidade_cliente_aves": aves,
                "mortalidade_cliente_kg": two_decimal(kg),
                "mortalidade_doa_aves": 0,
                "mortalidade_doa_kg": 0.0,
                "media_carregamento": two_decimal(media_carga),
                "media_cliente": two_decimal(media_cliente),
                "kg_afeta_carga": two_decimal(kg),
                "preco_cliente": money(preco),
                "preco_compra": money(preco),
                "fonte_kg": fonte_kg,
                "fonte_preco": fonte_preco,
                "alertas": alertas,
                "valor_afetado": money(kg * preco),
                "latitude": item.get("latitude"),
                "longitude": item.get("longitude"),
                "foto": arquivo_foto(foto),
                "id_foto": clean_text((foto or {}).get("id_foto")),
                **(foto or {}),
            }
        )

    # Keep photo-only evidence visible even if the operational mortality number was not filled yet.
    existing_photo_ids = {clean_text(item.get("id_foto")) for item in records if clean_text(item.get("id_foto"))}
    for foto in fotos:
        foto_id = clean_text(foto.get("id_foto"))
        if foto_id in existing_photo_ids:
            continue
        codigo = upper_text(foto.get("codigo_programacao"))
        is_cliente = bool(foto.get("cod_cliente"))
        data, hora = data_hora(foto.get("registrado_em"))
        records.append(
            {
                **foto,
                "escopo": "CLIENTE" if is_cliente else "DOA_TRANSBORDO",
                "codigo_programacao": codigo,
                "num_nf": "",
                "motorista": upper_text(foto.get("motorista_nome")),
                "data": data,
                "hora": hora,
                "mortalidade_cliente_aves": 0,
                "mortalidade_cliente_kg": 0.0,
                "mortalidade_doa_aves": 0,
                "mortalidade_doa_kg": 0.0,
                "media_carregamento": 0.0,
                "kg_afeta_carga": 0.0,
                "preco_compra": 0.0,
                "fonte_kg": "FOTO_SEM_NUMERO",
                "fonte_preco": "FOTO_SEM_NUMERO",
                "alertas": ["FOTO_SEM_OCORRENCIA_NUMERICA"],
                "valor_afetado": 0.0,
                "foto": arquivo_foto(foto),
            }
        )

    def iso_date(value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        raw = raw.replace("T", " ").split()[0]
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            return raw
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", raw):
            day, month, year = raw.split("/")
            return f"{year}-{month}-{day}"
        return ""

    def period_start() -> str:
        txt = upper_text(periodo)
        if txt in {"", "TODAS", "TODO", "ALL"}:
            return ""
        try:
            days = max(int(re.sub(r"\D+", "", txt) or "0"), 0)
        except Exception:
            days = 0
        if days <= 0:
            return ""
        return (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")

    start_date = iso_date(data_inicio) or period_start()
    end_date = iso_date(data_fim)
    codigo_filter = upper_text(codigo_programacao)
    motorista_filter = upper_text(motorista)
    nf_filter = upper_text(nf)
    escopo_filter = upper_text(escopo)
    busca_filter = upper_text(busca)

    def matches_filter(item: dict[str, Any]) -> bool:
        item_date = iso_date(item.get("data") or item.get("registrado_em"))
        if start_date and (not item_date or item_date < start_date):
            return False
        if end_date and (not item_date or item_date > end_date):
            return False
        if codigo_filter and codigo_filter not in upper_text(item.get("codigo_programacao")):
            return False
        if motorista_filter and motorista_filter not in upper_text(item.get("motorista") or item.get("motorista_nome")):
            return False
        if nf_filter and nf_filter not in upper_text(item.get("num_nf") or item.get("nota_fiscal")):
            return False
        if escopo_filter in {"CLIENTE", "DOA", "DOA_TRANSBORDO", "TRANSBORDO"}:
            item_scope = upper_text(item.get("escopo"))
            if escopo_filter == "CLIENTE" and item_scope != "CLIENTE":
                return False
            if escopo_filter != "CLIENTE" and item_scope == "CLIENTE":
                return False
        if busca_filter:
            haystack = " | ".join(
                upper_text(item.get(key))
                for key in (
                    "codigo_programacao",
                    "num_nf",
                    "motorista",
                    "cod_cliente",
                    "nome_cliente",
                    "cliente_nome",
                    "pedido",
                    "status_pedido",
                    "obs",
                    "foto",
                    "arquivo_nome",
                )
            )
            if busca_filter not in haystack:
                return False
        return True

    records = [item for item in records if matches_filter(item)]
    records = records[:limit]
    por_programacao: dict[str, dict[str, Any]] = {}
    for item in records:
        codigo = upper_text(item.get("codigo_programacao"))
        agg = por_programacao.setdefault(
            codigo,
            {
                "codigo_programacao": codigo,
                "num_nf": clean_text(item.get("num_nf")),
                "motorista": upper_text(item.get("motorista")),
                "fotos": 0,
                "mortalidade_cliente_aves": 0,
                "mortalidade_cliente_kg": 0.0,
                "mortalidade_doa_aves": 0,
                "mortalidade_doa_kg": 0.0,
                "kg_afetado": 0.0,
                "valor_afetado": 0.0,
            },
        )
        if clean_text(item.get("id_foto")):
            agg["fotos"] += 1
        agg["mortalidade_cliente_aves"] += safe_int(item.get("mortalidade_cliente_aves"), 0)
        agg["mortalidade_cliente_kg"] = two_decimal(safe_float(agg["mortalidade_cliente_kg"], 0.0) + safe_float(item.get("mortalidade_cliente_kg"), 0.0))
        agg["mortalidade_doa_aves"] += safe_int(item.get("mortalidade_doa_aves"), 0)
        agg["mortalidade_doa_kg"] = two_decimal(safe_float(agg["mortalidade_doa_kg"], 0.0) + safe_float(item.get("mortalidade_doa_kg"), 0.0))
        agg["kg_afetado"] = two_decimal(safe_float(agg["kg_afetado"], 0.0) + safe_float(item.get("kg_afeta_carga"), 0.0))
        agg["valor_afetado"] = money(safe_float(agg["valor_afetado"], 0.0) + safe_float(item.get("valor_afetado"), 0.0))

    total_cliente_aves = sum(safe_int(item.get("mortalidade_cliente_aves"), 0) for item in records)
    total_cliente_kg = sum(safe_float(item.get("mortalidade_cliente_kg"), 0.0) for item in records)
    total_doa_aves = sum(safe_int(item.get("mortalidade_doa_aves"), 0) for item in records)
    total_doa_kg = sum(safe_float(item.get("mortalidade_doa_kg"), 0.0) for item in records)
    total_kg = sum(safe_float(item.get("kg_afeta_carga"), 0.0) for item in records)
    total_valor = sum(safe_float(item.get("valor_afetado"), 0.0) for item in records)
    total_valor_cliente = sum(safe_float(item.get("valor_afetado"), 0.0) for item in records if upper_text(item.get("escopo")) == "CLIENTE")
    total_valor_doa = sum(safe_float(item.get("valor_afetado"), 0.0) for item in records if upper_text(item.get("escopo")) != "CLIENTE")
    return {
        "filtros": {
            "periodo": periodo,
            "data_inicio": start_date,
            "data_fim": end_date,
            "codigo_programacao": codigo_filter,
            "motorista": motorista_filter,
            "nf": nf_filter,
            "escopo": escopo_filter,
            "busca": busca_filter,
            "limit": limit,
        },
        "kpis": {
            "registros": len(records),
            "fotos": len([item for item in records if clean_text(item.get("id_foto"))]),
            "mortalidade_cliente_aves": total_cliente_aves,
            "mortalidade_cliente_kg": two_decimal(total_cliente_kg),
            "mortalidade_doa_aves": total_doa_aves,
            "mortalidade_doa_kg": two_decimal(total_doa_kg),
            "kg_afetado": two_decimal(total_kg),
            "valor_afetado": money(total_valor),
            "valor_cliente": money(total_valor_cliente),
            "valor_operacao": money(total_valor_doa),
            "programacoes": len(por_programacao),
        },
        "por_programacao": list(por_programacao.values()),
        "fotos": records,
    }


@router.get("/mortalidade/fotos/{id_foto}/arquivo")
async def obter_arquivo_foto_mortalidade(
    id_foto: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    foto_id = str(id_foto or "").strip()
    if not foto_id:
        raise HTTPException(status_code=404, detail="Foto nao encontrada.")
    result = await db.execute(
        text(
            """
            SELECT path_local, storage_path, arquivo_nome, mime_type
              FROM rota_fotos
             WHERE id_foto=:id_foto
             LIMIT 1
            """
        ),
        {"id_foto": foto_id},
    )
    row = result.mappings().first()
    if not row:
        raise HTTPException(status_code=404, detail="Foto nao encontrada.")
    candidates = [row.get("storage_path"), row.get("path_local"), row.get("arquivo_nome")]
    for candidate in candidates:
        path = Path(str(candidate or "").strip())
        if path.exists() and path.is_file():
            return FileResponse(path, media_type=str(row.get("mime_type") or "application/octet-stream"), filename=path.name)
    raise HTTPException(status_code=404, detail="Arquivo da foto nao encontrado no servidor.")


@router.post("/mortalidade/manual", response_model=MortalidadeManualResponse)
async def registrar_mortalidade_manual(
    request: Request,
    codigo_programacao: str = Form(...),
    nota_fiscal: str = Form(""),
    pedido: str = Form(""),
    cliente: str = Form(...),
    vendedor: str = Form(""),
    preco_cliente: float = Form(0),
    media: float = Form(0),
    data_pedido: str = Form(""),
    mortalidade_aves: int = Form(...),
    motivo: str = Form(...),
    foto: UploadFile | None = File(default=None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    codigo_input = upper_text(codigo_programacao)
    if not codigo_input:
        raise HTTPException(status_code=400, detail="Informe o numero da programacao.")
    programacao = await get_programacao_by_codigo(db, codigo_input)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada.")
    assert_can_mutate(programacao)

    aves = max(safe_int(mortalidade_aves, 0), 0)
    if aves <= 0:
        raise HTTPException(status_code=400, detail="Informe a quantidade de aves mortas.")
    media_aplicada = max(safe_float(media, 0.0), 0.0)
    preco = max(safe_float(preco_cliente, 0.0), 0.0)
    cliente_nome = upper_text(cliente)
    pedido_txt = str(pedido or "").strip()
    motivo_txt = str(motivo or "").strip()
    if not cliente_nome:
        raise HTTPException(status_code=400, detail="Informe o cliente.")
    if not motivo_txt:
        raise HTTPException(status_code=400, detail="Informe o motivo da ocorrencia.")

    codigo = upper_text(programacao.codigo_programacao)
    nf_txt = upper_text(nota_fiscal)
    if nf_txt:
        programacao.nf_numero = nf_txt
        programacao.num_nf = nf_txt
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    metadata = {
        "origem": "MANUAL",
        "nota_fiscal": nf_txt,
        "pedido": pedido_txt,
        "cliente": cliente_nome,
        "vendedor": upper_text(vendedor),
        "preco_cliente": money(preco),
        "media": two_decimal(media_aplicada),
        "data_pedido": str(data_pedido or "").strip(),
        "mortalidade_aves": aves,
        "mortalidade_kg": two_decimal(aves * media_aplicada),
        "valor_desconto": money(aves * media_aplicada * preco),
        "motivo": motivo_txt,
        "registrado_em": now,
        "registrado_por": getattr(current_user, "username", "") or getattr(current_user, "nome", "") or "",
    }
    foto_ref = await save_mortalidade_manual_photo(codigo=codigo, pedido=pedido_txt, foto=foto)
    if foto_ref:
        foto_ref.update(
            {
                "codigo_programacao": codigo,
                "categoria": "MORTALIDADE_CLIENTE_MANUAL",
                "tipo_registro": "MORTALIDADE_MANUAL",
                "cod_cliente": cliente_nome,
                "cliente_nome": cliente_nome,
                "pedido": pedido_txt,
                "motorista_nome": upper_text(programacao.motorista),
                "payload": dict(metadata),
            }
        )
        metadata["foto"] = {key: value for key, value in foto_ref.items() if key != "payload"}

    result = await db.execute(
        select(ProgramacaoItemControleDB)
        .where(func.upper(ProgramacaoItemControleDB.codigo_programacao) == codigo)
        .where(func.upper(func.coalesce(ProgramacaoItemControleDB.cod_cliente, "")) == cliente_nome)
        .where(func.coalesce(ProgramacaoItemControleDB.pedido, "") == pedido_txt)
        .limit(1)
    )
    controle = result.scalar_one_or_none()
    if controle is None:
        controle = ProgramacaoItemControleDB(
            codigo_programacao=codigo,
            cod_cliente=cliente_nome,
            pedido=pedido_txt,
        )
        db.add(controle)

    controle.status_pedido = controle.status_pedido or "ENTREGUE"
    controle.alteracao_tipo = "MORTALIDADE_MANUAL"
    controle.alteracao_detalhe = json.dumps(metadata, ensure_ascii=False, sort_keys=True)
    controle.mortalidade_aves = aves
    controle.media_aplicada = media_aplicada
    controle.peso_previsto = two_decimal(aves * media_aplicada)
    controle.preco_atual = preco
    controle.alterado_em = now
    controle.alterado_por = metadata["registrado_por"] or "WEB"
    controle.timestamp_entrega = controle.timestamp_entrega or now
    controle.updated_at = now
    if foto_ref:
        foto_json = json.dumps(foto_ref, ensure_ascii=False, sort_keys=True)
        controle.foto_mortalidade_ref_json = foto_json
        controle.foto_mortalidade_path = foto_ref.get("path_local", "")
        controle.mortalidade_foto_path = foto_ref.get("path_local", "")
        await db.execute(
            text(
                """
                INSERT INTO rota_fotos (
                    id_foto, codigo_programacao, categoria, tipo_registro, cod_cliente,
                    cliente_nome, pedido, path_local, storage_path, arquivo_nome,
                    mime_type, tamanho_bytes, motorista_nome, registrado_em, payload_json
                ) VALUES (
                    :id_foto, :codigo_programacao, :categoria, :tipo_registro, :cod_cliente,
                    :cliente_nome, :pedido, :path_local, :storage_path, :arquivo_nome,
                    :mime_type, :tamanho_bytes, :motorista_nome, :registrado_em, :payload_json
                )
                ON CONFLICT(id_foto) DO UPDATE SET
                    path_local=excluded.path_local,
                    storage_path=excluded.storage_path,
                    arquivo_nome=excluded.arquivo_nome,
                    payload_json=excluded.payload_json
                """
            ),
            {
                **foto_ref,
                "payload_json": foto_json,
            },
        )

    await db.execute(
        text(
            """
            INSERT INTO programacao_itens_log (
                codigo_programacao, cod_cliente, pedido, evento, payload_json, registrado_em, created_at
            ) VALUES (
                :codigo_programacao, :cod_cliente, :pedido, 'mortalidade_manual', :payload_json, :registrado_em, :created_at
            )
            """
        ),
        {
            "codigo_programacao": codigo,
            "cod_cliente": cliente_nome,
            "pedido": pedido_txt,
            "payload_json": json.dumps(metadata, ensure_ascii=False, sort_keys=True),
            "registrado_em": now,
            "created_at": now,
        },
    )
    record_audit_log(
        db,
        action="mortalidade_manual_registrada",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata=metadata,
    )
    await db.commit()
    return MortalidadeManualResponse(
        codigo_programacao=codigo,
        pedido=pedido_txt,
        cliente=cliente_nome,
        mortalidade_aves=aves,
        mortalidade_kg=two_decimal(aves * media_aplicada),
        valor_desconto=money(aves * media_aplicada * preco),
        foto=foto_ref,
    )


@router.get("/{codigo_programacao}/bundle", response_model=DespesasBundleResponse)
async def carregar_despesas_bundle(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_open_prestacao(programacao)
    if await sync_diarias_despesas(db, programacao):
        await refresh_km_metrics(db, programacao)
        await db.commit()
        await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.get("/{codigo_programacao}/pdf")
async def despesas_pdf(
    codigo_programacao: str,
    reimpressao: bool = Query(False),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    if await sync_diarias_despesas(db, programacao):
        await refresh_km_metrics(db, programacao)
        await db.commit()
        await db.refresh(programacao)
    bundle = await serialize_bundle(db, programacao)
    codigo = upper_text(programacao.codigo_programacao)
    recebimentos = await recebimento_rows(db, codigo)
    itens = await item_rows(db, codigo)
    controles = await controle_rows(db, codigo)
    transferencias_pdf = await transferencias_operacionais_pdf(db, codigo)
    equipe_txt = await resolve_equipe_pdf(db, programacao.equipe)
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4, pageCompression=0)
    pdf.setTitle(f"PRESTACAO DE CONTAS - PROGRAMACAO {codigo}")
    draw_despesas_pdf_page(
        pdf,
        bundle,
        programacao=programacao,
        recebimentos=recebimentos,
        itens=itens,
        controles=controles,
        transferencias_operacionais=transferencias_pdf,
        equipe_txt=equipe_txt,
        reimpressao=reimpressao,
    )
    pdf.save()
    buffer.seek(0)
    safe_name = codigo.replace("/", "_") or "PROGRAMACAO"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="PRESTACAO_{safe_name}.pdf"'},
    )


@router.post("/{codigo_programacao}/despesas", response_model=DespesaItem, status_code=status.HTTP_201_CREATED)
async def criar_despesa(
    codigo_programacao: str,
    payload: DespesaPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    codigo = upper_text(programacao.codigo_programacao)
    despesa = DespesaDB(
        codigo_programacao=codigo,
        descricao=upper_text(payload.descricao),
        valor=money(payload.valor),
        data_registro=payload.data_registro or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        tipo_despesa=upper_text(payload.tipo_despesa or "ROTA"),
        categoria=upper_text(payload.categoria),
        motorista=upper_text(programacao.motorista),
        veiculo=upper_text(programacao.veiculo),
        observacao=str(payload.observacao or "").strip(),
    )
    db.add(despesa)
    record_audit_log(
        db,
        action="despesa_criada",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata={"descricao": despesa.descricao, "valor": despesa.valor, "categoria": despesa.categoria},
    )
    await db.flush()
    await registrar_roteiro_operacional(
        db,
        tipo_evento="DESPESA_WEB",
        codigo_programacao=codigo,
        origem="WEB",
        destino=despesa.categoria or despesa.descricao,
        motorista_nome=programacao.motorista,
        data_hora=despesa.data_registro,
        observacao=despesa.observacao or despesa.descricao,
        payload={"despesa_id": despesa.id, "descricao": despesa.descricao, "valor": despesa.valor, "categoria": despesa.categoria},
    )
    await refresh_km_metrics(db, programacao)
    await db.commit()
    await db.refresh(despesa)
    return despesa_to_response(despesa)


@router.patch("/{codigo_programacao}/despesas/{despesa_id}", response_model=DespesaItem)
async def atualizar_despesa(
    codigo_programacao: str,
    despesa_id: int,
    payload: DespesaPatchPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    despesa = await db.get(DespesaDB, despesa_id)
    codigo = upper_text(programacao.codigo_programacao)
    if not despesa or upper_text(despesa.codigo_programacao) != codigo:
        raise HTTPException(status_code=404, detail="Despesa nao encontrada")

    if "descricao" in payload.model_fields_set and payload.descricao is not None:
        despesa.descricao = upper_text(payload.descricao)
    if "valor" in payload.model_fields_set and payload.valor is not None:
        despesa.valor = money(payload.valor)
    if "categoria" in payload.model_fields_set:
        despesa.categoria = upper_text(payload.categoria)
    if "tipo_despesa" in payload.model_fields_set and payload.tipo_despesa is not None:
        despesa.tipo_despesa = upper_text(payload.tipo_despesa or "ROTA")
    if "observacao" in payload.model_fields_set:
        despesa.observacao = str(payload.observacao or "").strip()
    if "data_registro" in payload.model_fields_set and payload.data_registro:
        despesa.data_registro = payload.data_registro

    record_audit_log(
        db,
        action="despesa_atualizada",
        actor_user=current_user,
        entity_type="despesa",
        entity_id=str(despesa.id),
        ip_address=client_ip_from_request(request),
        metadata={"codigo_programacao": codigo, "valor": despesa.valor},
    )
    await refresh_km_metrics(db, programacao)
    await db.commit()
    await db.refresh(despesa)
    return despesa_to_response(despesa)


@router.delete("/{codigo_programacao}/despesas/{despesa_id}", response_model=DespesasBundleResponse)
async def excluir_despesa(
    codigo_programacao: str,
    despesa_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    despesa = await db.get(DespesaDB, despesa_id)
    codigo = upper_text(programacao.codigo_programacao)
    if not despesa or upper_text(despesa.codigo_programacao) != codigo:
        raise HTTPException(status_code=404, detail="Despesa nao encontrada")
    await db.delete(despesa)
    record_audit_log(
        db,
        action="despesa_excluida",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"despesa_id": despesa_id},
    )
    await refresh_km_metrics(db, programacao)
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.put("/{codigo_programacao}/rota", response_model=DespesasBundleResponse)
async def salvar_rota_despesas(
    codigo_programacao: str,
    payload: RotaPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    programacao.km_inicial = safe_float(payload.km_inicial, 0.0)
    programacao.km_final = safe_float(payload.km_final, 0.0)
    programacao.litros = safe_float(payload.litros, 0.0)
    programacao.rota_observacao = str(payload.rota_observacao or "").strip()
    await refresh_km_metrics(db, programacao)
    await registrar_roteiro_operacional(
        db,
        tipo_evento="ROTA_WEB",
        codigo_programacao=upper_text(programacao.codigo_programacao),
        origem="WEB",
        destino=upper_text(programacao.veiculo),
        motorista_nome=programacao.motorista,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao=programacao.rota_observacao,
        payload={
            "km_inicial": programacao.km_inicial,
            "km_final": programacao.km_final,
            "km_rodado": programacao.km_rodado,
            "litros": programacao.litros,
            "media_km_l": programacao.media_km_l,
        },
    )
    record_audit_log(
        db,
        action="despesas_rota_salva",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        ip_address=client_ip_from_request(request),
        metadata={"km_rodado": programacao.km_rodado, "media_km_l": programacao.media_km_l},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.put("/{codigo_programacao}/nf", response_model=DespesasBundleResponse)
async def salvar_nf_despesas(
    codigo_programacao: str,
    payload: NfPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    codigo = upper_text(programacao.codigo_programacao)
    kg_itens, _receita_itens, _caixas_itens = await itens_totais(db, codigo)
    kg_carregado = safe_float(payload.nf_kg_carregado, 0.0) if payload.nf_kg_carregado is not None else safe_float(payload.nf_kg, 0.0)
    kg_vendido = safe_float(payload.nf_kg_vendido, 0.0) if payload.nf_kg_vendido is not None else kg_itens
    saldo = safe_float(payload.nf_saldo, 0.0) if payload.nf_saldo is not None else max(kg_carregado - kg_vendido, 0.0)

    nf_numero = upper_text(payload.nf_numero)
    programacao.nf_numero = nf_numero
    programacao.num_nf = nf_numero
    programacao.nf_kg = safe_float(payload.nf_kg, 0.0)
    programacao.nf_preco = safe_float(payload.nf_preco, 0.0)
    programacao.nf_caixas = safe_int(payload.nf_caixas, 0)
    programacao.nf_kg_carregado = kg_carregado
    programacao.nf_kg_vendido = kg_vendido
    programacao.nf_saldo = max(saldo, 0.0)
    programacao.media = safe_float(payload.nf_media_carregada, 0.0)
    programacao.qnt_aves_caixa_final = safe_int(payload.nf_caixa_final, 0)
    programacao.aves_caixa_final = safe_int(payload.nf_caixa_final, 0)
    programacao.mortalidade_transbordo_aves = safe_int(payload.mortalidade_transbordo_aves, 0)
    programacao.mortalidade_transbordo_kg = safe_float(payload.mortalidade_transbordo_kg, 0.0)
    programacao.obs_transbordo = str(payload.obs_transbordo or "").strip()
    await registrar_roteiro_operacional(
        db,
        tipo_evento="NF_WEB",
        codigo_programacao=codigo,
        origem="WEB",
        destino=programacao.local_carregamento or programacao.local_carregado or "",
        motorista_nome=programacao.motorista,
        caixas=programacao.nf_caixas,
        kg=programacao.nf_kg_carregado or programacao.nf_kg,
        media=programacao.media,
        aves_por_caixa=programacao.aves_caixa_final,
        nf_numero=nf_numero,
        nf_preco=programacao.nf_preco,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao=programacao.obs_transbordo,
        payload=payload.model_dump(),
    )
    record_audit_log(
        db,
        action="despesas_nf_salva",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=codigo,
        ip_address=client_ip_from_request(request),
        metadata={"nf_numero": nf_numero, "nf_kg": programacao.nf_kg, "nf_preco": programacao.nf_preco},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.put("/{codigo_programacao}/financeiro", response_model=DespesasBundleResponse)
async def salvar_financeiro_despesas(
    codigo_programacao: str,
    payload: FinanceiroPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    cedulas = {}
    for ced in CEDULAS:
        qtd = max(safe_int(payload.cedulas.get(str(ced), 0), 0), 0)
        setattr(programacao, f"ced_{ced}_qtd", qtd)
        cedulas[str(ced)] = qtd
    programacao.valor_dinheiro = cedulas_total(cedulas)
    programacao.adiantamento = safe_float(payload.adiantamento, 0.0)
    programacao.adiantamento_rota = safe_float(payload.adiantamento, 0.0)
    programacao.pix_motorista = safe_float(payload.pix_motorista, 0.0)
    await registrar_roteiro_operacional(
        db,
        tipo_evento="FINANCEIRO_WEB",
        codigo_programacao=upper_text(programacao.codigo_programacao),
        origem="WEB",
        destino="PRESTACAO",
        motorista_nome=programacao.motorista,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao="Conferencia financeira da prestacao",
        payload={"cedulas": cedulas, "valor_dinheiro": programacao.valor_dinheiro, "adiantamento": programacao.adiantamento, "pix_motorista": programacao.pix_motorista},
    )
    record_audit_log(
        db,
        action="despesas_financeiro_salvo",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        ip_address=client_ip_from_request(request),
        metadata={"valor_dinheiro": programacao.valor_dinheiro, "pix_motorista": programacao.pix_motorista},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)


@router.post("/{codigo_programacao}/finalizar", response_model=DespesasBundleResponse)
async def finalizar_prestacao_despesas(
    codigo_programacao: str,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    assert_can_mutate(programacao)
    await sync_diarias_despesas(db, programacao)
    await refresh_km_metrics(db, programacao)
    programacao.prestacao_status = "FECHADA"
    programacao.status = "FINALIZADA"
    programacao.status_operacional = "FINALIZADA"
    programacao.finalizada_no_app = 1
    await registrar_roteiro_operacional(
        db,
        tipo_evento="PRESTACAO_FINALIZADA_WEB",
        codigo_programacao=upper_text(programacao.codigo_programacao),
        origem="WEB",
        destino="FECHADA",
        motorista_nome=programacao.motorista,
        data_hora=datetime.now().isoformat(timespec="seconds"),
        observacao="Prestacao finalizada no web",
        payload={"prestacao_status": "FECHADA", "status": "FINALIZADA"},
    )
    record_audit_log(
        db,
        action="despesas_prestacao_finalizada",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"prestacao_status": "FECHADA"},
    )
    await db.commit()
    await db.refresh(programacao)
    return await serialize_bundle(db, programacao)
