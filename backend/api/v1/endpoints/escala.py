# backend/api/v1/endpoints/escala.py
"""
Escala endpoints mirroring the desktop EscalaPage read model.
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta
from io import BytesIO
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import normalize_time, safe_float, safe_int
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.cadastro import AjudanteDB, EscalaFolgaDB, MotoristaDB
from backend.models.programacao import ProgramacaoDB
from backend.models.user import User

router = APIRouter()


class EscalaKpis(BaseModel):
    rotas: int = 0
    motoristas: int = 0
    ajudantes: int = 0
    folgas_motoristas: int = 0
    folgas_ajudantes: int = 0
    mortalidade_media: float = 0
    km_total: float = 0
    km_medio_motorista: float = 0
    media_km_l: float = 0
    horas_medias_motorista: float = 0


class EscalaPessoa(BaseModel):
    nome: str
    rotas: int = 0
    em_rota: int = 0
    ativas: int = 0
    finalizadas: int = 0
    canceladas: int = 0
    local: str = "-"
    km_rodado: float = 0
    horas_trab: float = 0
    carga: str = "equilibrada"
    em_folga: bool = False


class EscalaFolga(BaseModel):
    id: int
    tipo: str
    pessoa_id: str = ""
    pessoa_codigo: str = ""
    pessoa_nome: str
    data_inicio: str
    data_fim: str
    motivo: str = ""
    status: str = "ATIVA"


class EscalaPessoaOption(BaseModel):
    tipo: str
    pessoa_id: str = ""
    pessoa_codigo: str = ""
    pessoa_nome: str
    label: str


class EscalaFolgaCreate(BaseModel):
    tipo: str
    pessoa_id: str = ""
    pessoa_codigo: str = ""
    pessoa_nome: str
    data_inicio: str
    data_fim: str
    motivo: str = ""


class EscalaChartItem(BaseModel):
    nome: str
    horas: float = 0
    dias: int = 0


class EscalaResumoResponse(BaseModel):
    periodo: str
    status: str
    kpis: EscalaKpis
    resumo: str
    recomendacoes: str
    motoristas: list[EscalaPessoa]
    ajudantes: list[EscalaPessoa]
    chart: list[EscalaChartItem]
    folgas: list[EscalaFolga] = []


def pdf_money_number(value: Any, places: int = 2) -> str:
    number = safe_float(value, 0.0)
    return f"{number:.{places}f}".replace(".", ",")


def pdf_wrap_lines(text: Any, max_chars: int = 72, max_lines: int | None = None) -> list[str]:
    lines: list[str] = []
    for raw in str(text or "").splitlines():
        words = raw.strip().split()
        if not words:
            lines.append("")
            continue
        current = ""
        for word in words:
            candidate = word if not current else f"{current} {word}"
            if len(candidate) <= max_chars:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = word[:max_chars]
        if current:
            lines.append(current)
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][: max(0, max_chars - 3)] + "..."
    return lines


def pdf_draw_line(pdf: Any, y: float, text: Any, *, x: int = 40, size: int = 9, bold: bool = False) -> float:
    pdf.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    pdf.drawString(x, y, str(text or "")[:126])
    return y - 13


def pdf_draw_wrapped(pdf: Any, y: float, title: str, text: Any, *, x: int = 40, max_lines: int = 7) -> float:
    y = pdf_draw_line(pdf, y, title, size=10, bold=True)
    pdf.setFont("Helvetica", 8)
    for line in pdf_wrap_lines(text, max_chars=78, max_lines=max_lines):
        pdf.drawString(x, y, line[:118])
        y -= 11
    return y


def pdf_draw_table(
    pdf: Any,
    y: float,
    title: str,
    headers: list[str],
    rows: list[list[Any]],
    *,
    width: float,
    height: float,
) -> float:
    def new_page(page_title: str) -> float:
        pdf.showPage()
        y_new = height - 52
        return pdf_draw_line(pdf, y_new, page_title, size=13, bold=True) - 8

    if y < 100:
        y = new_page("RELATORIO DE ESCALA")
    y = pdf_draw_line(pdf, y, title, size=10, bold=True)
    y = pdf_draw_line(pdf, y, " | ".join(headers), size=7, bold=True)
    for row in rows:
        if y < 58:
            y = new_page(f"{title} (continua)")
            y = pdf_draw_line(pdf, y, " | ".join(headers), size=7, bold=True)
        text = " | ".join(str(value or "") for value in row)
        pdf.setFont("Helvetica", 7)
        pdf.drawString(40, y, text[:142])
        y -= 11
    if not rows:
        y = pdf_draw_line(pdf, y, "Sem dados para o filtro selecionado.", size=8)
    pdf.setFont("Helvetica", 7)
    pdf.drawRightString(width - 40, 34, datetime.now().strftime("%d/%m/%Y %H:%M"))
    return y - 8


def draw_escala_pdf(pdf: Any, data: EscalaResumoResponse) -> None:
    width, height = pdf._pagesize
    kpis = data.kpis
    y = height - 52
    y = pdf_draw_line(pdf, y, "RELATORIO DE ESCALA", size=14, bold=True)
    y = pdf_draw_line(pdf, y, f"Periodo: {data.periodo} | Status: {data.status} | Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 8
    y = pdf_draw_line(pdf, y, "RESUMO EXECUTIVO", bold=True)
    for line in (
        f"Rotas: {kpis.rotas} | Motoristas: {kpis.motoristas} | Ajudantes: {kpis.ajudantes}",
        f"Folgas no periodo: {kpis.folgas_motoristas} motorista(s) | {kpis.folgas_ajudantes} ajudante(s)",
        f"KM total: {pdf_money_number(kpis.km_total, 1)} | KM medio/motorista: {pdf_money_number(kpis.km_medio_motorista, 1)}",
        f"Media km/L: {pdf_money_number(kpis.media_km_l)} | Horas medias/motorista: {pdf_money_number(kpis.horas_medias_motorista)}",
        f"Ocorrencias media/rota: {pdf_money_number(kpis.mortalidade_media)}",
    ):
        y = pdf_draw_line(pdf, y, line)
    y -= 8
    y = pdf_draw_wrapped(pdf, y, "RESUMO OPERACIONAL", data.resumo, max_lines=7)
    y -= 4
    y = pdf_draw_wrapped(pdf, y, "RECOMENDACOES", data.recomendacoes, max_lines=7)
    y -= 8

    motoristas_rows = [
        [
            item.nome,
            item.rotas,
            item.em_rota,
            item.ativas,
            item.finalizadas,
            item.canceladas,
            item.local,
            pdf_money_number(item.km_rodado, 1),
            pdf_money_number(item.horas_trab),
            item.carga.upper(),
        ]
        for item in data.motoristas[:20]
    ]
    y = pdf_draw_table(
        pdf,
        y,
        "DISTRIBUICAO POR MOTORISTA",
        ["Motorista", "Rotas", "Em rota", "Ativas", "Final.", "Canc.", "Local", "KM", "Horas", "Carga"],
        motoristas_rows,
        width=width,
        height=height,
    )
    if len(data.motoristas) > 20:
        y = pdf_draw_line(pdf, y, f"* Exibindo os 20 primeiros motoristas de {len(data.motoristas)} registros.", size=8)
    y -= 4

    ajudantes_rows = [
        [
            item.nome,
            item.rotas,
            item.em_rota,
            item.ativas,
            item.finalizadas,
            item.canceladas,
            pdf_money_number(item.km_rodado, 1),
            pdf_money_number(item.horas_trab),
            item.carga.upper(),
        ]
        for item in data.ajudantes[:20]
    ]
    y = pdf_draw_table(
        pdf,
        y,
        "DISTRIBUICAO POR AJUDANTE",
        ["Ajudante", "Rotas", "Em rota", "Ativas", "Final.", "Canc.", "KM", "Horas", "Carga"],
        ajudantes_rows,
        width=width,
        height=height,
    )
    if len(data.ajudantes) > 20:
        y = pdf_draw_line(pdf, y, f"* Exibindo os 20 primeiros ajudantes de {len(data.ajudantes)} registros.", size=8)

    folgas_rows = [
        [item.tipo, item.pessoa_nome, item.data_inicio, item.data_fim, item.motivo]
        for item in data.folgas[:24]
    ]
    y -= 4
    y = pdf_draw_table(
        pdf,
        y,
        "FOLGAS NO PERIODO",
        ["Tipo", "Pessoa", "Inicio", "Fim", "Motivo"],
        folgas_rows,
        width=width,
        height=height,
    )
    if len(data.folgas) > 24:
        pdf_draw_line(pdf, y, f"* Exibindo as 24 primeiras folgas de {len(data.folgas)} registros.", size=8)


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


def periodo_ref_range(periodo: str) -> tuple[str, str]:
    periodo_norm = upper_text(periodo or "30")
    hoje = datetime.now().date()
    if periodo_norm == "TODAS":
        return "1900-01-01", "2999-12-31"
    dias = max(safe_int(periodo_norm, 30), 0)
    return (hoje - timedelta(days=dias)).isoformat(), hoje.isoformat()


def serialize_folga(row: EscalaFolgaDB) -> EscalaFolga:
    return EscalaFolga(
        id=safe_int(row.id, 0),
        tipo=upper_text(row.tipo),
        pessoa_id=str(row.pessoa_id or ""),
        pessoa_codigo=upper_text(row.pessoa_codigo),
        pessoa_nome=upper_text(row.pessoa_nome),
        data_inicio=str(row.data_inicio or ""),
        data_fim=str(row.data_fim or ""),
        motivo=str(row.motivo or ""),
        status=upper_text(row.status) or "ATIVA",
    )


async def fetch_folgas_ativas(db: AsyncSession, ref_inicio: str, ref_fim: str) -> list[EscalaFolga]:
    result = await db.execute(
        select(EscalaFolgaDB)
        .where(func.upper(func.coalesce(EscalaFolgaDB.status, "ATIVA")) == "ATIVA")
        .order_by(EscalaFolgaDB.data_inicio.asc(), EscalaFolgaDB.pessoa_nome.asc())
    )
    folgas = []
    for row in result.scalars().all():
        if date_ranges_overlap(row.data_inicio, row.data_fim, ref_inicio, ref_fim):
            folgas.append(serialize_folga(row))
    return folgas


def folga_sets(folgas: list[EscalaFolga]) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    motoristas = {"nomes": set(), "codigos": set(), "ids": set()}
    ajudantes = {"nomes": set(), "ids": set()}
    for item in folgas:
        tipo = upper_text(item.tipo)
        if tipo == "MOTORISTA":
            if item.pessoa_nome:
                motoristas["nomes"].add(upper_text(item.pessoa_nome))
            if item.pessoa_codigo:
                motoristas["codigos"].add(upper_text(item.pessoa_codigo))
            if item.pessoa_id:
                motoristas["ids"].add(str(item.pessoa_id or "").strip())
        elif tipo == "AJUDANTE":
            if item.pessoa_nome:
                ajudantes["nomes"].add(upper_text(item.pessoa_nome))
            if item.pessoa_id:
                ajudantes["ids"].add(str(item.pessoa_id or "").strip())
    return motoristas, ajudantes


def pessoa_em_folga(
    *,
    tipo: str,
    nome: Any = "",
    codigo: Any = "",
    pessoa_id: Any = "",
    folga_motoristas: dict[str, set[str]] | None = None,
    folga_ajudantes: dict[str, set[str]] | None = None,
) -> bool:
    tipo_norm = upper_text(tipo)
    nome_norm = upper_text(nome)
    codigo_norm = upper_text(codigo)
    id_norm = str(pessoa_id or "").strip()
    if tipo_norm == "MOTORISTA":
        data = folga_motoristas or {"nomes": set(), "codigos": set(), "ids": set()}
        return bool(
            (nome_norm and nome_norm in data.get("nomes", set()))
            or (codigo_norm and codigo_norm in data.get("codigos", set()))
            or (id_norm and id_norm in data.get("ids", set()))
        )
    if tipo_norm == "AJUDANTE":
        data = folga_ajudantes or {"nomes": set(), "ids": set()}
        return bool(
            (nome_norm and nome_norm in data.get("nomes", set()))
            or (id_norm and id_norm in data.get("ids", set()))
        )
    return False


def count_folgas_pessoas(folgas: list[EscalaFolga], tipo: str) -> int:
    tipo_norm = upper_text(tipo)
    keys = set()
    for item in folgas:
        if upper_text(item.tipo) != tipo_norm:
            continue
        if tipo_norm == "MOTORISTA":
            key = upper_text(item.pessoa_codigo) or str(item.pessoa_id or "").strip() or upper_text(item.pessoa_nome)
        else:
            key = str(item.pessoa_id or "").strip() or upper_text(item.pessoa_nome)
        if key:
            keys.add(key)
    return len(keys)


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


def parse_data_hora(data_raw: Any, hora_raw: Any) -> datetime | None:
    dt_data = parse_data_programacao(data_raw)
    if dt_data is None:
        return None
    hora_txt = str(hora_raw or "").strip()
    if not hora_txt:
        return dt_data.replace(hour=0, minute=0, second=0, microsecond=0)
    normalized = normalize_time(hora_txt)
    if not normalized:
        return dt_data.replace(hour=0, minute=0, second=0, microsecond=0)
    parts = (normalized + ":00:00").split(":")
    return dt_data.replace(
        hour=safe_int(parts[0], 0),
        minute=safe_int(parts[1], 0),
        second=safe_int(parts[2], 0),
        microsecond=0,
    )


def calc_horas_trabalhadas(data_saida: Any, hora_saida: Any, data_chegada: Any, hora_chegada: Any) -> float:
    dt_saida = parse_data_hora(data_saida, hora_saida)
    dt_chegada = parse_data_hora(data_chegada, hora_chegada)
    if not dt_saida or not dt_chegada:
        return 0.0
    diff = (dt_chegada - dt_saida).total_seconds() / 3600.0
    if diff <= 0:
        return 0.0
    return min(round(diff, 2), 72.0)


def status_value(programacao: ProgramacaoDB) -> str:
    status = upper_text(programacao.status_operacional) or upper_text(programacao.status)
    if not status and safe_int(programacao.finalizada_no_app, 0) == 1:
        return "FINALIZADA"
    return status


def is_em_rota(status: str) -> bool:
    return upper_text(status) in {"EM_ROTA", "EM ROTA", "INICIADA"}


def is_ativa(status: str) -> bool:
    return upper_text(status) in {"ATIVA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"}


def is_finalizada(status: str) -> bool:
    return upper_text(status) in {"FINALIZADA", "FINALIZADO"}


def is_cancelada(status: str) -> bool:
    return upper_text(status) in {"CANCELADA", "CANCELADO"}


def status_match(status: str, filtro: str) -> bool:
    filtro_norm = upper_text(filtro)
    if filtro_norm == "TODOS":
        return True
    if filtro_norm == "ATIVAS":
        return is_ativa(status)
    if filtro_norm in {"FINALIZADA", "FINALIZADO"}:
        return is_finalizada(status)
    if filtro_norm in {"CANCELADA", "CANCELADO"}:
        return is_cancelada(status)
    if filtro_norm in {"EM_ROTA", "EM ROTA"}:
        return is_em_rota(status)
    return upper_text(status) == filtro_norm


def cadastro_ativo_para_escala(status: Any) -> bool:
    status_norm = upper_text(status or "ATIVO")
    return status_norm not in {"INATIVO", "DESATIVADO", "BLOQUEADO", "CANCELADO", "CANCELADA"}


def split_ajudantes(equipe_raw: Any, ajudante_map: dict[str, str]) -> list[str]:
    raw = str(equipe_raw or "").strip()
    if not raw:
        return []
    out = []
    seen = set()
    for part in re.split(r"[|,;/]+", raw):
        token = part.strip()
        if not token:
            continue
        nome = ajudante_map.get(token) if token.isdigit() else None
        nome = nome or upper_text(token)
        if nome in {"NAN", "NONE", "-", "SEM EQUIPE"}:
            continue
        if nome not in seen:
            seen.add(nome)
            out.append(nome)
    return out


def default_stats() -> dict[str, Any]:
    return {
        "codigo": "",
        "pessoa_id": "",
        "rotas": 0,
        "em_rota": 0,
        "ativas": 0,
        "finalizadas": 0,
        "canceladas": 0,
        "kg": 0.0,
        "km_rodado": 0.0,
        "horas_trab": 0.0,
        "mort_aves": 0.0,
        "litros": 0.0,
        "media_km_l_sum": 0.0,
        "media_km_l_count": 0,
        "custo_km_sum": 0.0,
        "custo_km_count": 0,
        "local_ref": "-",
        "last_dt": None,
        "local_counts": {},
        "dias_set": set(),
    }


def add_status_counts(stats: dict[str, Any], status: str) -> None:
    if is_em_rota(status):
        stats["em_rota"] += 1
    if is_ativa(status):
        stats["ativas"] += 1
    if is_finalizada(status):
        stats["finalizadas"] += 1
    if is_cancelada(status):
        stats["canceladas"] += 1


def tag_por_carga(
    rotas: int,
    media_rotas: float,
    km_rodado: float,
    media_km: float,
    horas_trab: float,
    media_horas: float,
    mort_aves: float,
    media_mort: float,
    custo_km: float,
    media_custo_km: float,
) -> str:
    idx_rotas = float(rotas) / media_rotas if media_rotas > 0 else 0.0
    idx_km = float(km_rodado) / media_km if media_km > 0 else 0.0
    idx_horas = float(horas_trab) / media_horas if media_horas > 0 else 0.0
    idx_mort = float(mort_aves) / media_mort if media_mort > 0 else (1.5 if mort_aves > 0 else 0.0)
    idx_custo = float(custo_km) / media_custo_km if media_custo_km > 0 else 0.0
    idx = max(idx_rotas, idx_km, idx_horas, idx_mort * 1.2, idx_custo * 1.1)
    if idx > 1.25:
        return "sobrecarga"
    if idx > 1.05:
        return "alerta"
    return "equilibrada"


def media_custo(stats: dict[str, Any]) -> float:
    count = safe_int(stats.get("custo_km_count"), 0)
    return safe_float(stats.get("custo_km_sum"), 0.0) / float(count) if count > 0 else 0.0


def serialize_pessoas(
    rows: list[tuple[str, dict[str, Any]]],
    qtd_rotas: int,
    *,
    include_local: bool,
    folga_motoristas: dict[str, set[str]] | None = None,
    folga_ajudantes: dict[str, set[str]] | None = None,
) -> list[EscalaPessoa]:
    if not rows:
        return []
    media_rotas = qtd_rotas / float(len(rows)) if rows else 0.0
    media_km = sum(safe_float(d.get("km_rodado"), 0.0) for _n, d in rows) / float(len(rows))
    media_horas = sum(safe_float(d.get("horas_trab"), 0.0) for _n, d in rows) / float(len(rows))
    media_mort = sum(safe_float(d.get("mort_aves"), 0.0) for _n, d in rows) / float(len(rows))
    media_custo_km = sum(media_custo(d) for _n, d in rows) / float(len(rows))
    out = []
    for nome, data in rows:
        carga = tag_por_carga(
            safe_int(data.get("rotas"), 0),
            media_rotas,
            safe_float(data.get("km_rodado"), 0.0),
            media_km,
            safe_float(data.get("horas_trab"), 0.0),
            media_horas,
            safe_float(data.get("mort_aves"), 0.0),
            media_mort,
            media_custo(data),
            media_custo_km,
        )
        out.append(
            EscalaPessoa(
                nome=nome,
                rotas=safe_int(data.get("rotas"), 0),
                em_rota=safe_int(data.get("em_rota"), 0),
                ativas=safe_int(data.get("ativas"), 0),
                finalizadas=safe_int(data.get("finalizadas"), 0),
                canceladas=safe_int(data.get("canceladas"), 0),
                local=str(data.get("local_ref") or "-") if include_local else "-",
                km_rodado=round(safe_float(data.get("km_rodado"), 0.0), 1),
                horas_trab=round(safe_float(data.get("horas_trab"), 0.0), 2),
                carga=carga,
                em_folga=pessoa_em_folga(
                    tipo="MOTORISTA" if include_local else "AJUDANTE",
                    nome=nome,
                    codigo=data.get("codigo"),
                    pessoa_id=data.get("pessoa_id"),
                    folga_motoristas=folga_motoristas,
                    folga_ajudantes=folga_ajudantes,
                ),
            )
        )
    return out


def recomendacoes_distribuicao(
    qtd_rotas: int,
    mot_rows: list[tuple[str, dict[str, Any]]],
    aj_rows: list[tuple[str, dict[str, Any]]],
    *,
    folgas: list[EscalaFolga] | None = None,
) -> str:
    if qtd_rotas <= 0:
        return "Recomendacoes: sem dados no filtro para sugerir distribuicao."
    folgas = folgas or []
    folga_mot, folga_aju = folga_sets(folgas)
    mot_rows_reco = [
        (nome, data)
        for nome, data in mot_rows
        if not pessoa_em_folga(
            tipo="MOTORISTA",
            nome=nome,
            codigo=data.get("codigo"),
            pessoa_id=data.get("pessoa_id"),
            folga_motoristas=folga_mot,
        )
    ]
    aj_rows_reco = [
        (nome, data)
        for nome, data in aj_rows
        if not pessoa_em_folga(
            tipo="AJUDANTE",
            nome=nome,
            pessoa_id=data.get("pessoa_id"),
            folga_ajudantes=folga_aju,
        )
    ]
    recs = []
    local_counts = {}
    for _nome, data in mot_rows_reco:
        for local, qtd in (data.get("local_counts") or {}).items():
            local_counts[local] = safe_int(local_counts.get(local), 0) + safe_int(qtd, 0)
    local_alvo = sorted(local_counts.items(), key=lambda item: (-item[1], item[0]))[0][0] if local_counts else "-"

    if mot_rows_reco:
        media_horas = max(sum(safe_float(d.get("horas_trab"), 0.0) for _n, d in mot_rows_reco) / float(len(mot_rows_reco)), 1.0)

        def score_motorista(data: dict[str, Any]) -> float:
            return (
                safe_float(data.get("rotas"), 0.0)
                + safe_float(data.get("em_rota"), 0.0) * 0.85
                + (safe_float(data.get("horas_trab"), 0.0) / media_horas) * 0.70
            )

        prox_motorista = sorted(mot_rows_reco, key=lambda item: (score_motorista(item[1]), item[0]))[0][0]
        recs.append(f"Proxima rota sugerida (motorista): {prox_motorista}")

        media_rotas = qtd_rotas / float(len(mot_rows_reco))
        sobrecarregados = [
            nome
            for nome, data in mot_rows_reco
            if safe_float(data.get("rotas"), 0.0) > media_rotas * 1.25
            or safe_float(data.get("horas_trab"), 0.0) > media_horas * 1.20
        ]
        if sobrecarregados:
            recs.append("Evitar novas rotas para motoristas sobrecarregados: " + ", ".join(sobrecarregados[:4]))

    if aj_rows_reco:
        media_horas_aj = max(sum(safe_float(d.get("horas_trab"), 0.0) for _n, d in aj_rows_reco) / float(len(aj_rows_reco)), 1.0)

        def score_ajudante(data: dict[str, Any]) -> float:
            return (
                safe_float(data.get("rotas"), 0.0)
                + safe_float(data.get("em_rota"), 0.0) * 0.90
                + (safe_float(data.get("horas_trab"), 0.0) / media_horas_aj) * 0.65
            )

        prox_aj = [nome for nome, _data in sorted(aj_rows_reco, key=lambda item: (score_ajudante(item[1]), item[0]))[:2]]
        if prox_aj:
            recs.append("Ajudantes sugeridos para proxima equipe: " + " / ".join(prox_aj))

    if folgas:
        folga_text = [
            f"{item.tipo}: {item.pessoa_nome} ate {item.data_fim}" for item in folgas[:6]
        ]
        if len(folgas) > len(folga_text):
            folga_text.append(f"+{len(folgas) - len(folga_text)} folga(s)")
        recs.append("Folgas no periodo: " + "; ".join(folga_text))

    if not recs:
        return "Recomendacoes: distribuicao esta equilibrada no filtro atual."
    return f"Recomendacoes (local-alvo: {local_alvo}):\n- " + "\n- ".join(recs)


@router.get("/pessoas", response_model=list[EscalaPessoaOption])
async def escala_pessoas(
    tipo: str = "MOTORISTA",
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    tipo_norm = upper_text(tipo or "MOTORISTA")
    out: list[EscalaPessoaOption] = []
    if tipo_norm == "AJUDANTE":
        result = await db.execute(select(AjudanteDB).order_by(AjudanteDB.nome.asc(), AjudanteDB.sobrenome.asc()))
        for item in result.scalars().all():
            if not cadastro_ativo_para_escala(item.status):
                continue
            nome = upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
            if not nome:
                continue
            out.append(
                EscalaPessoaOption(
                    tipo="AJUDANTE",
                    pessoa_id=str(item.id or ""),
                    pessoa_codigo="",
                    pessoa_nome=nome,
                    label=nome,
                )
            )
        return out

    result = await db.execute(select(MotoristaDB).order_by(MotoristaDB.nome.asc(), MotoristaDB.codigo.asc()))
    for item in result.scalars().all():
        if not cadastro_ativo_para_escala(item.status):
            continue
        nome = upper_text(item.nome)
        codigo = upper_text(item.codigo)
        if not nome:
            continue
        out.append(
            EscalaPessoaOption(
                tipo="MOTORISTA",
                pessoa_id=str(item.id or ""),
                pessoa_codigo=codigo,
                pessoa_nome=nome,
                label=f"{nome} ({codigo})" if codigo else nome,
            )
        )
    return out


@router.get("/folgas", response_model=list[EscalaFolga])
async def escala_folgas(
    periodo: str = "30",
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    ref_inicio, ref_fim = periodo_ref_range(periodo)
    return await fetch_folgas_ativas(db, ref_inicio, ref_fim)


@router.post("/folgas", response_model=EscalaFolga)
async def criar_escala_folga(
    payload: EscalaFolgaCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    tipo = upper_text(payload.tipo)
    if tipo not in {"MOTORISTA", "AJUDANTE"}:
        raise HTTPException(status_code=400, detail="Tipo de folga invalido.")
    nome = upper_text(payload.pessoa_nome)
    if not nome:
        raise HTTPException(status_code=400, detail="Informe a pessoa da folga.")
    inicio_dt = parse_date_safe(payload.data_inicio)
    fim_dt = parse_date_safe(payload.data_fim) or inicio_dt
    if not inicio_dt or not fim_dt:
        raise HTTPException(status_code=400, detail="Informe datas validas no formato YYYY-MM-DD.")
    if fim_dt < inicio_dt:
        raise HTTPException(status_code=400, detail="A data final da folga nao pode ser menor que a inicial.")

    pessoa_id = str(payload.pessoa_id or "").strip()
    pessoa_codigo = upper_text(payload.pessoa_codigo)
    if tipo == "MOTORISTA":
        stmt = select(MotoristaDB)
        if pessoa_id.isdigit():
            stmt = stmt.where(MotoristaDB.id == safe_int(pessoa_id, 0))
        elif pessoa_codigo:
            stmt = stmt.where(func.upper(func.coalesce(MotoristaDB.codigo, "")) == pessoa_codigo)
        else:
            stmt = stmt.where(func.upper(func.coalesce(MotoristaDB.nome, "")) == nome)
        cadastro_result = await db.execute(stmt.limit(1))
        cadastro = cadastro_result.scalar_one_or_none()
        if not cadastro or not cadastro_ativo_para_escala(cadastro.status):
            raise HTTPException(status_code=422, detail="Selecione um motorista ativo do cadastro.")
        pessoa_id = str(cadastro.id or "")
        pessoa_codigo = upper_text(cadastro.codigo)
        nome = upper_text(cadastro.nome)
    else:
        if pessoa_id.isdigit():
            cadastro_result = await db.execute(select(AjudanteDB).where(AjudanteDB.id == safe_int(pessoa_id, 0)).limit(1))
            cadastro = cadastro_result.scalar_one_or_none()
        else:
            cadastro_result = await db.execute(select(AjudanteDB).order_by(AjudanteDB.nome.asc(), AjudanteDB.sobrenome.asc()))
            cadastro = None
            for item in cadastro_result.scalars().all():
                display = upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
                if display == nome:
                    cadastro = item
                    break
        if not cadastro or not cadastro_ativo_para_escala(cadastro.status):
            raise HTTPException(status_code=422, detail="Selecione um ajudante ativo do cadastro.")
        pessoa_id = str(cadastro.id or "")
        pessoa_codigo = ""
        nome = upper_text(f"{cadastro.nome or ''} {cadastro.sobrenome or ''}".strip())

    result = await db.execute(
        select(EscalaFolgaDB).where(
            func.upper(func.coalesce(EscalaFolgaDB.status, "ATIVA")) == "ATIVA",
            func.upper(func.coalesce(EscalaFolgaDB.tipo, "")) == tipo,
        )
    )
    for row in result.scalars().all():
        mesmo_recurso = False
        if tipo == "MOTORISTA":
            mesmo_recurso = (pessoa_codigo and pessoa_codigo == upper_text(row.pessoa_codigo)) or nome == upper_text(row.pessoa_nome)
        else:
            mesmo_recurso = (pessoa_id and pessoa_id == str(row.pessoa_id or "").strip()) or nome == upper_text(row.pessoa_nome)
        if mesmo_recurso and date_ranges_overlap(row.data_inicio, row.data_fim, inicio_dt.isoformat(), fim_dt.isoformat()):
            raise HTTPException(status_code=409, detail="Ja existe folga ativa para esta pessoa no periodo informado.")

    row = EscalaFolgaDB(
        tipo=tipo,
        pessoa_id=pessoa_id,
        pessoa_codigo=pessoa_codigo,
        pessoa_nome=nome,
        data_inicio=inicio_dt.isoformat(),
        data_fim=fim_dt.isoformat(),
        motivo=str(payload.motivo or "").strip(),
        status="ATIVA",
        criado_em=datetime.now().isoformat(timespec="seconds"),
        atualizado_em=datetime.now().isoformat(timespec="seconds"),
    )
    db.add(row)
    await db.commit()
    await db.refresh(row)
    return serialize_folga(row)


@router.patch("/folgas/{folga_id}/encerrar", response_model=EscalaFolga)
async def encerrar_folga(
    folga_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    row = await db.get(EscalaFolgaDB, folga_id)
    if not row:
        raise HTTPException(status_code=404, detail="Folga nao encontrada.")
    if upper_text(row.status or "ATIVA") != "ATIVA":
        raise HTTPException(status_code=409, detail="Esta folga ja esta encerrada.")
    row.status = "ENCERRADA"
    row.atualizado_em = datetime.now().isoformat(timespec="seconds")
    await db.commit()
    await db.refresh(row)
    return serialize_folga(row)


@router.get("/resumo", response_model=EscalaResumoResponse)
async def resumo_escala(
    periodo: str = "30",
    status: str = "ATIVAS",
    limit: int = 5000,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    periodo_norm = upper_text(periodo or "30")
    status_norm = upper_text(status or "ATIVAS")
    if periodo_norm not in {"7", "15", "30", "60", "90", "180", "TODAS"}:
        periodo_norm = "30"
    if status_norm not in {"ATIVAS", "TODOS", "ATIVA", "EM_ROTA", "CARREGADA", "FINALIZADA", "CANCELADA"}:
        status_norm = "ATIVAS"

    cutoff = None
    if periodo_norm != "TODAS":
        cutoff = datetime.now() - timedelta(days=safe_int(periodo_norm, 30))
    ref_inicio, ref_fim = periodo_ref_range(periodo_norm)
    folgas = await fetch_folgas_ativas(db, ref_inicio, ref_fim)
    folga_mot, folga_aju = folga_sets(folgas)

    ajudantes_result = await db.execute(select(AjudanteDB).order_by(AjudanteDB.id.asc()))
    ajudante_map = {
        str(item.id): upper_text(f"{item.nome or ''} {item.sobrenome or ''}".strip())
        for item in ajudantes_result.scalars().all()
    }
    ajudante_id_por_nome = {nome: ajudante_id for ajudante_id, nome in ajudante_map.items() if nome}

    stmt = select(ProgramacaoDB).order_by(ProgramacaoDB.id.desc()).limit(max(min(limit, 10000), 1))
    result = await db.execute(stmt)
    programacoes = list(result.scalars().all())

    mot: dict[str, dict[str, Any]] = {}
    aju: dict[str, dict[str, Any]] = {}
    qtd_rotas = 0

    for programacao in programacoes:
        status_programacao = status_value(programacao)
        if not status_match(status_programacao, status_norm):
            continue

        data_ref = programacao.data_saida or programacao.data_criacao or programacao.data or ""
        dt_prog = parse_data_programacao(data_ref)
        if cutoff is not None and dt_prog is not None and dt_prog < cutoff:
            continue

        qtd_rotas += 1
        motorista = upper_text(programacao.motorista) or "SEM MOTORISTA"
        motorista_codigo = upper_text(
            getattr(programacao, "motorista_codigo", "")
            or getattr(programacao, "codigo_motorista", "")
        )
        motorista_id = str(getattr(programacao, "motorista_id", "") or "").strip()
        local_rota = upper_text(programacao.local_rota or programacao.tipo_rota) or "-"
        horas_trab = calc_horas_trabalhadas(
            programacao.data_saida,
            programacao.hora_saida,
            programacao.data_chegada,
            programacao.hora_chegada,
        )
        km_inicial = safe_float(programacao.km_inicial, 0.0)
        km_final = safe_float(programacao.km_final, 0.0)
        km_rodado = safe_float(programacao.km_rodado, 0.0)
        if km_rodado <= 0 and km_final > km_inicial:
            km_rodado = km_final - km_inicial
        litros = safe_float(programacao.litros, 0.0)
        media_km_l = safe_float(programacao.media_km_l, 0.0)
        if media_km_l <= 0 and litros > 0 and km_rodado > 0:
            media_km_l = km_rodado / litros
        custo_km = safe_float(programacao.custo_km, 0.0)
        mort_aves = safe_float(programacao.mortalidade_transbordo_aves, 0.0)

        motorista_stats = mot.setdefault(motorista, default_stats())
        if motorista_codigo and not motorista_stats.get("codigo"):
            motorista_stats["codigo"] = motorista_codigo
        if motorista_id and not motorista_stats.get("pessoa_id"):
            motorista_stats["pessoa_id"] = motorista_id
        motorista_stats["rotas"] += 1
        motorista_stats["kg"] += safe_float(programacao.kg_estimado, 0.0)
        motorista_stats["km_rodado"] += km_rodado
        motorista_stats["horas_trab"] += horas_trab
        motorista_stats["mort_aves"] += mort_aves
        motorista_stats["litros"] += litros
        if media_km_l > 0:
            motorista_stats["media_km_l_sum"] += media_km_l
            motorista_stats["media_km_l_count"] += 1
        if custo_km > 0:
            motorista_stats["custo_km_sum"] += custo_km
            motorista_stats["custo_km_count"] += 1
        add_status_counts(motorista_stats, status_programacao)

        for dt_item in (
            dt_prog,
            parse_data_hora(programacao.data_saida, programacao.hora_saida),
            parse_data_hora(programacao.data_chegada, programacao.hora_chegada),
        ):
            if isinstance(dt_item, datetime):
                motorista_stats["dias_set"].add(dt_item.date().isoformat())
                if motorista_stats.get("last_dt") is None or dt_item > motorista_stats["last_dt"]:
                    motorista_stats["last_dt"] = dt_item
        if local_rota != "-":
            counts = motorista_stats["local_counts"]
            counts[local_rota] = safe_int(counts.get(local_rota), 0) + 1
            motorista_stats["local_ref"] = sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0]

        for nome_ajudante in split_ajudantes(programacao.equipe, ajudante_map):
            ajudante_stats = aju.setdefault(nome_ajudante, default_stats())
            ajudante_id = ajudante_id_por_nome.get(nome_ajudante)
            if ajudante_id:
                ajudante_stats["pessoa_id"] = ajudante_id
            ajudante_stats["rotas"] += 1
            ajudante_stats["km_rodado"] += km_rodado
            ajudante_stats["horas_trab"] += horas_trab
            ajudante_stats["mort_aves"] += mort_aves
            ajudante_stats["litros"] += litros
            if media_km_l > 0:
                ajudante_stats["media_km_l_sum"] += media_km_l
                ajudante_stats["media_km_l_count"] += 1
            if custo_km > 0:
                ajudante_stats["custo_km_sum"] += custo_km
                ajudante_stats["custo_km_count"] += 1
            add_status_counts(ajudante_stats, status_programacao)
            if isinstance(dt_prog, datetime):
                ajudante_stats["dias_set"].add(dt_prog.date().isoformat())
                if ajudante_stats.get("last_dt") is None or dt_prog > ajudante_stats["last_dt"]:
                    ajudante_stats["last_dt"] = dt_prog

    mot_rows = sorted(mot.items(), key=lambda item: (-safe_int(item[1].get("rotas"), 0), item[0]))
    aj_rows = sorted(aju.items(), key=lambda item: (-safe_int(item[1].get("rotas"), 0), item[0]))

    qtd_motoristas = len(mot_rows)
    qtd_ajudantes = len(aj_rows)
    total_km = sum(safe_float(data.get("km_rodado"), 0.0) for _nome, data in mot_rows)
    total_horas = sum(safe_float(data.get("horas_trab"), 0.0) for _nome, data in mot_rows)
    total_mort = sum(safe_float(data.get("mort_aves"), 0.0) for _nome, data in mot_rows)
    total_litros = sum(safe_float(data.get("litros"), 0.0) for _nome, data in mot_rows)
    total_media_km_l_sum = sum(safe_float(data.get("media_km_l_sum"), 0.0) for _nome, data in mot_rows)
    total_media_km_l_count = sum(safe_int(data.get("media_km_l_count"), 0) for _nome, data in mot_rows)

    km_medio = total_km / float(qtd_motoristas) if qtd_motoristas else 0.0
    horas_medias = total_horas / float(qtd_motoristas) if qtd_motoristas else 0.0
    mort_media = total_mort / float(qtd_rotas) if qtd_rotas else 0.0
    media_km_l = 0.0
    if total_media_km_l_count > 0:
        media_km_l = total_media_km_l_sum / float(total_media_km_l_count)
    elif total_litros > 0 and total_km > 0:
        media_km_l = total_km / total_litros

    if qtd_motoristas:
        media_rotas = qtd_rotas / float(qtd_motoristas)
        mais_carregado = mot_rows[0][0]
        carga_max = safe_float(mot_rows[0][1].get("rotas"), 0.0)
        nivel = "EQUILIBRADA"
        if media_rotas > 0 and carga_max > media_rotas * 1.25:
            nivel = "SOBRECARGA"
        elif media_rotas > 0 and carga_max > media_rotas * 1.05:
            nivel = "ALERTA"
        resumo = (
            f"Nivel da escala: {nivel} | Motorista mais carregado: {mais_carregado}\n"
            f"Rotas no filtro: {qtd_rotas} | Motoristas: {qtd_motoristas} | Ajudantes: {qtd_ajudantes}\n"
            f"KM total: {total_km:.1f} | KM medio/motorista: {km_medio:.1f}\n"
            f"Horas totais: {total_horas:.2f} | Horas medias/motorista: {horas_medias:.2f}\n"
            f"Ocorrencias media/rota: {mort_media:.2f} unid. | KM/L medio: {media_km_l:.2f}\n"
            "Legenda visual: verde=equilibrado | laranja=alerta | vermelho=sobrecarga"
        )
    else:
        resumo = f"Rotas no filtro: {qtd_rotas} | Sem motoristas no periodo/filtro selecionado."

    chart = [
        EscalaChartItem(
            nome=nome,
            horas=round(safe_float(data.get("horas_trab"), 0.0), 2),
            dias=len(data.get("dias_set") or set()),
        )
        for nome, data in sorted(
            mot_rows,
            key=lambda item: (-safe_float(item[1].get("horas_trab"), 0.0), -len(item[1].get("dias_set") or set()), item[0]),
        )[:8]
    ]

    return EscalaResumoResponse(
        periodo=periodo_norm,
        status=status_norm,
        kpis=EscalaKpis(
            rotas=qtd_rotas,
            motoristas=qtd_motoristas,
            ajudantes=qtd_ajudantes,
            folgas_motoristas=count_folgas_pessoas(folgas, "MOTORISTA"),
            folgas_ajudantes=count_folgas_pessoas(folgas, "AJUDANTE"),
            mortalidade_media=round(mort_media, 2),
            km_total=round(total_km, 1),
            km_medio_motorista=round(km_medio, 1),
            media_km_l=round(media_km_l, 2),
            horas_medias_motorista=round(horas_medias, 2),
        ),
        resumo=resumo,
        recomendacoes=recomendacoes_distribuicao(qtd_rotas, mot_rows, aj_rows, folgas=folgas),
        motoristas=serialize_pessoas(
            mot_rows,
            qtd_rotas,
            include_local=True,
            folga_motoristas=folga_mot,
            folga_ajudantes=folga_aju,
        ),
        ajudantes=serialize_pessoas(
            aj_rows,
            qtd_rotas,
            include_local=False,
            folga_motoristas=folga_mot,
            folga_ajudantes=folga_aju,
        ),
        chart=chart,
        folgas=folgas,
    )


@router.get("/pdf")
async def escala_pdf(
    periodo: str = "30",
    status: str = "ATIVAS",
    limit: int = 5000,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    data = await resumo_escala(periodo=periodo, status=status, limit=limit, db=db, current_user=current_user)
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:  # pragma: no cover - depends on optional package
        raise HTTPException(status_code=503, detail="Biblioteca ReportLab indisponivel.") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    draw_escala_pdf(pdf, data)
    pdf.save()
    buffer.seek(0)
    safe_periodo = upper_text(data.periodo).replace("/", "_") or "PERIODO"
    safe_status = upper_text(data.status).replace("/", "_") or "STATUS"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="ESCALA_{safe_periodo}_{safe_status}.pdf"'},
    )
