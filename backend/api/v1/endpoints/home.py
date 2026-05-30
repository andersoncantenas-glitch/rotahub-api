"""
Home dashboard endpoints mirroring the desktop HomePage operational overview.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy import distinct, exists, func, not_, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import safe_float, safe_int
from backend.config.database import get_db
from backend.config.settings import settings
from backend.models.cadastro import ClienteDB
from backend.models.despesa import DespesaDB
from backend.models.programacao import ProgramacaoDB, ProgramacaoItemControleDB, ProgramacaoItemDB
from backend.models.recebimento import RecebimentoDB
from backend.models.user import User
from backend.models.venda_importada import VendaImportadaDB
from backend.services.auth import get_current_user

router = APIRouter()

BLOCKED_STATUSES = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}
ROOT_DIR = Path(__file__).resolve().parents[4]


def clean_text(value: Any) -> str:
    return str(value or "").strip()


def upper_text(value: Any) -> str:
    return clean_text(value).upper()


def optional_float(value: Any) -> float | None:
    if value is None:
        return None
    text_value = clean_text(value).replace(",", ".")
    if not text_value:
        return None
    try:
        return float(text_value)
    except (TypeError, ValueError):
        return None


def parse_json_dict(value: Any) -> dict[str, Any]:
    try:
        data = json.loads(str(value or "{}"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def active_programacao_filters():
    status = func.upper(func.trim(func.coalesce(ProgramacaoDB.status, "")))
    operational = func.upper(func.trim(func.coalesce(ProgramacaoDB.status_operacional, "")))
    return (
        status.not_in(BLOCKED_STATUSES),
        operational.not_in(BLOCKED_STATUSES),
        func.coalesce(ProgramacaoDB.finalizada_no_app, 0) == 0,
        func.trim(func.coalesce(ProgramacaoDB.data_chegada, "")) == "",
        func.trim(func.coalesce(ProgramacaoDB.hora_chegada, "")) == "",
        func.coalesce(ProgramacaoDB.km_final, 0) == 0,
    )


def route_status(programacao: ProgramacaoDB) -> str:
    return upper_text(programacao.status_operacional or programacao.status or "ATIVA")


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
    data = parse_json_dict(value)
    return upper_text(data.get("carga_raiz_programacao") or data.get("carga_origem_programacao") or fallback)


async def kg_transferido_convertido_por_destino(db: AsyncSession, codigos: list[str]) -> dict[str, float]:
    codigos_norm = {upper_text(codigo) for codigo in codigos if upper_text(codigo)}
    if not codigos_norm:
        return {}
    try:
        result = await db.execute(
            text(
                """
                SELECT codigo_origem, codigo_destino, qtd_caixas, qtd_convertida, status, snapshot
                  FROM transferencias
                 WHERE TRIM(COALESCE(codigo_destino, '')) <> ''
                """
            )
        )
    except Exception:
        return {}
    transferencias = [dict(row) for row in result.mappings().all()]
    roots = {
        carga_raiz_from_snapshot(row.get("snapshot"), row.get("codigo_origem"))
        for row in transferencias
        if upper_text(row.get("codigo_destino")) in codigos_norm
    }
    root_map: dict[str, ProgramacaoDB] = {}
    if roots:
        root_result = await db.execute(select(ProgramacaoDB).where(func.upper(ProgramacaoDB.codigo_programacao).in_(roots)))
        root_map = {upper_text(item.codigo_programacao): item for item in root_result.scalars().all()}
    out: dict[str, float] = {}
    for row in transferencias:
        destino = upper_text(row.get("codigo_destino"))
        if destino not in codigos_norm:
            continue
        status_value = upper_text(row.get("status"))
        if status_value in {"CANCELADA", "CANCELADO", "RECUSADA", "RECUSADO"}:
            continue
        qtd = transferencia_qtd_convertida(row)
        if qtd <= 0:
            continue
        origem = upper_text(row.get("codigo_origem"))
        raiz = carga_raiz_from_snapshot(row.get("snapshot"), origem)
        root = root_map.get(raiz)
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
        if kg_base <= 0 or caixas_base <= 0:
            continue
        out[destino] = round(out.get(destino, 0.0) + (qtd * (kg_base / caixas_base)), 2)
    return out


def serialize_programacao(programacao: ProgramacaoDB, kg_transferido_convertido: float = 0.0) -> dict[str, Any]:
    data_saida = clean_text(programacao.saida_data or programacao.data_saida)
    hora_saida = clean_text(programacao.saida_hora or programacao.hora_saida)
    nf_caixas = safe_int(programacao.nf_caixas, 0)
    caixas_carregadas = safe_int(programacao.caixas_carregadas or programacao.qnt_cx_carregada or nf_caixas, 0)
    nf_kg_carregado = safe_float(programacao.nf_kg_carregado or programacao.kg_carregado, 0.0)
    nf_kg = safe_float(programacao.nf_kg, 0.0)
    nf_saldo = safe_float(programacao.nf_saldo, 0.0)
    if abs(nf_saldo) < 0.0001 and nf_kg > 0 and nf_kg_carregado > 0:
        nf_saldo = round(max(nf_kg - nf_kg_carregado, 0.0), 2)
    nf_preco = safe_float(programacao.nf_preco, 0.0)
    nf_saldo_valor = round(max(nf_saldo, 0.0) * nf_preco, 2)
    nf_kg_vendido = safe_float(programacao.nf_kg_vendido, 0.0) or safe_float(kg_transferido_convertido, 0.0)
    return {
        "id": programacao.id,
        "codigo_programacao": upper_text(programacao.codigo_programacao),
        "programacao_id": programacao.id,
        "codigo_exibicao": upper_text(programacao.codigo_programacao) or f"ID {programacao.id}",
        "data": clean_text(programacao.data_criacao or programacao.data),
        "motorista": upper_text(programacao.motorista),
        "motorista_codigo": upper_text(programacao.motorista_codigo or programacao.codigo_motorista),
        "veiculo": upper_text(programacao.veiculo),
        "equipe": clean_text(programacao.equipe),
        "rota": upper_text(programacao.local_rota or programacao.tipo_rota),
        "local_carregamento": clean_text(
            programacao.local_carregamento
            or programacao.granja_carregada
            or programacao.local_carregado
            or programacao.local_carreg
        ),
        "status": route_status(programacao),
        "prestacao_status": upper_text(programacao.prestacao_status or "PENDENTE"),
        "total_caixas": safe_int(programacao.total_caixas, 0),
        "quilos": safe_float(programacao.quilos or programacao.kg_estimado, 0.0),
        "adiantamento": safe_float(programacao.adiantamento or programacao.adiantamento_rota, 0.0),
        "num_nf": clean_text(programacao.nf_numero or programacao.num_nf),
        "data_saida": data_saida,
        "hora_saida": hora_saida,
        "saida_data": data_saida,
        "saida_hora": hora_saida,
        "inicio_carregamento": clean_text(programacao.inicio_carregamento),
        "fim_carregamento": clean_text(programacao.fim_carregamento),
        "data_chegada": clean_text(programacao.data_chegada),
        "hora_chegada": clean_text(programacao.hora_chegada),
        "km_inicial": safe_float(programacao.km_inicial, 0.0),
        "km_final": safe_float(programacao.km_final, 0.0),
        "km_rodado": safe_float(programacao.km_rodado, 0.0),
        "nf_kg": nf_kg,
        "nf_preco": nf_preco,
        "nf_caixas": nf_caixas,
        "caixas_carregadas": caixas_carregadas,
        "qnt_cx_carregada": safe_int(programacao.qnt_cx_carregada, 0),
        "kg_carregado": nf_kg_carregado,
        "nf_kg_carregado": nf_kg_carregado,
        "nf_kg_vendido": nf_kg_vendido,
        "kg_transferido_convertido": safe_float(kg_transferido_convertido, 0.0),
        "nf_saldo": nf_saldo,
        "nf_saldo_valor": nf_saldo_valor,
        "desconto_fornecedor": nf_saldo_valor,
        "media": safe_float(programacao.media, 0.0),
        "qnt_aves_por_cx": safe_int(programacao.qnt_aves_por_cx, 0),
        "aves_caixa_final": safe_int(programacao.aves_caixa_final or programacao.qnt_aves_caixa_final, 0),
        "qnt_aves_caixa_final": safe_int(programacao.qnt_aves_caixa_final, 0),
    }


def read_local_version() -> str:
    try:
        from version import APP_VERSION

        return str(APP_VERSION or "").strip() or "-"
    except Exception:
        return "-"


def read_manifest() -> dict[str, Any]:
    manifest_path = ROOT_DIR / "updates" / "version.json"
    try:
        return json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}
    except Exception:
        return {}


def database_label() -> str:
    url = settings.DATABASE_URL
    if ":///" not in url:
        return url.split("://", 1)[0]
    path = url.split(":///", 1)[1]
    return str((ROOT_DIR / path).resolve()) if path.startswith("./") else path


async def count_scalar(db: AsyncSession, stmt) -> int:
    result = await db.execute(stmt)
    return safe_int(result.scalar_one_or_none(), 0)


@router.get("/overview")
async def home_overview(
    limit: int = 30,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    del current_user
    limit = max(1, min(limit, 100))
    active_filters = active_programacao_filters()

    rotas_stmt = (
        select(ProgramacaoDB)
        .where(*active_filters)
        .order_by(ProgramacaoDB.id.desc())
        .limit(limit)
    )
    rotas_result = await db.execute(rotas_stmt)
    rotas = rotas_result.scalars().all()
    kg_transferido_por_destino = await kg_transferido_convertido_por_destino(
        db,
        [programacao.codigo_programacao for programacao in rotas],
    )

    active_count = await count_scalar(
        db,
        select(func.count(ProgramacaoDB.id)).where(*active_filters),
    )
    vendas_count = await count_scalar(db, select(func.count(VendaImportadaDB.id)))
    clientes_ativos = await count_scalar(
        db,
        select(func.count(distinct(ProgramacaoItemDB.cod_cliente)))
        .select_from(ProgramacaoItemDB)
        .join(ProgramacaoDB, ProgramacaoDB.codigo_programacao == ProgramacaoItemDB.codigo_programacao)
        .where(*active_filters, func.trim(func.coalesce(ProgramacaoItemDB.cod_cliente, "")) != ""),
    )
    if clientes_ativos == 0:
        clientes_ativos = await count_scalar(db, select(func.count(ClienteDB.id)))

    prestacoes_pendentes = await count_scalar(
        db,
        select(func.count(ProgramacaoDB.id)).where(
            func.upper(func.trim(func.coalesce(ProgramacaoDB.status, ""))).not_in({"CANCELADA", "CANCELADO"}),
            func.upper(func.trim(func.coalesce(ProgramacaoDB.prestacao_status, "PENDENTE"))) != "FECHADA",
        ),
    )
    despesa_exists = exists(
        select(DespesaDB.id).where(DespesaDB.codigo_programacao == ProgramacaoDB.codigo_programacao)
    ).correlate(ProgramacaoDB)
    sem_despesa = await count_scalar(
        db,
        select(func.count(ProgramacaoDB.id)).where(*active_filters, not_(despesa_exists)),
    )

    manifest = read_manifest()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "generated_at": now,
        "metrics": {
            "programacoes_ativas": active_count,
            "vendas_importadas": vendas_count,
            "clientes_ativos": clientes_ativos,
        },
        "pendencias": {
            "rotas_abertas": active_count,
            "prestacoes_pendentes": prestacoes_pendentes,
            "sem_despesa": sem_despesa,
        },
        "rotas": [
            serialize_programacao(
                programacao,
                kg_transferido_por_destino.get(upper_text(programacao.codigo_programacao), 0.0),
            )
            for programacao in rotas
        ],
        "sistema": {
            "versao_local": read_local_version(),
            "versao_disponivel": clean_text(manifest.get("version")) or "-",
            "banco": database_label(),
            "ambiente": settings.ENVIRONMENT,
            "debug": settings.DEBUG,
            "api": "online",
            "data_hora_atual": now,
            "alerta": clean_text(manifest.get("alert")) or "",
        },
    }


@router.get("/rotas/{codigo_programacao}/preview")
async def home_rota_preview(
    codigo_programacao: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    del current_user
    codigo = upper_text(codigo_programacao)
    stmt = select(ProgramacaoDB).where(func.upper(func.trim(ProgramacaoDB.codigo_programacao)) == codigo)
    if codigo.isdigit():
        stmt = select(ProgramacaoDB).where(
            (func.upper(func.trim(func.coalesce(ProgramacaoDB.codigo_programacao, ""))) == codigo)
            | (ProgramacaoDB.id == int(codigo))
        )
    result = await db.execute(stmt)
    programacao = result.scalar_one_or_none()
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")

    items_result = await db.execute(
        select(ProgramacaoItemDB)
        .where(ProgramacaoItemDB.codigo_programacao == programacao.codigo_programacao)
        .order_by(ProgramacaoItemDB.id.asc())
    )
    itens = items_result.scalars().all()

    controles_result = await db.execute(
        select(ProgramacaoItemControleDB).where(
            ProgramacaoItemControleDB.codigo_programacao == programacao.codigo_programacao
        )
    )
    controles = controles_result.scalars().all()
    controles_por_chave: dict[tuple[str, str], ProgramacaoItemControleDB] = {}
    controles_por_cliente: dict[str, ProgramacaoItemControleDB] = {}
    for controle in controles:
        cod_cliente = upper_text(controle.cod_cliente)
        pedido = upper_text(controle.pedido)
        if cod_cliente:
            controles_por_cliente.setdefault(cod_cliente, controle)
        if cod_cliente or pedido:
            controles_por_chave[(cod_cliente, pedido)] = controle

    recebimentos_result = await db.execute(
        select(RecebimentoDB)
        .where(RecebimentoDB.codigo_programacao == programacao.codigo_programacao)
        .order_by(RecebimentoDB.id.asc())
    )
    recebimentos = recebimentos_result.scalars().all()

    despesas_result = await db.execute(
        select(DespesaDB)
        .where(DespesaDB.codigo_programacao == programacao.codigo_programacao)
        .order_by(DespesaDB.id.asc())
    )
    despesas = despesas_result.scalars().all()

    item_rows = []
    total_caixas = 0
    total_caixas_programadas = 0
    total_caixas_entregues = 0
    total_kg = 0.0
    total_controle = 0.0
    total_entregues = 0
    total_pendentes = 0
    total_alterados = 0
    total_cancelados = 0
    total_com_localizacao = 0
    for item in itens:
        key = (upper_text(item.cod_cliente), upper_text(item.pedido))
        controle = controles_por_chave.get(key) or controles_por_cliente.get(upper_text(item.cod_cliente))
        caixas_original = safe_int(item.qnt_caixas, 0)
        preco_original = safe_float(item.preco, 0.0)
        caixas = safe_int(controle.caixas_atual if controle and controle.caixas_atual is not None else item.qnt_caixas, 0)
        kg = safe_float(item.kg, 0.0)
        preco_atual = safe_float(
            controle.preco_atual if controle and controle.preco_atual is not None else item.preco,
            0.0,
        )
        valor_recebido = safe_float(controle.valor_recebido if controle else 0, 0.0)
        status_pedido = (
            upper_text(controle.status_pedido if controle and controle.status_pedido else item.status_pedido)
            or "PENDENTE"
        )
        lat_evento = optional_float(
            (controle.lat_entrega if controle and controle.lat_entrega is not None else None)
            if controle
            else None
        )
        if lat_evento is None:
            lat_evento = optional_float(controle.lat_evento if controle else None)
        lon_evento = optional_float(
            (controle.lon_entrega if controle and controle.lon_entrega is not None else None)
            if controle
            else None
        )
        if lon_evento is None:
            lon_evento = optional_float(controle.lon_evento if controle else None)
        tem_localizacao = lat_evento is not None and lon_evento is not None
        if status_pedido in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO"}:
            total_entregues += 1
            total_caixas_entregues += max(caixas if caixas > 0 else caixas_original, 0)
        elif status_pedido in {"CANCELADO", "CANCELADA"}:
            total_cancelados += 1
        elif status_pedido == "ALTERADO":
            total_alterados += 1
        else:
            total_pendentes += 1
        if tem_localizacao:
            total_com_localizacao += 1

        endereco_evento = clean_text(controle.endereco_evento if controle else "")
        cidade_evento = clean_text(controle.cidade_evento if controle else "")
        bairro_evento = clean_text(controle.bairro_evento if controle else "")
        alterado_em = clean_text(
            (controle.alterado_em if controle and controle.alterado_em else "")
            or item.alterado_em
            or getattr(controle, "updated_at", "")
        )
        total_caixas += caixas
        total_caixas_programadas += caixas_original
        total_kg += kg
        total_controle += valor_recebido
        item_rows.append(
            {
                "cod_cliente": upper_text(item.cod_cliente),
                "nome_cliente": clean_text(item.nome_cliente),
                "pedido": clean_text(item.pedido),
                "produto": clean_text(item.produto),
                "endereco": clean_text(item.endereco),
                "vendedor": clean_text(item.vendedor),
                "caixas_original": caixas_original,
                "caixas": caixas,
                "caixas_atual": caixas,
                "delta_caixas": caixas - caixas_original,
                "kg": kg,
                "preco_original": preco_original,
                "preco": preco_atual,
                "preco_atual": preco_atual,
                "delta_preco": preco_atual - preco_original,
                "valor_original": caixas_original * preco_original,
                "valor_atual": caixas * preco_atual,
                "status_pedido": status_pedido,
                "alterado_em": alterado_em,
                "alterado_por": clean_text(controle.alterado_por if controle else item.alterado_por),
                "alteracao_tipo": clean_text(getattr(controle, "alteracao_tipo", "") if controle else ""),
                "alteracao_detalhe": clean_text(getattr(controle, "alteracao_detalhe", "") if controle else ""),
                "valor_recebido": valor_recebido,
                "forma_recebimento": clean_text(controle.forma_recebimento if controle else ""),
                "obs_recebimento": clean_text(controle.obs_recebimento if controle else ""),
                "mortalidade_aves": safe_int(controle.mortalidade_aves if controle else 0, 0),
                "media_aplicada": optional_float(getattr(controle, "media_aplicada", None) if controle else None),
                "peso_previsto": safe_float(controle.peso_previsto if controle else 0, 0.0),
                "ordem_sugerida": safe_int(controle.ordem_sugerida if controle else item.ordem_sugerida, 0),
                "distancia": safe_float(
                    controle.distancia if controle and controle.distancia is not None else item.distancia,
                    0.0,
                ),
                "eta": clean_text(controle.eta if controle and controle.eta else item.eta),
                "confianca_localizacao": optional_float(
                    controle.confianca_localizacao
                    if controle and controle.confianca_localizacao is not None
                    else item.confianca_localizacao
                ),
                "lat_evento": lat_evento,
                "lon_evento": lon_evento,
                "lat_entrega": optional_float(getattr(controle, "lat_entrega", None) if controle else None),
                "lon_entrega": optional_float(getattr(controle, "lon_entrega", None) if controle else None),
                "accuracy_entrega": optional_float(getattr(controle, "accuracy_entrega", None) if controle else None),
                "timestamp_entrega": clean_text(getattr(controle, "timestamp_entrega", "") if controle else ""),
                "latitude": lat_evento,
                "longitude": lon_evento,
                "endereco_evento": endereco_evento,
                "cidade_evento": cidade_evento,
                "bairro_evento": bairro_evento,
                "localizacao": endereco_evento or clean_text(item.endereco),
                "tem_localizacao": tem_localizacao,
                "map_url": f"https://www.google.com/maps?q={lat_evento},{lon_evento}" if tem_localizacao else "",
                "foto_mortalidade": parse_json_dict(getattr(controle, "foto_mortalidade_ref_json", "") if controle else ""),
                "foto_mortalidade_path": clean_text(getattr(controle, "foto_mortalidade_path", "") if controle else ""),
                "mortalidade_foto_path": clean_text(getattr(controle, "mortalidade_foto_path", "") if controle else ""),
                "status_origem": "APP MOTORISTA" if controle else "PROGRAMACAO",
            }
        )

    recebimento_rows = [
        {
            "cod_cliente": upper_text(item.cod_cliente),
            "nome_cliente": clean_text(item.nome_cliente),
            "valor": safe_float(item.valor, 0.0),
            "forma_pagamento": upper_text(item.forma_pagamento),
            "num_nf": clean_text(item.num_nf),
            "data_registro": clean_text(item.data_registro),
            "observacao": clean_text(item.observacao),
        }
        for item in recebimentos
    ]
    despesa_rows = [
        {
            "id": item.id,
            "descricao": clean_text(item.descricao),
            "valor": safe_float(item.valor, 0.0),
            "tipo_despesa": upper_text(item.tipo_despesa or "ROTA"),
            "categoria": upper_text(item.categoria),
            "data_registro": clean_text(item.data_registro),
            "observacao": clean_text(item.observacao),
        }
        for item in despesas
    ]
    kg_transferido_por_destino = await kg_transferido_convertido_por_destino(db, [programacao.codigo_programacao])
    kg_transferido_convertido = kg_transferido_por_destino.get(upper_text(programacao.codigo_programacao), 0.0)
    if total_kg <= 0 and kg_transferido_convertido > 0:
        total_kg = kg_transferido_convertido
    total_recebido = sum(item["valor"] for item in recebimento_rows) + total_controle
    total_despesas = sum(item["valor"] for item in despesa_rows)
    programacao_row = serialize_programacao(programacao, kg_transferido_convertido)
    if safe_int(programacao_row.get("caixas_carregadas"), 0) <= 0:
        programacao_row["caixas_carregadas"] = total_caixas_programadas
        programacao_row["nf_caixas"] = total_caixas_programadas
    if safe_float(programacao_row.get("kg_carregado"), 0.0) <= 0:
        programacao_row["kg_carregado"] = total_kg
        programacao_row["nf_kg_carregado"] = total_kg
    if (
        safe_float(programacao_row.get("nf_saldo"), 0.0) <= 0
        and safe_float(programacao_row.get("nf_kg"), 0.0) > 0
        and safe_float(programacao_row.get("kg_carregado"), 0.0) > 0
    ):
        programacao_row["nf_saldo"] = round(
            max(safe_float(programacao_row.get("nf_kg"), 0.0) - safe_float(programacao_row.get("kg_carregado"), 0.0), 0.0),
            2,
        )
    programacao_row["nf_saldo_valor"] = round(
        max(safe_float(programacao_row.get("nf_saldo"), 0.0), 0.0) * safe_float(programacao_row.get("nf_preco"), 0.0),
        2,
    )
    programacao_row["desconto_fornecedor"] = programacao_row["nf_saldo_valor"]
    return {
        "programacao": programacao_row,
        "resumo": {
            "clientes": len(itens),
            "caixas": total_caixas,
            "caixas_programadas": total_caixas_programadas,
            "caixas_entregues": total_caixas_entregues,
            "kg": total_kg,
            "recebido": total_recebido,
            "despesas": total_despesas,
            "adiantamento": safe_float(programacao.adiantamento or programacao.adiantamento_rota, 0.0),
            "saldo": total_recebido - total_despesas - safe_float(programacao.adiantamento or programacao.adiantamento_rota, 0.0),
            "entregues": total_entregues,
            "pendentes": total_pendentes,
            "alterados": total_alterados,
            "cancelados": total_cancelados,
            "com_localizacao": total_com_localizacao,
        },
        "itens": item_rows,
        "recebimentos": recebimento_rows,
        "despesas": despesa_rows,
    }
