import json
import logging
from datetime import datetime
from typing import Any

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import safe_float, safe_int

logger = logging.getLogger(__name__)


def upper_text(value: Any) -> str:
    return str(value or "").strip().upper()


def _json_payload(value: Any) -> str:
    try:
        return json.dumps(value or {}, ensure_ascii=False, sort_keys=True)
    except Exception:
        return json.dumps({"raw": str(value)}, ensure_ascii=False)


def _data_hora(value: Any = None) -> str:
    txt = str(value or "").strip()
    if txt:
        return txt
    return datetime.now().isoformat(timespec="seconds")


async def registrar_roteiro_operacional(
    db: AsyncSession,
    *,
    tipo_evento: str,
    codigo_programacao: str,
    origem: str = "",
    destino: str = "",
    motorista_codigo: str = "",
    motorista_nome: str = "",
    pedido: str = "",
    cod_cliente: str = "",
    cliente_nome: str = "",
    caixas: Any = 0,
    kg: Any = 0.0,
    media: Any = 0.0,
    aves_por_caixa: Any = 0,
    nf_numero: str = "",
    nf_preco: Any = 0.0,
    lotes: Any = "",
    data_hora: Any = None,
    observacao: str = "",
    payload: Any = None,
) -> None:
    try:
        lotes_value = _json_payload(lotes) if isinstance(lotes, (dict, list)) else str(lotes or "").strip()
        await db.execute(
            text(
                """
                INSERT INTO roteiro_operacional (
                    tipo_evento, codigo_programacao, origem, destino, motorista_codigo, motorista_nome,
                    pedido, cod_cliente, cliente_nome, caixas, kg, media, aves_por_caixa,
                    nf_numero, nf_preco, lotes, data_hora, observacao, payload_json, created_at
                ) VALUES (
                    :tipo_evento, :codigo_programacao, :origem, :destino, :motorista_codigo, :motorista_nome,
                    :pedido, :cod_cliente, :cliente_nome, :caixas, :kg, :media, :aves_por_caixa,
                    :nf_numero, :nf_preco, :lotes, :data_hora, :observacao, :payload_json, :created_at
                )
                """
            ),
            {
                "tipo_evento": upper_text(tipo_evento),
                "codigo_programacao": upper_text(codigo_programacao),
                "origem": str(origem or "").strip(),
                "destino": str(destino or "").strip(),
                "motorista_codigo": upper_text(motorista_codigo),
                "motorista_nome": upper_text(motorista_nome),
                "pedido": str(pedido or "").strip(),
                "cod_cliente": upper_text(cod_cliente),
                "cliente_nome": upper_text(cliente_nome),
                "caixas": safe_int(caixas, 0),
                "kg": safe_float(kg, 0.0),
                "media": safe_float(media, 0.0),
                "aves_por_caixa": safe_int(aves_por_caixa, 0),
                "nf_numero": str(nf_numero or "").strip(),
                "nf_preco": safe_float(nf_preco, 0.0),
                "lotes": lotes_value,
                "data_hora": _data_hora(data_hora),
                "observacao": str(observacao or "").strip(),
                "payload_json": _json_payload(payload),
                "created_at": datetime.now().isoformat(timespec="seconds"),
            },
        )
    except Exception:
        logger.debug("Falha ao registrar roteiro_operacional.", exc_info=True)


async def listar_roteiro_operacional(db: AsyncSession, codigo_programacao: str) -> list[dict[str, Any]]:
    codigo = upper_text(codigo_programacao)
    rows: list[dict[str, Any]] = []
    try:
        result = await db.execute(
            text(
                """
                SELECT *
                  FROM roteiro_operacional
                 WHERE UPPER(TRIM(COALESCE(codigo_programacao,''))) = UPPER(TRIM(:codigo))
                 ORDER BY COALESCE(NULLIF(TRIM(data_hora), ''), created_at, ''), id
                """
            ),
            {"codigo": codigo},
        )
        rows = [dict(row) for row in result.mappings().all()]
    except Exception:
        logger.debug("Falha ao listar roteiro_operacional.", exc_info=True)
    try:
        legacy = await db.execute(
            text(
                """
                SELECT
                    evento AS tipo_evento,
                    codigo_programacao,
                    'LEGADO' AS origem,
                    '' AS destino,
                    '' AS motorista_codigo,
                    '' AS motorista_nome,
                    pedido,
                    cod_cliente,
                    '' AS cliente_nome,
                    0 AS caixas,
                    0 AS kg,
                    0 AS media,
                    0 AS aves_por_caixa,
                    '' AS nf_numero,
                    0 AS nf_preco,
                    '' AS lotes,
                    COALESCE(NULLIF(TRIM(registrado_em), ''), created_at, '') AS data_hora,
                    '' AS observacao,
                    payload_json,
                    created_at,
                    id
                  FROM programacao_itens_log
                 WHERE UPPER(TRIM(COALESCE(codigo_programacao,''))) = UPPER(TRIM(:codigo))
                 ORDER BY COALESCE(NULLIF(TRIM(registrado_em), ''), created_at, ''), id
                """
            ),
            {"codigo": codigo},
        )
        existing = {
            (
                str(row.get("tipo_evento") or ""),
                str(row.get("pedido") or ""),
                str(row.get("cod_cliente") or ""),
                str(row.get("data_hora") or ""),
            )
            for row in rows
        }
        for row in legacy.mappings().all():
            item = dict(row)
            key = (
                str(item.get("tipo_evento") or ""),
                str(item.get("pedido") or ""),
                str(item.get("cod_cliente") or ""),
                str(item.get("data_hora") or ""),
            )
            if key not in existing:
                rows.append(item)
    except Exception:
        logger.debug("Falha ao listar logs legados do roteiro.", exc_info=True)
    rows.sort(key=lambda item: (str(item.get("data_hora") or item.get("created_at") or ""), int(item.get("id") or 0)))
    return rows
