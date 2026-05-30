# backend/api/v1/endpoints/compras.py
"""Compras, NF-e recebidas e base inicial do manifestador fiscal."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import uuid4
import xml.etree.ElementTree as ET

from fastapi import APIRouter, Depends, File, HTTPException, Query, Request, UploadFile, status
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.ext.asyncio import AsyncSession

from backend.api.v1.endpoints.programacao import upper_text
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.user import User
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()

NATUREZAS_OPERACAO = [
    "COMPRA/AQUISICAO",
    "COMPRAS USO E CONSUMO",
    "BONIFICACAO DE ENTRADA",
    "COMPRAS DE ATIVO IMOBILIZADO",
    "OUTRAS ENTRADAS",
    "OUTRAS ENTRADAS - REMESSA PALLET",
]

DEFAULT_LOGISTICA = {
    "produto_padrao": "CARGA",
    "unidade_padrao": "KG",
}


class ComprasNfeRow(BaseModel):
    id: int
    chave_acesso: str = ""
    serie: str = ""
    numero: str = ""
    fornecedor_documento: str = ""
    fornecedor_razao: str = ""
    emissao: str = ""
    valor_total: float = 0
    situacao_nfe: str = "Autorizado"
    nsu: str = ""
    natureza_operacao: str = ""
    fornecedor_id: int | None = None
    origem: str = "XML"
    xml_disponivel: bool = False
    estoque_fiscal_status: str = "PENDENTE"
    estoque_fisico_status: str = "PENDENTE"
    estoque_kg_entrada: float = 0
    estoque_kg_saldo: float = 0
    estoque_caixas_entrada: int = 0
    produto_id: int | None = None
    produto: str = ""


class ComprasNfeListResponse(BaseModel):
    naturezas: list[str]
    total: int
    total_valor: float
    rows: list[ComprasNfeRow]


class ComprasNfeImportResponse(BaseModel):
    ok: bool
    message: str
    nfe: ComprasNfeRow
    fornecedor_criado: bool = False


class EstoqueEntradaPayload(BaseModel):
    quantidade_kg: float = 0
    quantidade_caixas: int = 0
    produto: str = ""
    observacao: str = ""


class EstoqueAjustePayload(BaseModel):
    tipo_estoque: str = "FISICO"
    tipo_movimento: str = "ENTRADA"
    produto_id: int | None = None
    produto: str = ""
    quantidade_kg: float = 0
    quantidade_caixas: int = 0
    observacao: str = ""


class EstoqueResumoResponse(BaseModel):
    fiscal_entrada_kg: float = 0
    fiscal_saida_kg: float = 0
    fiscal_saldo_kg: float = 0
    fisico_entrada_kg: float = 0
    fisico_saida_kg: float = 0
    fisico_saldo_kg: float = 0
    notas_pendentes: int = 0
    saidas_pendentes_sefaz: int = 0
    movimentos: list[dict[str, Any]] = []
    produtos: list[dict[str, Any]] = []


class EstoqueSyncResponse(BaseModel):
    ok: bool
    message: str
    saidas_criadas: int = 0
    programacoes_vinculadas: int = 0


class EstoqueManutencaoResponse(BaseModel):
    ok: bool
    message: str
    movimentos_criados: int = 0


def money(value: Any) -> float:
    try:
        return round(float(value or 0), 2)
    except Exception:
        return 0.0


def clean_digits(value: Any) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


async def logistica_defaults(db: AsyncSession) -> dict[str, str]:
    try:
        result = await db.execute(
            text(
                """
                SELECT produto_padrao, unidade_padrao
                  FROM empresa_configuracao_logistica
                 ORDER BY company_id ASC
                 LIMIT 1
                """
            )
        )
        row = result.mappings().first()
        if row:
            return {
                "produto_padrao": upper_text(row.get("produto_padrao") or DEFAULT_LOGISTICA["produto_padrao"]),
                "unidade_padrao": upper_text(row.get("unidade_padrao") or DEFAULT_LOGISTICA["unidade_padrao"]),
            }
    except SQLAlchemyError:
        pass
    return dict(DEFAULT_LOGISTICA)


def produto_codigo_base(value: Any) -> str:
    text_value = upper_text(value)
    code = "".join(ch if ch.isalnum() else "-" for ch in text_value).strip("-")
    while "--" in code:
        code = code.replace("--", "-")
    return (code or "PRODUTO")[:40]


def xml_text(root: ET.Element, local_name: str) -> str:
    for node in root.iter():
        if str(node.tag).split("}")[-1] == local_name:
            return str(node.text or "").strip()
    return ""


def find_first(root: ET.Element, local_name: str) -> ET.Element | None:
    for node in root.iter():
        if str(node.tag).split("}")[-1] == local_name:
            return node
    return None


def child_text(node: ET.Element | None, local_name: str) -> str:
    if node is None:
        return ""
    for child in node.iter():
        if str(child.tag).split("}")[-1] == local_name:
            return str(child.text or "").strip()
    return ""


def parse_nfe_xml(contents: bytes) -> dict[str, Any]:
    try:
        root = ET.fromstring(contents)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"XML invalido: {exc}") from exc

    inf_nfe = find_first(root, "infNFe")
    chave = ""
    if inf_nfe is not None:
        chave = str(inf_nfe.attrib.get("Id") or "").replace("NFe", "").strip()
    emit = find_first(root, "emit")
    total = find_first(root, "ICMSTot")
    documento = clean_digits(child_text(emit, "CNPJ") or child_text(emit, "CPF"))
    razao = upper_text(child_text(emit, "xNome"))
    produtos_by_key: dict[str, dict[str, Any]] = {}
    for det in [node for node in root.iter() if str(node.tag).split("}")[-1] == "det"]:
        prod = find_first(det, "prod")
        produto_nome = upper_text(child_text(prod, "xProd")) or "CARGA"
        codigo_produto = upper_text(child_text(prod, "cProd"))
        unidade = upper_text(child_text(prod, "uCom"))
        quantidade = money(child_text(prod, "qCom"))
        valor_total_item = money(child_text(prod, "vProd"))
        key = produto_nome
        item = produtos_by_key.setdefault(
            key,
            {
                "codigo": codigo_produto,
                "produto": produto_nome,
                "unidade": unidade or "KG",
                "quantidade_kg": 0.0,
                "quantidade_caixas": 0,
                "valor_total": 0.0,
            },
        )
        if unidade in {"KG", "KILO", "QUILO"}:
            item["quantidade_kg"] = money(item["quantidade_kg"] + quantidade)
        elif unidade in {"CX", "CAIXA", "CAIXAS"}:
            item["quantidade_caixas"] = int(item["quantidade_caixas"] or 0) + int(round(quantidade))
        item["valor_total"] = money(item["valor_total"] + valor_total_item)
    produtos = list(produtos_by_key.values())
    estoque_kg = money(sum(item["quantidade_kg"] for item in produtos))
    produto_nome_resumo = produtos[0]["produto"] if len(produtos) == 1 else f"MISTO ({len(produtos)} PRODUTOS)"
    return {
        "chave_acesso": chave or clean_digits(xml_text(root, "chNFe")),
        "serie": child_text(find_first(root, "ide"), "serie"),
        "numero": child_text(find_first(root, "ide"), "nNF"),
        "fornecedor_documento": documento,
        "fornecedor_razao": razao,
        "emissao": child_text(find_first(root, "ide"), "dhEmi") or child_text(find_first(root, "ide"), "dEmi"),
        "valor_total": money(child_text(total, "vNF")),
        "situacao_nfe": "Autorizado",
        "produto": produto_nome_resumo or "CARGA",
        "produtos": produtos or [{"codigo": "", "produto": "CARGA", "unidade": "KG", "quantidade_kg": 0, "quantidade_caixas": 0, "valor_total": 0}],
        "estoque_kg_entrada": money(estoque_kg),
    }


async def ensure_compras_schema(db: AsyncSession) -> None:
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                nome TEXT NOT NULL,
                descricao TEXT,
                categoria TEXT DEFAULT 'AVES',
                unidade TEXT DEFAULT 'KG',
                unidade_estoque TEXT DEFAULT 'KG',
                controla_estoque_fisico INTEGER DEFAULT 1,
                controla_estoque_fiscal INTEGER DEFAULT 1,
                estoque_min_kg REAL DEFAULT 0,
                estoque_min_caixas INTEGER DEFAULT 0,
                ncm TEXT,
                cest TEXT,
                cfop_entrada TEXT,
                cfop_saida TEXT,
                ean TEXT,
                custo_padrao REAL DEFAULT 0,
                preco_padrao REAL DEFAULT 0,
                status TEXT DEFAULT 'ATIVO'
            )
            """
        )
    )
    await db.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_compras_produtos_codigo ON produtos(codigo)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_produtos_nome ON produtos(nome)"))
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS compras_nfe (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chave_acesso TEXT UNIQUE,
                serie TEXT,
                numero TEXT,
                fornecedor_documento TEXT,
                fornecedor_razao TEXT,
                emissao TEXT,
                valor_total REAL DEFAULT 0,
                situacao_nfe TEXT DEFAULT 'Autorizado',
                nsu TEXT,
                natureza_operacao TEXT,
                fornecedor_id INTEGER,
                origem TEXT DEFAULT 'XML',
                xml_path TEXT,
                pdf_path TEXT,
                estoque_fiscal_status TEXT DEFAULT 'PENDENTE',
                estoque_fisico_status TEXT DEFAULT 'PENDENTE',
                estoque_kg_entrada REAL DEFAULT 0,
                estoque_kg_saldo REAL DEFAULT 0,
                estoque_caixas_entrada INTEGER DEFAULT 0,
                produto_id INTEGER,
                produto TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT
            )
            """
        )
    )
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_chave ON compras_nfe(chave_acesso)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_fornecedor ON compras_nfe(fornecedor_documento)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_emissao ON compras_nfe(emissao)"))
    existing_columns = {str(row[1]) for row in (await db.execute(text("PRAGMA table_info(compras_nfe)"))).all()}
    for column, definition in {
        "estoque_fiscal_status": "TEXT DEFAULT 'PENDENTE'",
        "estoque_fisico_status": "TEXT DEFAULT 'PENDENTE'",
        "estoque_kg_entrada": "REAL DEFAULT 0",
        "estoque_kg_saldo": "REAL DEFAULT 0",
        "estoque_caixas_entrada": "INTEGER DEFAULT 0",
        "produto_id": "INTEGER",
        "produto": "TEXT",
        "codigo_programacao": "TEXT",
        "vinculada_em": "TEXT",
        "vinculada_por": "TEXT",
    }.items():
        if column not in existing_columns:
            await db.execute(text(f"ALTER TABLE compras_nfe ADD COLUMN {column} {definition}"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_numero ON compras_nfe(numero)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_vinculo ON compras_nfe(codigo_programacao)"))
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS compras_nfe_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                compra_id INTEGER NOT NULL,
                produto_id INTEGER,
                produto TEXT NOT NULL,
                codigo_produto TEXT,
                unidade TEXT DEFAULT 'KG',
                quantidade_kg REAL DEFAULT 0,
                quantidade_caixas INTEGER DEFAULT 0,
                valor_total REAL DEFAULT 0,
                company_id INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    item_columns = {str(row[1]) for row in (await db.execute(text("PRAGMA table_info(compras_nfe_itens)"))).all()}
    for column, definition in {
        "compra_id": "INTEGER",
        "produto_id": "INTEGER",
        "produto": "TEXT",
        "codigo_produto": "TEXT",
        "unidade": "TEXT DEFAULT 'KG'",
        "quantidade_kg": "REAL DEFAULT 0",
        "quantidade_caixas": "INTEGER DEFAULT 0",
        "valor_total": "REAL DEFAULT 0",
        "company_id": "INTEGER",
    }.items():
        if column not in item_columns:
            await db.execute(text(f"ALTER TABLE compras_nfe_itens ADD COLUMN {column} {definition}"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_itens_compra ON compras_nfe_itens(compra_id)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_compras_nfe_itens_produto ON compras_nfe_itens(produto_id, produto)"))
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS estoque_movimentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo_estoque TEXT NOT NULL,
                tipo_movimento TEXT NOT NULL,
                origem TEXT NOT NULL,
                origem_id INTEGER,
                chave_acesso TEXT,
                numero_nf TEXT,
                codigo_programacao TEXT,
                produto_id INTEGER,
                produto TEXT,
                quantidade_kg REAL DEFAULT 0,
                quantidade_caixas INTEGER DEFAULT 0,
                valor_unitario REAL DEFAULT 0,
                valor_total REAL DEFAULT 0,
                natureza_operacao TEXT,
                status_fiscal TEXT DEFAULT 'NAO_APLICAVEL',
                observacao TEXT,
                company_id INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    movimento_columns = {str(row[1]) for row in (await db.execute(text("PRAGMA table_info(estoque_movimentos)"))).all()}
    for column, definition in {
        "produto_id": "INTEGER",
        "produto": "TEXT",
        "company_id": "INTEGER",
    }.items():
        if column not in movimento_columns:
            await db.execute(text(f"ALTER TABLE estoque_movimentos ADD COLUMN {column} {definition}"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_estoque_movimentos_tipo ON estoque_movimentos(tipo_estoque, tipo_movimento)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_estoque_movimentos_produto ON estoque_movimentos(produto_id, produto)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_estoque_movimentos_nf ON estoque_movimentos(numero_nf)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_estoque_movimentos_prog ON estoque_movimentos(codigo_programacao)"))


async def ensure_produto_catalogo(db: AsyncSession, nome: str, *, codigo: str = "", unidade: str = "") -> int | None:
    defaults = await logistica_defaults(db)
    produto_padrao = defaults["produto_padrao"]
    unidade_padrao = upper_text(unidade or defaults["unidade_padrao"]) or "KG"
    nome_norm = upper_text(nome or produto_padrao) or produto_padrao
    codigo_explicit = bool(upper_text(codigo))
    codigo_norm = produto_codigo_base(codigo or nome_norm)
    if codigo_explicit and codigo_norm:
        result = await db.execute(text("SELECT id FROM produtos WHERE UPPER(COALESCE(codigo,''))=:codigo LIMIT 1"), {"codigo": codigo_norm})
        produto_id = result.scalar_one_or_none()
        if produto_id:
            return int(produto_id)
    result = await db.execute(text("SELECT id FROM produtos WHERE UPPER(COALESCE(nome,''))=:nome LIMIT 1"), {"nome": nome_norm})
    produto_id = result.scalar_one_or_none()
    if produto_id:
        return int(produto_id)
    codigo_final = codigo_norm
    for suffix in range(0, 1000):
        candidate = codigo_final if suffix == 0 else f"{codigo_norm[:34]}-{suffix:03d}"
        result = await db.execute(text("SELECT id FROM produtos WHERE UPPER(COALESCE(codigo,''))=:codigo LIMIT 1"), {"codigo": candidate})
        if result.scalar_one_or_none():
            continue
        await db.execute(
            text(
                """
                INSERT INTO produtos (
                    codigo, nome, descricao, categoria, unidade, unidade_estoque,
                    controla_estoque_fisico, controla_estoque_fiscal, status
                ) VALUES (
                    :codigo, :nome, :descricao, :categoria, :unidade, :unidade, 1, 1, 'ATIVO'
                )
                """
            ),
            {
                "codigo": candidate,
                "nome": nome_norm,
                "categoria": "AVES" if "AVE" in nome_norm or "FRANGO" in nome_norm else "GERAL",
                "unidade": unidade_padrao,
                "descricao": "Cadastro automatico criado a partir de compras/estoque.",
            },
        )
        break
    else:
        raise HTTPException(status_code=409, detail="Nao foi possivel gerar codigo unico para o produto.")
    result = await db.execute(text("SELECT id FROM produtos WHERE UPPER(COALESCE(nome,''))=:nome LIMIT 1"), {"nome": nome_norm})
    produto_id = result.scalar_one_or_none()
    return int(produto_id) if produto_id else None


async def ensure_fornecedor_from_xml(db: AsyncSession, nfe: dict[str, Any], natureza: str) -> tuple[int | None, bool]:
    documento = clean_digits(nfe.get("fornecedor_documento"))
    if not documento:
        return None, False
    result = await db.execute(text("SELECT id FROM fornecedores WHERE documento=:documento LIMIT 1"), {"documento": documento})
    fornecedor_id = result.scalar_one_or_none()
    if fornecedor_id:
        return int(fornecedor_id), False
    perfil = "DISTRIBUICAO_GERAL" if natureza == "COMPRA/AQUISICAO" else "OUTROS"
    await db.execute(
        text(
            """
            INSERT INTO fornecedores (
                razao_social, nome_fantasia, documento, tipo_pessoa, perfil_fornecedor,
                status, observacao
            ) VALUES (
                :razao_social, :nome_fantasia, :documento, :tipo_pessoa, :perfil_fornecedor,
                'ATIVO', 'Cadastro automatico criado pela importacao de XML'
            )
            """
        ),
        {
            "razao_social": upper_text(nfe.get("fornecedor_razao")) or documento,
            "nome_fantasia": upper_text(nfe.get("fornecedor_razao")),
            "documento": documento,
            "tipo_pessoa": "CPF" if len(documento) == 11 else "CNPJ",
            "perfil_fornecedor": perfil,
        },
    )
    result = await db.execute(text("SELECT id FROM fornecedores WHERE documento=:documento LIMIT 1"), {"documento": documento})
    return int(result.scalar_one_or_none() or 0) or None, True


def row_from_mapping(item: dict[str, Any]) -> ComprasNfeRow:
    return ComprasNfeRow(
        id=int(item.get("id") or 0),
        chave_acesso=str(item.get("chave_acesso") or ""),
        serie=str(item.get("serie") or ""),
        numero=str(item.get("numero") or ""),
        fornecedor_documento=str(item.get("fornecedor_documento") or ""),
        fornecedor_razao=str(item.get("fornecedor_razao") or ""),
        emissao=str(item.get("emissao") or ""),
        valor_total=money(item.get("valor_total")),
        situacao_nfe=str(item.get("situacao_nfe") or "Autorizado"),
        nsu=str(item.get("nsu") or ""),
        natureza_operacao=str(item.get("natureza_operacao") or ""),
        fornecedor_id=int(item["fornecedor_id"]) if item.get("fornecedor_id") else None,
        origem=str(item.get("origem") or "XML"),
        xml_disponivel=bool(str(item.get("xml_path") or "").strip()),
        estoque_fiscal_status=str(item.get("estoque_fiscal_status") or "PENDENTE"),
        estoque_fisico_status=str(item.get("estoque_fisico_status") or "PENDENTE"),
        estoque_kg_entrada=money(item.get("estoque_kg_entrada")),
        estoque_kg_saldo=money(item.get("estoque_kg_saldo")),
        estoque_caixas_entrada=int(item.get("estoque_caixas_entrada") or 0),
        produto_id=int(item["produto_id"]) if item.get("produto_id") else None,
        produto=str(item.get("produto") or ""),
    )


@router.get("/nfe", response_model=ComprasNfeListResponse)
async def listar_nfe_compras(
    natureza: str = Query(default="TODAS"),
    limit: int = Query(default=500, ge=1, le=5000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    await ensure_compras_schema(db)
    natureza_norm = upper_text(natureza or "TODAS")
    sql = "SELECT * FROM compras_nfe"
    params: dict[str, Any] = {"limit": limit}
    if natureza_norm and natureza_norm != "TODAS":
        sql += " WHERE natureza_operacao=:natureza"
        params["natureza"] = natureza_norm
    sql += " ORDER BY COALESCE(emissao, created_at) DESC, id DESC LIMIT :limit"
    result = await db.execute(text(sql), params)
    rows = [row_from_mapping(dict(row)) for row in result.mappings().all()]
    return ComprasNfeListResponse(
        naturezas=NATUREZAS_OPERACAO,
        total=len(rows),
        total_valor=money(sum(item.valor_total for item in rows)),
        rows=rows,
    )


@router.post("/nfe/importar-xml", response_model=ComprasNfeImportResponse, status_code=status.HTTP_201_CREATED)
async def importar_nfe_xml(
    request: Request,
    natureza_operacao: str = Query(default="COMPRA/AQUISICAO"),
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_compras_schema(db)
    natureza = upper_text(natureza_operacao or "COMPRA/AQUISICAO")
    if natureza not in NATUREZAS_OPERACAO:
        natureza = "COMPRA/AQUISICAO"
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=422, detail="Arquivo XML vazio.")
    nfe = parse_nfe_xml(contents)
    if not nfe["chave_acesso"]:
        raise HTTPException(status_code=422, detail="Nao foi possivel identificar a chave de acesso no XML.")

    fornecedor_id, fornecedor_criado = await ensure_fornecedor_from_xml(db, nfe, natureza)
    nfe_produtos = [item for item in nfe.get("produtos", []) if isinstance(item, dict)]
    produto_ids: list[int] = []
    for item in nfe_produtos:
        produto_id_item = await ensure_produto_catalogo(db, item.get("produto") or "", codigo=item.get("codigo") or "", unidade=item.get("unidade") or "")
        if produto_id_item:
            item["produto_id"] = produto_id_item
            produto_ids.append(produto_id_item)
    produto_id = produto_ids[0] if len(set(produto_ids)) == 1 else None
    root = Path(".rotahub_runtime") / "compras_xml"
    root.mkdir(parents=True, exist_ok=True)
    xml_path = root / f"{nfe['chave_acesso']}_{uuid4().hex[:8]}.xml"
    xml_path.write_bytes(contents)
    await db.execute(
        text(
            """
            INSERT INTO compras_nfe (
                chave_acesso, serie, numero, fornecedor_documento, fornecedor_razao, emissao,
                valor_total, situacao_nfe, nsu, natureza_operacao, fornecedor_id, origem,
                xml_path, estoque_kg_entrada, estoque_kg_saldo, produto_id, produto, updated_at
            ) VALUES (
                :chave_acesso, :serie, :numero, :fornecedor_documento, :fornecedor_razao, :emissao,
                :valor_total, :situacao_nfe, :nsu, :natureza_operacao, :fornecedor_id, 'XML',
                :xml_path, :estoque_kg_entrada, :estoque_kg_entrada, :produto_id, :produto, CURRENT_TIMESTAMP
            )
            ON CONFLICT(chave_acesso) DO UPDATE SET
                serie=excluded.serie,
                numero=excluded.numero,
                fornecedor_documento=excluded.fornecedor_documento,
                fornecedor_razao=excluded.fornecedor_razao,
                emissao=excluded.emissao,
                valor_total=excluded.valor_total,
                situacao_nfe=excluded.situacao_nfe,
                natureza_operacao=excluded.natureza_operacao,
                fornecedor_id=excluded.fornecedor_id,
                xml_path=excluded.xml_path,
                estoque_kg_entrada=excluded.estoque_kg_entrada,
                estoque_kg_saldo=CASE
                    WHEN COALESCE(compras_nfe.estoque_kg_saldo, 0) <= 0 THEN excluded.estoque_kg_entrada
                    ELSE compras_nfe.estoque_kg_saldo
                END,
                produto_id=excluded.produto_id,
                produto=excluded.produto,
                updated_at=CURRENT_TIMESTAMP
            """
        ),
        {
            **nfe,
            "nsu": "",
            "natureza_operacao": natureza,
            "fornecedor_id": fornecedor_id,
            "xml_path": str(xml_path),
            "produto_id": produto_id,
        },
    )
    compra_result = await db.execute(text("SELECT id FROM compras_nfe WHERE chave_acesso=:chave LIMIT 1"), {"chave": nfe["chave_acesso"]})
    compra_id = int(compra_result.scalar_one_or_none() or 0)
    if compra_id:
        await db.execute(text("DELETE FROM compras_nfe_itens WHERE compra_id=:compra_id"), {"compra_id": compra_id})
        for item in nfe_produtos:
            await db.execute(
                text(
                    """
                    INSERT INTO compras_nfe_itens (
                        compra_id, produto_id, produto, codigo_produto, unidade,
                        quantidade_kg, quantidade_caixas, valor_total
                    ) VALUES (
                        :compra_id, :produto_id, :produto, :codigo_produto, :unidade,
                        :quantidade_kg, :quantidade_caixas, :valor_total
                    )
                    """
                ),
                {
                    "compra_id": compra_id,
                    "produto_id": item.get("produto_id"),
                    "produto": upper_text(item.get("produto") or "CARGA"),
                    "codigo_produto": upper_text(item.get("codigo") or ""),
                    "unidade": upper_text(item.get("unidade") or "KG"),
                    "quantidade_kg": money(item.get("quantidade_kg")),
                    "quantidade_caixas": int(item.get("quantidade_caixas") or 0),
                    "valor_total": money(item.get("valor_total")),
                },
            )
    record_audit_log(
        db,
        action="compras_nfe_xml_importado",
        actor_user=current_user,
        entity_type="compras_nfe",
        entity_id=compra_id or nfe["chave_acesso"],
        ip_address=client_ip_from_request(request),
        metadata={
            "chave_acesso": nfe["chave_acesso"],
            "numero": nfe.get("numero"),
            "natureza_operacao": natureza,
            "fornecedor_criado": fornecedor_criado,
            "produtos": [
                {
                    "produto_id": item.get("produto_id"),
                    "produto": upper_text(item.get("produto")),
                    "quantidade_kg": money(item.get("quantidade_kg")),
                    "quantidade_caixas": int(item.get("quantidade_caixas") or 0),
                }
                for item in nfe_produtos
            ],
        },
    )
    await db.commit()
    result = await db.execute(text("SELECT * FROM compras_nfe WHERE chave_acesso=:chave LIMIT 1"), {"chave": nfe["chave_acesso"]})
    row = row_from_mapping(dict(result.mappings().first() or {}))
    return ComprasNfeImportResponse(
        ok=True,
        message="XML importado e fornecedor verificado.",
        nfe=row,
        fornecedor_criado=fornecedor_criado,
    )


@router.get("/estoque/resumo", response_model=EstoqueResumoResponse)
async def resumo_estoque_compras(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    await ensure_compras_schema(db)
    result = await db.execute(
        text(
            """
            SELECT
                tipo_estoque,
                tipo_movimento,
                COALESCE(SUM(quantidade_kg), 0) AS kg
              FROM estoque_movimentos
             GROUP BY tipo_estoque, tipo_movimento
            """
        )
    )
    totals: dict[tuple[str, str], float] = {}
    for row in result.mappings().all():
        totals[(str(row["tipo_estoque"]), str(row["tipo_movimento"]))] = money(row["kg"])
    fiscal_entrada = totals.get(("FISCAL", "ENTRADA"), 0.0)
    fiscal_saida = totals.get(("FISCAL", "SAIDA"), 0.0)
    fisico_entrada = totals.get(("FISICO", "ENTRADA"), 0.0)
    fisico_saida = totals.get(("FISICO", "SAIDA"), 0.0)
    pendentes = await db.execute(
        text("SELECT COUNT(*) FROM compras_nfe WHERE COALESCE(estoque_fiscal_status, 'PENDENTE')='PENDENTE'")
    )
    saidas_pendentes = await db.execute(
        text("SELECT COUNT(*) FROM estoque_movimentos WHERE tipo_estoque='FISCAL' AND tipo_movimento='SAIDA' AND status_fiscal='PENDENTE_SEFAZ'")
    )
    movements = await db.execute(
        text(
            """
            SELECT tipo_estoque, tipo_movimento, numero_nf, codigo_programacao, produto,
                   quantidade_kg, quantidade_caixas, valor_total, status_fiscal, created_at
              FROM estoque_movimentos
             ORDER BY id DESC
             LIMIT 12
            """
        )
    )
    produtos = await db.execute(
        text(
            """
            SELECT
                COALESCE(p.id, 0) AS produto_id,
                COALESCE(p.codigo, '') AS codigo,
                COALESCE(p.nome, em.produto, 'SEM PRODUTO') AS nome,
                SUM(CASE WHEN em.tipo_estoque='FISCAL' AND em.tipo_movimento='ENTRADA' THEN COALESCE(em.quantidade_kg, 0) ELSE 0 END) AS fiscal_entrada_kg,
                SUM(CASE WHEN em.tipo_estoque='FISCAL' AND em.tipo_movimento='SAIDA' THEN COALESCE(em.quantidade_kg, 0) ELSE 0 END) AS fiscal_saida_kg,
                SUM(CASE WHEN em.tipo_estoque='FISICO' AND em.tipo_movimento='ENTRADA' THEN COALESCE(em.quantidade_kg, 0) ELSE 0 END) AS fisico_entrada_kg,
                SUM(CASE WHEN em.tipo_estoque='FISICO' AND em.tipo_movimento='SAIDA' THEN COALESCE(em.quantidade_kg, 0) ELSE 0 END) AS fisico_saida_kg
              FROM estoque_movimentos em
              LEFT JOIN produtos p ON p.id=em.produto_id
             GROUP BY COALESCE(p.id, 0), COALESCE(p.codigo, ''), COALESCE(p.nome, em.produto, 'SEM PRODUTO')
             ORDER BY nome
            """
        )
    )
    return EstoqueResumoResponse(
        fiscal_entrada_kg=fiscal_entrada,
        fiscal_saida_kg=fiscal_saida,
        fiscal_saldo_kg=money(fiscal_entrada - fiscal_saida),
        fisico_entrada_kg=fisico_entrada,
        fisico_saida_kg=fisico_saida,
        fisico_saldo_kg=money(fisico_entrada - fisico_saida),
        notas_pendentes=int(pendentes.scalar_one_or_none() or 0),
        saidas_pendentes_sefaz=int(saidas_pendentes.scalar_one_or_none() or 0),
        movimentos=[dict(row) for row in movements.mappings().all()],
        produtos=[
            {
                **dict(row),
                "fiscal_saldo_kg": money(row["fiscal_entrada_kg"] - row["fiscal_saida_kg"]),
                "fisico_saldo_kg": money(row["fisico_entrada_kg"] - row["fisico_saida_kg"]),
            }
            for row in produtos.mappings().all()
        ],
    )


@router.post("/estoque/ajuste", response_model=EstoqueManutencaoResponse)
async def ajustar_estoque_manual(
    payload: EstoqueAjustePayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_compras_schema(db)
    tipo_estoque = upper_text(payload.tipo_estoque or "FISICO")
    tipo_movimento = upper_text(payload.tipo_movimento or "ENTRADA")
    if tipo_estoque not in {"FISICO", "FISCAL"}:
        raise HTTPException(status_code=422, detail="Tipo de estoque invalido.")
    if tipo_movimento not in {"ENTRADA", "SAIDA"}:
        raise HTTPException(status_code=422, detail="Tipo de movimento invalido.")
    quantidade_kg = money(payload.quantidade_kg)
    quantidade_caixas = max(int(payload.quantidade_caixas or 0), 0)
    if quantidade_kg <= 0 and quantidade_caixas <= 0:
        raise HTTPException(status_code=422, detail="Informe KG ou caixas para ajustar o estoque.")
    defaults = await logistica_defaults(db)
    produto = upper_text(payload.produto or defaults["produto_padrao"])
    produto_id = int(payload.produto_id or 0) or await ensure_produto_catalogo(db, produto)
    await db.execute(
        text(
            """
            INSERT INTO estoque_movimentos (
                tipo_estoque, tipo_movimento, origem, produto_id, produto,
                quantidade_kg, quantidade_caixas, natureza_operacao, status_fiscal, observacao
            ) VALUES (
                :tipo_estoque, :tipo_movimento, 'AJUSTE_MANUAL', :produto_id, :produto,
                :quantidade_kg, :quantidade_caixas, 'AJUSTE MANUAL DE ESTOQUE',
                :status_fiscal, :observacao
            )
            """
        ),
        {
            "tipo_estoque": tipo_estoque,
            "tipo_movimento": tipo_movimento,
            "produto_id": produto_id,
            "produto": produto,
            "quantidade_kg": quantidade_kg,
            "quantidade_caixas": quantidade_caixas,
            "status_fiscal": "PENDENTE_SEFAZ" if tipo_estoque == "FISCAL" else "NAO_APLICAVEL",
            "observacao": upper_text(payload.observacao or "AJUSTE MANUAL"),
        },
    )
    record_audit_log(
        db,
        action="estoque_ajuste_manual",
        actor_user=current_user,
        entity_type="estoque_movimentos",
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={
            "tipo_estoque": tipo_estoque,
            "tipo_movimento": tipo_movimento,
            "produto_id": produto_id,
            "produto": produto,
            "quantidade_kg": quantidade_kg,
            "quantidade_caixas": quantidade_caixas,
            "observacao": upper_text(payload.observacao),
        },
    )
    await db.commit()
    return EstoqueManutencaoResponse(ok=True, message="Ajuste manual registrado no estoque.", movimentos_criados=1)


@router.post("/estoque/zerar-fisico", response_model=EstoqueManutencaoResponse)
async def zerar_estoque_fisico(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_compras_schema(db)
    saldos = await db.execute(
        text(
            """
            SELECT
                COALESCE(produto_id, 0) AS produto_id,
                COALESCE(produto, 'SEM PRODUTO') AS produto,
                SUM(CASE WHEN tipo_movimento='ENTRADA' THEN COALESCE(quantidade_kg, 0) ELSE -COALESCE(quantidade_kg, 0) END) AS saldo_kg,
                SUM(CASE WHEN tipo_movimento='ENTRADA' THEN COALESCE(quantidade_caixas, 0) ELSE -COALESCE(quantidade_caixas, 0) END) AS saldo_caixas
              FROM estoque_movimentos
             WHERE tipo_estoque='FISICO'
             GROUP BY COALESCE(produto_id, 0), COALESCE(produto, 'SEM PRODUTO')
            HAVING ABS(saldo_kg) > 0.0001 OR ABS(saldo_caixas) > 0
            """
        )
    )
    movimentos = 0
    audit_items: list[dict[str, Any]] = []
    for row in saldos.mappings().all():
        saldo_kg = money(row["saldo_kg"])
        saldo_caixas = int(row["saldo_caixas"] or 0)
        if saldo_kg == 0 and saldo_caixas == 0:
            continue
        tipo_movimento = "SAIDA" if saldo_kg > 0 or saldo_caixas > 0 else "ENTRADA"
        await db.execute(
            text(
                """
                INSERT INTO estoque_movimentos (
                    tipo_estoque, tipo_movimento, origem, produto_id, produto,
                    quantidade_kg, quantidade_caixas, natureza_operacao, status_fiscal, observacao
                ) VALUES (
                    'FISICO', :tipo_movimento, 'ZERAGEM_MANUAL', :produto_id, :produto,
                    :quantidade_kg, :quantidade_caixas, 'ZERAGEM MANUAL DE ESTOQUE FISICO',
                    'NAO_APLICAVEL', 'ZERAGEM MANUAL DO ESTOQUE FISICO'
                )
                """
            ),
            {
                "tipo_movimento": tipo_movimento,
                "produto_id": int(row["produto_id"] or 0) or None,
                "produto": upper_text(row["produto"]),
                "quantidade_kg": abs(saldo_kg),
                "quantidade_caixas": abs(saldo_caixas),
            },
        )
        movimentos += 1
        audit_items.append(
            {
                "produto_id": int(row["produto_id"] or 0) or None,
                "produto": upper_text(row["produto"]),
                "saldo_kg_zerado": saldo_kg,
                "saldo_caixas_zerado": saldo_caixas,
                "movimento_compensatorio": tipo_movimento,
            }
        )
    record_audit_log(
        db,
        action="estoque_fisico_zerado_manual",
        actor_user=current_user,
        entity_type="estoque_movimentos",
        severity="critical",
        ip_address=client_ip_from_request(request),
        metadata={"movimentos_criados": movimentos, "produtos": audit_items[:100]},
    )
    await db.commit()
    return EstoqueManutencaoResponse(
        ok=True,
        message="Estoque fisico zerado por movimentos compensatorios.",
        movimentos_criados=movimentos,
    )


@router.post("/nfe/{nfe_id}/confirmar-entrada", response_model=ComprasNfeRow)
async def confirmar_entrada_estoque(
    nfe_id: int,
    payload: EstoqueEntradaPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_compras_schema(db)
    result = await db.execute(text("SELECT * FROM compras_nfe WHERE id=:id LIMIT 1"), {"id": nfe_id})
    nfe = result.mappings().first()
    if not nfe:
        raise HTTPException(status_code=404, detail="NF-e nao encontrada.")
    itens_result = await db.execute(
        text(
            """
            SELECT produto_id, produto, quantidade_kg, quantidade_caixas, valor_total
              FROM compras_nfe_itens
             WHERE compra_id=:compra_id
             ORDER BY id
            """
        ),
        {"compra_id": nfe_id},
    )
    itens = [dict(row) for row in itens_result.mappings().all()]
    if not itens:
        defaults = await logistica_defaults(db)
        produto = upper_text(payload.produto or nfe.get("produto") or defaults["produto_padrao"])
        itens = [
            {
                "produto_id": await ensure_produto_catalogo(db, produto),
                "produto": produto,
                "quantidade_kg": money(payload.quantidade_kg or nfe.get("estoque_kg_entrada") or 0),
                "quantidade_caixas": int(payload.quantidade_caixas or nfe.get("estoque_caixas_entrada") or 0),
                "valor_total": money(nfe.get("valor_total")),
            }
        ]
    total_itens_kg = money(sum(money(item.get("quantidade_kg")) for item in itens))
    total_itens_caixas = sum(int(item.get("quantidade_caixas") or 0) for item in itens)
    override_kg = money(payload.quantidade_kg)
    override_caixas = int(payload.quantidade_caixas or 0)

    for item in itens:
        defaults = await logistica_defaults(db)
        produto = upper_text(item.get("produto") or defaults["produto_padrao"])
        produto_id = int(item.get("produto_id") or 0) or await ensure_produto_catalogo(db, produto)
        item_kg_original = money(item.get("quantidade_kg"))
        item_caixas_original = int(item.get("quantidade_caixas") or 0)
        kg = money((item_kg_original / total_itens_kg) * override_kg) if override_kg > 0 and total_itens_kg > 0 else item_kg_original
        caixas = int(round((item_caixas_original / total_itens_caixas) * override_caixas)) if override_caixas > 0 and total_itens_caixas > 0 else item_caixas_original
        valor_total_item = money(item.get("valor_total"))
        valor_unitario = money(valor_total_item / kg) if kg > 0 else 0
        for tipo_estoque, status_fiscal in [("FISCAL", "AUTORIZADO_ENTRADA"), ("FISICO", "NAO_APLICAVEL")]:
            exists = await db.execute(
                text(
                    """
                    SELECT id FROM estoque_movimentos
                     WHERE tipo_estoque=:tipo_estoque AND tipo_movimento='ENTRADA'
                       AND origem='COMPRAS_NFE' AND origem_id=:origem_id
                       AND COALESCE(produto_id, 0)=COALESCE(:produto_id, 0)
                       AND UPPER(COALESCE(produto, ''))=UPPER(:produto)
                     LIMIT 1
                    """
                ),
                {"tipo_estoque": tipo_estoque, "origem_id": nfe_id, "produto_id": produto_id or 0, "produto": produto},
            )
            if exists.scalar_one_or_none():
                continue
            await db.execute(
                text(
                    """
                    INSERT INTO estoque_movimentos (
                        tipo_estoque, tipo_movimento, origem, origem_id, chave_acesso, numero_nf,
                        produto_id, produto, quantidade_kg, quantidade_caixas, valor_unitario, valor_total,
                        natureza_operacao, status_fiscal, observacao
                    ) VALUES (
                        :tipo_estoque, 'ENTRADA', 'COMPRAS_NFE', :origem_id, :chave_acesso, :numero_nf,
                        :produto_id, :produto, :quantidade_kg, :quantidade_caixas, :valor_unitario, :valor_total,
                        :natureza_operacao, :status_fiscal, :observacao
                    )
                    """
                ),
                {
                    "tipo_estoque": tipo_estoque,
                    "origem_id": nfe_id,
                    "chave_acesso": nfe.get("chave_acesso"),
                    "numero_nf": nfe.get("numero"),
                    "produto": produto,
                    "produto_id": produto_id,
                    "quantidade_kg": kg,
                    "quantidade_caixas": caixas,
                    "valor_unitario": valor_unitario,
                    "valor_total": valor_total_item,
                    "natureza_operacao": nfe.get("natureza_operacao"),
                    "status_fiscal": status_fiscal,
                    "observacao": upper_text(payload.observacao),
                },
            )
    kg = money(override_kg or total_itens_kg or nfe.get("estoque_kg_entrada") or 0)
    caixas = int(override_caixas or total_itens_caixas or nfe.get("estoque_caixas_entrada") or 0)
    defaults = await logistica_defaults(db)
    produto = upper_text(nfe.get("produto") or (itens[0].get("produto") if len(itens) == 1 else f"MISTO ({len(itens)} PRODUTOS)") or defaults["produto_padrao"])
    produto_id = int(itens[0].get("produto_id") or 0) if len(itens) == 1 else None
    await db.execute(
        text(
            """
            UPDATE compras_nfe
               SET estoque_fiscal_status='ENTRADA_CONFIRMADA',
                   estoque_fisico_status='ENTRADA_CONFIRMADA',
                   estoque_kg_entrada=:kg,
                   estoque_kg_saldo=CASE WHEN COALESCE(estoque_kg_saldo, 0) <= 0 THEN :kg ELSE estoque_kg_saldo END,
                   estoque_caixas_entrada=:caixas,
                   produto_id=:produto_id,
                   produto=:produto,
                   updated_at=CURRENT_TIMESTAMP
             WHERE id=:id
            """
        ),
        {"kg": kg, "caixas": caixas, "produto_id": produto_id, "produto": produto, "id": nfe_id},
    )
    record_audit_log(
        db,
        action="estoque_entrada_confirmada",
        actor_user=current_user,
        entity_type="compras_nfe",
        entity_id=nfe_id,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={
            "chave_acesso": nfe.get("chave_acesso"),
            "numero_nf": nfe.get("numero"),
            "quantidade_kg_total": kg,
            "quantidade_caixas_total": caixas,
            "produtos": [
                {
                    "produto_id": int(item.get("produto_id") or 0) or None,
                    "produto": upper_text(item.get("produto")),
                    "quantidade_kg": money(item.get("quantidade_kg")),
                    "quantidade_caixas": int(item.get("quantidade_caixas") or 0),
                }
                for item in itens
            ],
        },
    )
    await db.commit()
    updated = await db.execute(text("SELECT * FROM compras_nfe WHERE id=:id LIMIT 1"), {"id": nfe_id})
    return row_from_mapping(dict(updated.mappings().first() or {}))


@router.post("/estoque/sincronizar-programacoes", response_model=EstoqueSyncResponse)
async def sincronizar_saidas_programacoes(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_compras_schema(db)
    result = await db.execute(
        text(
            """
            SELECT p.id, p.codigo_programacao, p.num_nf, p.nf_numero, p.nf_kg_carregado,
                   p.kg_carregado, p.nf_kg, p.nf_caixas, p.total_caixas,
                   c.id AS compra_id, c.chave_acesso, c.numero, c.valor_total, c.estoque_kg_saldo,
                   COALESCE(pi.produto_id, c.produto_id) AS produto_id,
                   COALESCE(pi.produto, c.produto, '') AS produto,
                   COUNT(pi.id) AS item_count,
                   COALESCE(SUM(pi.kg), 0) AS item_kg,
                   COALESCE(SUM(pi.qnt_caixas), 0) AS item_caixas,
                   COALESCE(MAX(cni.valor_total), 0) AS compra_item_valor_total,
                   COALESCE(MAX(cni.quantidade_kg), 0) AS compra_item_kg
              FROM programacoes p
              JOIN compras_nfe c
                ON TRIM(COALESCE(c.numero, '')) <> ''
               AND UPPER(TRIM(COALESCE(c.numero, ''))) IN (
                    UPPER(TRIM(COALESCE(p.num_nf, ''))),
                    UPPER(TRIM(COALESCE(p.nf_numero, '')))
               )
              LEFT JOIN programacao_itens pi
                ON UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) = UPPER(TRIM(COALESCE(p.codigo_programacao, '')))
              LEFT JOIN compras_nfe_itens cni
                ON cni.compra_id=c.id
               AND (
                    (COALESCE(cni.produto_id, 0) <> 0 AND COALESCE(cni.produto_id, 0)=COALESCE(pi.produto_id, c.produto_id, 0))
                    OR UPPER(TRIM(COALESCE(cni.produto, '')))=UPPER(TRIM(COALESCE(pi.produto, c.produto, '')))
               )
             WHERE COALESCE(p.codigo_programacao, '') <> ''
             GROUP BY p.id, p.codigo_programacao, p.num_nf, p.nf_numero, p.nf_kg_carregado,
                      p.kg_carregado, p.nf_kg, p.nf_caixas, p.total_caixas,
                      c.id, c.chave_acesso, c.numero, c.valor_total, c.estoque_kg_saldo,
                      COALESCE(pi.produto_id, c.produto_id), COALESCE(pi.produto, c.produto, '')
            """
        )
    )
    saidas = 0
    vinculadas = 0
    saidas_audit: list[dict[str, Any]] = []
    for row in result.mappings().all():
        vinculadas += 1
        codigo = upper_text(row.get("codigo_programacao"))
        compra_id = int(row.get("compra_id") or 0)
        if not codigo or not compra_id:
            continue
        has_items = int(row.get("item_count") or 0) > 0
        kg = money(row.get("item_kg") if has_items else (row.get("nf_kg_carregado") or row.get("kg_carregado") or row.get("nf_kg") or 0))
        caixas = int(row.get("item_caixas") if has_items else (row.get("nf_caixas") or row.get("total_caixas") or 0))
        if kg <= 0 and caixas <= 0:
            continue
        valor_total = money(row.get("compra_item_valor_total") or row.get("valor_total"))
        defaults = await logistica_defaults(db)
        produto = upper_text(row.get("produto") or defaults["produto_padrao"])
        produto_id = int(row.get("produto_id") or 0) or await ensure_produto_catalogo(db, produto)
        compra_kg = money(row.get("compra_item_kg") or row.get("estoque_kg_saldo") or kg)
        valor_unitario = money(valor_total / compra_kg) if compra_kg > 0 else 0
        baixa_fisica_criada = False
        for tipo_estoque, status_fiscal in [("FISICO", "NAO_APLICAVEL"), ("FISCAL", "PENDENTE_SEFAZ")]:
            exists = await db.execute(
                text(
                    """
                    SELECT id FROM estoque_movimentos
                   WHERE tipo_estoque=:tipo_estoque AND tipo_movimento='SAIDA'
                     AND origem='PROGRAMACAO' AND codigo_programacao=:codigo
                     AND COALESCE(produto_id, 0)=COALESCE(:produto_id, 0)
                     AND UPPER(COALESCE(produto, ''))=UPPER(:produto)
                   LIMIT 1
                    """
                ),
                {"tipo_estoque": tipo_estoque, "codigo": codigo, "produto_id": produto_id or 0, "produto": produto},
            )
            if exists.scalar_one_or_none():
                continue
            await db.execute(
                text(
                    """
                    INSERT INTO estoque_movimentos (
                        tipo_estoque, tipo_movimento, origem, origem_id, chave_acesso, numero_nf,
                        codigo_programacao, produto_id, produto, quantidade_kg, quantidade_caixas,
                        valor_unitario, valor_total, natureza_operacao, status_fiscal, observacao
                    ) VALUES (
                        :tipo_estoque, 'SAIDA', 'PROGRAMACAO', :origem_id, :chave_acesso, :numero_nf,
                        :codigo_programacao, :produto_id, :produto, :quantidade_kg, :quantidade_caixas,
                        :valor_unitario, :valor_total, 'SAIDA POR PROGRAMACAO', :status_fiscal, :observacao
                    )
                    """
                ),
                {
                    "tipo_estoque": tipo_estoque,
                    "origem_id": row.get("id"),
                    "chave_acesso": row.get("chave_acesso"),
                    "numero_nf": row.get("numero"),
                    "codigo_programacao": codigo,
                    "produto_id": produto_id,
                    "produto": produto,
                    "quantidade_kg": kg,
                    "quantidade_caixas": caixas,
                    "valor_unitario": valor_unitario,
                    "valor_total": money(kg * valor_unitario),
                    "status_fiscal": status_fiscal,
                    "observacao": "Saida fiscal pendente de autorizacao SEFAZ" if tipo_estoque == "FISCAL" else "Baixa fisica por programacao",
                },
            )
            saidas += 1
            if tipo_estoque == "FISICO":
                baixa_fisica_criada = True
            saidas_audit.append(
                {
                    "codigo_programacao": codigo,
                    "produto_id": produto_id,
                    "produto": produto,
                    "tipo_estoque": tipo_estoque,
                    "quantidade_kg": kg,
                    "quantidade_caixas": caixas,
                    "numero_nf": row.get("numero"),
                }
            )
        if baixa_fisica_criada:
            await db.execute(
                text(
                    """
                    UPDATE compras_nfe
                       SET estoque_kg_saldo=MAX(COALESCE(estoque_kg_saldo, 0) - :kg, 0),
                           updated_at=CURRENT_TIMESTAMP
                     WHERE id=:id
                    """
                ),
                {"kg": kg, "id": compra_id},
            )
    record_audit_log(
        db,
        action="estoque_saidas_programacoes_sincronizadas",
        actor_user=current_user,
        entity_type="estoque_movimentos",
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={
            "saidas_criadas": saidas,
            "programacoes_vinculadas": vinculadas,
            "amostra_saidas": saidas_audit[:50],
        },
    )
    await db.commit()
    return EstoqueSyncResponse(
        ok=True,
        message="Programacoes vinculadas a notas foram sincronizadas com o estoque.",
        saidas_criadas=saidas,
        programacoes_vinculadas=vinculadas,
    )


@router.get("/nfe/{nfe_id}/xml")
async def download_nfe_xml(
    nfe_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    await ensure_compras_schema(db)
    result = await db.execute(text("SELECT chave_acesso, xml_path FROM compras_nfe WHERE id=:id LIMIT 1"), {"id": nfe_id})
    row = result.mappings().first()
    if not row:
        raise HTTPException(status_code=404, detail="NF-e nao encontrada.")
    path = Path(str(row.get("xml_path") or ""))
    if not path.exists():
        raise HTTPException(status_code=404, detail="XML nao encontrado no armazenamento local.")
    return FileResponse(path, media_type="application/xml", filename=f"{row.get('chave_acesso') or nfe_id}.xml")
