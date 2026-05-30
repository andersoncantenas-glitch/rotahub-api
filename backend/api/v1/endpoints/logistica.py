"""Configuracao operacional para adaptar o sistema a diferentes logisticas."""
from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query, Request
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from backend.api.v1.endpoints.programacao import upper_text
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.user import User
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()

DEFAULT_PERFIL = {
    "codigo": "DISTRIBUICAO_GERAL",
    "nome": "Distribuicao geral",
    "descricao": "Perfil logistico neutro para rotas com carga, volumes, ocorrencias operacionais e controle fiscal.",
    "produto_padrao": "CARGA",
    "unidade_padrao": "KG",
    "embalagem_label": "Volumes",
    "perda_label": "Ocorrencias operacionais",
    "quantidade_embalagem_label": "Qtd. por volume",
    "usa_mortalidade": 1,
    "usa_aves_por_caixa": 0,
    "usa_nota_fiscal_motorista": 1,
    "usa_estoque_fisico": 1,
    "usa_estoque_fiscal": 1,
}

LEGACY_PERDA_LABELS = {
    "MORTALIDADE",
    "MORTALIDADES",
    "PERDA",
    "PERDAS",
    "AVES MORTAS",
    "MORTES",
}


def normalize_perda_label(value: Any) -> str:
    label = str(value or "").strip()
    key = upper_text(label)
    if not label:
        return "Ocorrencias operacionais"
    if "MORTAL" in key or key in LEGACY_PERDA_LABELS:
        return "Ocorrencias operacionais"
    return label

PERFIS_SEED = [
    DEFAULT_PERFIL,
    {
        "codigo": "LOGISTICA_FRIOS",
        "nome": "Logistica de frios",
        "descricao": "Perfil para cargas refrigeradas ou congeladas, com controle por peso, volume e ocorrencias.",
        "produto_padrao": "CARGA REFRIGERADA",
        "unidade_padrao": "KG",
        "embalagem_label": "Volumes",
        "perda_label": "Ocorrencias operacionais",
        "quantidade_embalagem_label": "Qtd. por embalagem",
        "usa_mortalidade": 1,
        "usa_aves_por_caixa": 0,
        "usa_nota_fiscal_motorista": 1,
        "usa_estoque_fisico": 1,
        "usa_estoque_fiscal": 1,
    },
    {
        "codigo": "GRANJA",
        "nome": "Logistica de granja",
        "descricao": "Perfil para operacao de granja, mantendo termos configuraveis para carga, volumes e ocorrencias.",
        "produto_padrao": "CARGA",
        "unidade_padrao": "KG",
        "embalagem_label": "Caixas",
        "perda_label": "Ocorrencias operacionais",
        "quantidade_embalagem_label": "Itens por caixa",
        "usa_mortalidade": 1,
        "usa_aves_por_caixa": 1,
        "usa_nota_fiscal_motorista": 1,
        "usa_estoque_fisico": 1,
        "usa_estoque_fiscal": 1,
    },
    {
        "codigo": "OVOS",
        "nome": "Logistica de ovos",
        "descricao": "Perfil para distribuicao de ovos com controle por duzia, caixa, bandeja ou unidade.",
        "produto_padrao": "PRODUTO",
        "unidade_padrao": "DZ",
        "embalagem_label": "Caixas",
        "perda_label": "Ocorrencias operacionais",
        "quantidade_embalagem_label": "Qtd. por caixa",
        "usa_mortalidade": 1,
        "usa_aves_por_caixa": 0,
        "usa_nota_fiscal_motorista": 1,
        "usa_estoque_fisico": 1,
        "usa_estoque_fiscal": 1,
    },
    {
        "codigo": "DISTRIBUIDORA",
        "nome": "Distribuidora geral",
        "descricao": "Perfil para distribuidores de alimentos, limpeza, ferramentas e mercadorias em geral.",
        "produto_padrao": "PRODUTO",
        "unidade_padrao": "UN",
        "embalagem_label": "Embalagens",
        "perda_label": "Ocorrencias operacionais",
        "quantidade_embalagem_label": "Qtd. por embalagem",
        "usa_mortalidade": 1,
        "usa_aves_por_caixa": 0,
        "usa_nota_fiscal_motorista": 1,
        "usa_estoque_fisico": 1,
        "usa_estoque_fiscal": 1,
    },
    {
        "codigo": "BEBIDAS_REFRIGERADAS",
        "nome": "Distribuidora de bebidas e congelados",
        "descricao": "Perfil para bebidas e produtos refrigerados/congelados com controle por volume, peso ou unidade.",
        "produto_padrao": "PRODUTO",
        "unidade_padrao": "KG",
        "embalagem_label": "Volumes",
        "perda_label": "Ocorrencias operacionais",
        "quantidade_embalagem_label": "Qtd. por volume",
        "usa_mortalidade": 1,
        "usa_aves_por_caixa": 0,
        "usa_nota_fiscal_motorista": 1,
        "usa_estoque_fisico": 1,
        "usa_estoque_fiscal": 1,
    },
]

UNIDADES_SEED = [
    ("KG", "Quilograma", "PESO"),
    ("CX", "Caixa", "EMBALAGEM"),
    ("UN", "Unidade", "UNIDADE"),
    ("PC", "Peca", "UNIDADE"),
    ("LT", "Litro", "VOLUME"),
    ("TON", "Tonelada", "PESO"),
    ("M2", "Metro quadrado", "AREA"),
    ("DZ", "Duzia", "UNIDADE"),
    ("PALLET", "Pallet", "EMBALAGEM"),
]

OCORRENCIAS_SEED = [
    ("OCORRENCIA_OPERACIONAL", "Ocorrencia operacional", "PERDA", "DISTRIBUICAO_GERAL"),
    ("AVARIA", "Avaria", "PERDA", "DISTRIBUICAO_GERAL"),
    ("QUEBRA", "Quebra", "PERDA", "DISTRIBUICAO_GERAL"),
    ("DEVOLUCAO", "Devolucao", "RETORNO", "DISTRIBUICAO_GERAL"),
    ("FALTA", "Falta", "DIVERGENCIA", "DISTRIBUICAO_GERAL"),
    ("SOBRA", "Sobra", "DIVERGENCIA", "DISTRIBUICAO_GERAL"),
]


class LogisticaConfigPayload(BaseModel):
    company_id: int | None = Field(default=None, ge=1)
    perfil_codigo: str = Field(default="DISTRIBUICAO_GERAL", max_length=80)
    produto_padrao: str = Field(default="CARGA", max_length=120)
    unidade_padrao: str = Field(default="KG", max_length=20)
    embalagem_label: str = Field(default="Volumes", max_length=80)
    perda_label: str = Field(default="Ocorrencias operacionais", max_length=80)
    quantidade_embalagem_label: str = Field(default="Qtd. por volume", max_length=80)
    usa_mortalidade: int = Field(default=1, ge=0, le=1)
    usa_aves_por_caixa: int = Field(default=1, ge=0, le=1)
    usa_nota_fiscal_motorista: int = Field(default=1, ge=0, le=1)
    usa_estoque_fisico: int = Field(default=1, ge=0, le=1)
    usa_estoque_fiscal: int = Field(default=1, ge=0, le=1)

    @field_validator(
        "perfil_codigo",
        "produto_padrao",
        "unidade_padrao",
        "embalagem_label",
        "perda_label",
        "quantidade_embalagem_label",
        mode="before",
    )
    @classmethod
    def normalize_text(cls, value):
        return str(value or "").strip()


class LogisticaPerfilPayload(BaseModel):
    codigo: str = Field(max_length=80)
    nome: str = Field(max_length=120)
    descricao: str = Field(default="", max_length=500)
    produto_padrao: str = Field(default="PRODUTO", max_length=120)
    unidade_padrao: str = Field(default="UN", max_length=20)
    embalagem_label: str = Field(default="Embalagens", max_length=80)
    perda_label: str = Field(default="Ocorrencias operacionais", max_length=80)
    quantidade_embalagem_label: str = Field(default="Qtd. por embalagem", max_length=80)
    usa_mortalidade: int = Field(default=0, ge=0, le=1)
    usa_aves_por_caixa: int = Field(default=0, ge=0, le=1)
    usa_nota_fiscal_motorista: int = Field(default=1, ge=0, le=1)
    usa_estoque_fisico: int = Field(default=1, ge=0, le=1)
    usa_estoque_fiscal: int = Field(default=1, ge=0, le=1)
    status: str = Field(default="ATIVO", max_length=20)

    @field_validator(
        "codigo",
        "nome",
        "descricao",
        "produto_padrao",
        "unidade_padrao",
        "embalagem_label",
        "perda_label",
        "quantidade_embalagem_label",
        "status",
        mode="before",
    )
    @classmethod
    def normalize_text(cls, value):
        return str(value or "").strip()


class LogisticaUnidadePayload(BaseModel):
    codigo: str = Field(max_length=20)
    nome: str = Field(max_length=120)
    tipo: str = Field(default="UNIDADE", max_length=40)
    status: str = Field(default="ATIVO", max_length=20)

    @field_validator("codigo", "nome", "tipo", "status", mode="before")
    @classmethod
    def normalize_text(cls, value):
        return str(value or "").strip()


class LogisticaOcorrenciaPayload(BaseModel):
    codigo: str = Field(max_length=80)
    nome: str = Field(max_length=120)
    categoria: str = Field(default="PERDA", max_length=40)
    perfil_codigo: str = Field(default="DISTRIBUICAO_GERAL", max_length=80)
    status: str = Field(default="ATIVO", max_length=20)

    @field_validator("codigo", "nome", "categoria", "perfil_codigo", "status", mode="before")
    @classmethod
    def normalize_text(cls, value):
        return str(value or "").strip()


async def ensure_logistica_schema(db: AsyncSession) -> None:
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS perfis_operacionais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE,
                nome TEXT NOT NULL,
                descricao TEXT,
                produto_padrao TEXT DEFAULT 'CARGA',
                unidade_padrao TEXT DEFAULT 'KG',
                embalagem_label TEXT DEFAULT 'Volumes',
                perda_label TEXT DEFAULT 'Ocorrencias operacionais',
                quantidade_embalagem_label TEXT DEFAULT 'Qtd. por volume',
                usa_mortalidade INTEGER DEFAULT 1,
                usa_aves_por_caixa INTEGER DEFAULT 1,
                usa_nota_fiscal_motorista INTEGER DEFAULT 1,
                usa_estoque_fisico INTEGER DEFAULT 1,
                usa_estoque_fiscal INTEGER DEFAULT 1,
                status TEXT DEFAULT 'ATIVO',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS empresa_configuracao_logistica (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_id INTEGER NOT NULL UNIQUE,
                perfil_codigo TEXT NOT NULL DEFAULT 'DISTRIBUICAO_GERAL',
                produto_padrao TEXT DEFAULT 'CARGA',
                unidade_padrao TEXT DEFAULT 'KG',
                embalagem_label TEXT DEFAULT 'Volumes',
                perda_label TEXT DEFAULT 'Ocorrencias operacionais',
                quantidade_embalagem_label TEXT DEFAULT 'Qtd. por volume',
                usa_mortalidade INTEGER DEFAULT 1,
                usa_aves_por_caixa INTEGER DEFAULT 1,
                usa_nota_fiscal_motorista INTEGER DEFAULT 1,
                usa_estoque_fisico INTEGER DEFAULT 1,
                usa_estoque_fiscal INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS unidades_medida_logistica (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE,
                nome TEXT NOT NULL,
                tipo TEXT DEFAULT 'UNIDADE',
                status TEXT DEFAULT 'ATIVO'
            )
            """
        )
    )
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS tipos_ocorrencia_logistica (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE,
                nome TEXT NOT NULL,
                categoria TEXT DEFAULT 'PERDA',
                perfil_codigo TEXT DEFAULT 'DISTRIBUICAO_GERAL',
                status TEXT DEFAULT 'ATIVO'
            )
            """
        )
    )
    for perfil in PERFIS_SEED:
        exists = await db.execute(text("SELECT id FROM perfis_operacionais WHERE codigo=:codigo LIMIT 1"), {"codigo": perfil["codigo"]})
        if exists.scalar_one_or_none():
            continue
        await db.execute(
            text(
                """
                INSERT INTO perfis_operacionais (
                    codigo, nome, descricao, produto_padrao, unidade_padrao, embalagem_label,
                    perda_label, quantidade_embalagem_label, usa_mortalidade, usa_aves_por_caixa,
                    usa_nota_fiscal_motorista, usa_estoque_fisico, usa_estoque_fiscal
                ) VALUES (
                    :codigo, :nome, :descricao, :produto_padrao, :unidade_padrao, :embalagem_label,
                    :perda_label, :quantidade_embalagem_label, :usa_mortalidade, :usa_aves_por_caixa,
                    :usa_nota_fiscal_motorista, :usa_estoque_fisico, :usa_estoque_fiscal
                )
                """
            ),
            perfil,
        )
    for codigo, nome, tipo in UNIDADES_SEED:
        exists = await db.execute(text("SELECT id FROM unidades_medida_logistica WHERE codigo=:codigo LIMIT 1"), {"codigo": codigo})
        if not exists.scalar_one_or_none():
            await db.execute(
                text("INSERT INTO unidades_medida_logistica (codigo, nome, tipo) VALUES (:codigo, :nome, :tipo)"),
                {"codigo": codigo, "nome": nome, "tipo": tipo},
            )
    for codigo, nome, categoria, perfil_codigo in OCORRENCIAS_SEED:
        exists = await db.execute(text("SELECT id FROM tipos_ocorrencia_logistica WHERE codigo=:codigo LIMIT 1"), {"codigo": codigo})
        if not exists.scalar_one_or_none():
            await db.execute(
                text(
                    """
                    INSERT INTO tipos_ocorrencia_logistica (codigo, nome, categoria, perfil_codigo)
                    VALUES (:codigo, :nome, :categoria, :perfil_codigo)
                    """
                ),
                {"codigo": codigo, "nome": nome, "categoria": categoria, "perfil_codigo": perfil_codigo},
            )
    await db.commit()


async def default_company_id(db: AsyncSession) -> int:
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                status TEXT DEFAULT 'active',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    result = await db.execute(text("SELECT id FROM companies ORDER BY id ASC LIMIT 1"))
    company_id = result.scalar_one_or_none()
    if company_id:
        return int(company_id)
    await db.execute(
        text("INSERT INTO companies (code, name, status) VALUES ('default', 'Empresa Principal', 'active')")
    )
    await db.commit()
    result = await db.execute(text("SELECT id FROM companies ORDER BY id ASC LIMIT 1"))
    return int(result.scalar_one())


async def load_config(db: AsyncSession, company_id: int) -> dict[str, Any]:
    result = await db.execute(
        text("SELECT * FROM empresa_configuracao_logistica WHERE company_id=:company_id LIMIT 1"),
        {"company_id": company_id},
    )
    row = result.mappings().first()
    if row:
        config = dict(row)
        normalized = normalize_perda_label(config.get("perda_label"))
        if normalized != config.get("perda_label"):
            config["perda_label"] = normalized
            await db.execute(
                text(
                    """
                    UPDATE empresa_configuracao_logistica
                    SET perda_label=:perda_label, updated_at=CURRENT_TIMESTAMP
                    WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id, "perda_label": normalized},
            )
            await db.commit()
        return config
    payload = dict(DEFAULT_PERFIL)
    await db.execute(
        text(
            """
            INSERT INTO empresa_configuracao_logistica (
                company_id, perfil_codigo, produto_padrao, unidade_padrao, embalagem_label,
                perda_label, quantidade_embalagem_label, usa_mortalidade, usa_aves_por_caixa,
                usa_nota_fiscal_motorista, usa_estoque_fisico, usa_estoque_fiscal
            ) VALUES (
                :company_id, :perfil_codigo, :produto_padrao, :unidade_padrao, :embalagem_label,
                :perda_label, :quantidade_embalagem_label, :usa_mortalidade, :usa_aves_por_caixa,
                :usa_nota_fiscal_motorista, :usa_estoque_fisico, :usa_estoque_fiscal
            )
            """
        ),
        {"company_id": company_id, "perfil_codigo": payload["codigo"], **payload},
    )
    await db.commit()
    return await load_config(db, company_id)


@router.get("/config")
async def get_logistica_config(
    company_id: int | None = Query(default=None, ge=1),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    await ensure_logistica_schema(db)
    cid = int(company_id or await default_company_id(db))
    config = await load_config(db, cid)
    perfis = (await db.execute(text("SELECT * FROM perfis_operacionais WHERE status='ATIVO' ORDER BY nome"))).mappings().all()
    unidades = (await db.execute(text("SELECT * FROM unidades_medida_logistica WHERE status='ATIVO' ORDER BY codigo"))).mappings().all()
    ocorrencias = (await db.execute(text("SELECT * FROM tipos_ocorrencia_logistica WHERE status='ATIVO' ORDER BY nome"))).mappings().all()
    return {
        "company_id": cid,
        "config": config,
        "perfis": [dict(row) for row in perfis],
        "unidades": [dict(row) for row in unidades],
        "ocorrencias": [dict(row) for row in ocorrencias],
    }


@router.put("/config")
async def update_logistica_config(
    payload: LogisticaConfigPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_logistica_schema(db)
    company_id = int(payload.company_id or await default_company_id(db))
    perfil_codigo = upper_text(payload.perfil_codigo or "DISTRIBUICAO_GERAL")
    perfil = await db.execute(text("SELECT codigo FROM perfis_operacionais WHERE codigo=:codigo LIMIT 1"), {"codigo": perfil_codigo})
    if not perfil.scalar_one_or_none():
        raise HTTPException(status_code=422, detail="Perfil operacional invalido.")
    unidade = upper_text(payload.unidade_padrao or "KG")
    await db.execute(
        text(
            """
            INSERT INTO empresa_configuracao_logistica (
                company_id, perfil_codigo, produto_padrao, unidade_padrao, embalagem_label,
                perda_label, quantidade_embalagem_label, usa_mortalidade, usa_aves_por_caixa,
                usa_nota_fiscal_motorista, usa_estoque_fisico, usa_estoque_fiscal, updated_at
            ) VALUES (
                :company_id, :perfil_codigo, :produto_padrao, :unidade_padrao, :embalagem_label,
                :perda_label, :quantidade_embalagem_label, :usa_mortalidade, :usa_aves_por_caixa,
                :usa_nota_fiscal_motorista, :usa_estoque_fisico, :usa_estoque_fiscal, CURRENT_TIMESTAMP
            )
            ON CONFLICT(company_id) DO UPDATE SET
                perfil_codigo=excluded.perfil_codigo,
                produto_padrao=excluded.produto_padrao,
                unidade_padrao=excluded.unidade_padrao,
                embalagem_label=excluded.embalagem_label,
                perda_label=excluded.perda_label,
                quantidade_embalagem_label=excluded.quantidade_embalagem_label,
                usa_mortalidade=excluded.usa_mortalidade,
                usa_aves_por_caixa=excluded.usa_aves_por_caixa,
                usa_nota_fiscal_motorista=excluded.usa_nota_fiscal_motorista,
                usa_estoque_fisico=excluded.usa_estoque_fisico,
                usa_estoque_fiscal=excluded.usa_estoque_fiscal,
                updated_at=CURRENT_TIMESTAMP
            """
        ),
        {
            "company_id": company_id,
            "perfil_codigo": perfil_codigo,
            "produto_padrao": upper_text(payload.produto_padrao or "PRODUTO"),
            "unidade_padrao": unidade,
            "embalagem_label": payload.embalagem_label or "Embalagens",
            "perda_label": normalize_perda_label(payload.perda_label),
            "quantidade_embalagem_label": payload.quantidade_embalagem_label or "Qtd. por volume",
            "usa_mortalidade": int(payload.usa_mortalidade),
            "usa_aves_por_caixa": int(payload.usa_aves_por_caixa),
            "usa_nota_fiscal_motorista": int(payload.usa_nota_fiscal_motorista),
            "usa_estoque_fisico": int(payload.usa_estoque_fisico),
            "usa_estoque_fiscal": int(payload.usa_estoque_fiscal),
        },
    )
    record_audit_log(
        db,
        action="logistica_config_atualizada",
        actor_user=current_user,
        entity_type="empresa_configuracao_logistica",
        entity_id=company_id,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"company_id": company_id, "perfil_codigo": perfil_codigo, "unidade_padrao": unidade},
    )
    await db.commit()
    return {"ok": True, "config": await load_config(db, company_id)}


@router.post("/perfis")
async def upsert_logistica_perfil(
    payload: LogisticaPerfilPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_logistica_schema(db)
    codigo = upper_text(payload.codigo).replace(" ", "_")
    nome = payload.nome.strip()
    if not codigo or not nome:
        raise HTTPException(status_code=422, detail="Informe codigo e nome do perfil.")
    unidade = upper_text(payload.unidade_padrao or "UN")
    unidade_exists = await db.execute(text("SELECT codigo FROM unidades_medida_logistica WHERE codigo=:codigo LIMIT 1"), {"codigo": unidade})
    if not unidade_exists.scalar_one_or_none():
        await db.execute(
            text("INSERT INTO unidades_medida_logistica (codigo, nome, tipo) VALUES (:codigo, :nome, 'UNIDADE')"),
            {"codigo": unidade, "nome": unidade},
        )
    await db.execute(
        text(
            """
            INSERT INTO perfis_operacionais (
                codigo, nome, descricao, produto_padrao, unidade_padrao, embalagem_label,
                perda_label, quantidade_embalagem_label, usa_mortalidade, usa_aves_por_caixa,
                usa_nota_fiscal_motorista, usa_estoque_fisico, usa_estoque_fiscal, status, updated_at
            ) VALUES (
                :codigo, :nome, :descricao, :produto_padrao, :unidade_padrao, :embalagem_label,
                :perda_label, :quantidade_embalagem_label, :usa_mortalidade, :usa_aves_por_caixa,
                :usa_nota_fiscal_motorista, :usa_estoque_fisico, :usa_estoque_fiscal, :status, CURRENT_TIMESTAMP
            )
            ON CONFLICT(codigo) DO UPDATE SET
                nome=excluded.nome,
                descricao=excluded.descricao,
                produto_padrao=excluded.produto_padrao,
                unidade_padrao=excluded.unidade_padrao,
                embalagem_label=excluded.embalagem_label,
                perda_label=excluded.perda_label,
                quantidade_embalagem_label=excluded.quantidade_embalagem_label,
                usa_mortalidade=excluded.usa_mortalidade,
                usa_aves_por_caixa=excluded.usa_aves_por_caixa,
                usa_nota_fiscal_motorista=excluded.usa_nota_fiscal_motorista,
                usa_estoque_fisico=excluded.usa_estoque_fisico,
                usa_estoque_fiscal=excluded.usa_estoque_fiscal,
                status=excluded.status,
                updated_at=CURRENT_TIMESTAMP
            """
        ),
        {
            "codigo": codigo,
            "nome": nome,
            "descricao": payload.descricao,
            "produto_padrao": upper_text(payload.produto_padrao or "PRODUTO"),
            "unidade_padrao": unidade,
            "embalagem_label": payload.embalagem_label or "Embalagens",
            "perda_label": normalize_perda_label(payload.perda_label),
            "quantidade_embalagem_label": payload.quantidade_embalagem_label or "Qtd. por embalagem",
            "usa_mortalidade": int(payload.usa_mortalidade),
            "usa_aves_por_caixa": int(payload.usa_aves_por_caixa),
            "usa_nota_fiscal_motorista": int(payload.usa_nota_fiscal_motorista),
            "usa_estoque_fisico": int(payload.usa_estoque_fisico),
            "usa_estoque_fiscal": int(payload.usa_estoque_fiscal),
            "status": upper_text(payload.status or "ATIVO"),
        },
    )
    record_audit_log(
        db,
        action="logistica_perfil_salvo",
        actor_user=current_user,
        entity_type="perfis_operacionais",
        entity_id=codigo,
        severity="warning",
        ip_address=client_ip_from_request(request),
        metadata={"codigo": codigo, "unidade_padrao": unidade},
    )
    await db.commit()
    return {"ok": True, "codigo": codigo}


@router.post("/unidades")
async def upsert_logistica_unidade(
    payload: LogisticaUnidadePayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_logistica_schema(db)
    codigo = upper_text(payload.codigo)
    nome = payload.nome.strip()
    if not codigo or not nome:
        raise HTTPException(status_code=422, detail="Informe codigo e nome da unidade.")
    await db.execute(
        text(
            """
            INSERT INTO unidades_medida_logistica (codigo, nome, tipo, status)
            VALUES (:codigo, :nome, :tipo, :status)
            ON CONFLICT(codigo) DO UPDATE SET
                nome=excluded.nome,
                tipo=excluded.tipo,
                status=excluded.status
            """
        ),
        {
            "codigo": codigo,
            "nome": nome,
            "tipo": upper_text(payload.tipo or "UNIDADE"),
            "status": upper_text(payload.status or "ATIVO"),
        },
    )
    record_audit_log(
        db,
        action="logistica_unidade_salva",
        actor_user=current_user,
        entity_type="unidades_medida_logistica",
        entity_id=codigo,
        severity="info",
        ip_address=client_ip_from_request(request),
        metadata={"codigo": codigo},
    )
    await db.commit()
    return {"ok": True, "codigo": codigo}


@router.post("/ocorrencias")
async def upsert_logistica_ocorrencia(
    payload: LogisticaOcorrenciaPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    await ensure_logistica_schema(db)
    codigo = upper_text(payload.codigo).replace(" ", "_")
    nome = payload.nome.strip()
    perfil_codigo = upper_text(payload.perfil_codigo or "DISTRIBUICAO_GERAL")
    if not codigo or not nome:
        raise HTTPException(status_code=422, detail="Informe codigo e nome da ocorrencia.")
    perfil = await db.execute(text("SELECT codigo FROM perfis_operacionais WHERE codigo=:codigo LIMIT 1"), {"codigo": perfil_codigo})
    if not perfil.scalar_one_or_none():
        raise HTTPException(status_code=422, detail="Perfil operacional invalido.")
    await db.execute(
        text(
            """
            INSERT INTO tipos_ocorrencia_logistica (codigo, nome, categoria, perfil_codigo, status)
            VALUES (:codigo, :nome, :categoria, :perfil_codigo, :status)
            ON CONFLICT(codigo) DO UPDATE SET
                nome=excluded.nome,
                categoria=excluded.categoria,
                perfil_codigo=excluded.perfil_codigo,
                status=excluded.status
            """
        ),
        {
            "codigo": codigo,
            "nome": nome,
            "categoria": upper_text(payload.categoria or "PERDA"),
            "perfil_codigo": perfil_codigo,
            "status": upper_text(payload.status or "ATIVO"),
        },
    )
    record_audit_log(
        db,
        action="logistica_ocorrencia_salva",
        actor_user=current_user,
        entity_type="tipos_ocorrencia_logistica",
        entity_id=codigo,
        severity="info",
        ip_address=client_ip_from_request(request),
        metadata={"codigo": codigo, "perfil_codigo": perfil_codigo},
    )
    await db.commit()
    return {"ok": True, "codigo": codigo}
