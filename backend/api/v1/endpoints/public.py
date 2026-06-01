# backend/api/v1/endpoints/public.py
"""
Public marketing and signup endpoints.

These endpoints expose plan cards and provision the initial trial account for
the public landing page.
"""
from __future__ import annotations

import json
import re
import time
import unicodedata
from collections import defaultdict
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Request, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from backend.config.database import get_db
from backend.services.auth import get_password_hash

router = APIRouter()

PUBLIC_SIGNUP_MAX_ATTEMPTS = 5
PUBLIC_SIGNUP_WINDOW_SECONDS = 10 * 60
PUBLIC_SIGNUP_ATTEMPTS: dict[str, list[float]] = defaultdict(list)
TRIAL_PLAN_PRIORITY = ("enterprise", "professional", "growth", "starter")


PLAN_FALLBACKS = [
    {
        "code": "starter",
        "name": "Inicial 5 Veículos",
        "monthly_price": 199.0,
        "vehicle_limit": 5,
        "user_limit": 6,
        "description": "Para começar organizado: planejamento, recebimentos, custos e app motorista em até 5 veículos.",
        "summary": "O básico bem feito para parar de controlar rota no improviso.",
        "features": ["Cadastros da operação", "Planejamento de rotas", "Recebimentos e custos", "App do motorista"],
    },
    {
        "code": "growth",
        "name": "Crescimento 10 Veículos",
        "monthly_price": 399.0,
        "vehicle_limit": 10,
        "user_limit": 15,
        "description": "Para controlar ocorrências operacionais, rotas, financeiro básico e análise de custos em até 10 veículos.",
        "summary": "Mais visão para saber onde a rota está dando lucro ou prejuízo.",
        "features": ["Tudo do Inicial", "Ocorrências operacionais", "Rotas", "Análise de custos", "Financeiro básico"],
    },
    {
        "code": "professional",
        "name": "Profissional 15 Veículos",
        "monthly_price": 699.0,
        "vehicle_limit": 15,
        "user_limit": 30,
        "description": "Para operações mais exigentes: escala, relatórios avançados, rotas e controles completos em até 15 veículos.",
        "summary": "Controle completo para gestor acompanhar equipe, frota e resultado.",
        "features": ["Tudo do Crescimento", "Escala", "Relatórios avançados", "Controle por perfis", "Gestão completa"],
    },
    {
        "code": "enterprise",
        "name": "Empresarial Mais Veículos",
        "monthly_price": 999.0,
        "vehicle_limit": None,
        "user_limit": None,
        "description": "Para empresas com mais de 15 veículos, API, suporte prioritário e contrato ajustado à operação.",
        "summary": "Plano sob medida para operação maior, com implantação acompanhada.",
        "features": ["Tudo do Profissional", "Mais de 15 veículos", "API", "Contrato customizado", "Suporte prioritário"],
    },
]

PUBLIC_PLAN_CODES = ("starter", "growth", "professional", "enterprise")

FEATURE_PUBLIC_LABELS = {
    "cadastros": "Cadastros da operacao",
    "importar_vendas": "Importar pedidos",
    "programacao": "Planejamento de rotas",
    "recebimentos": "Recebimentos",
    "despesas": "Custos e despesas",
    "mortalidade": "Ocorrencias operacionais",
    "centro_custos": "Analise de custos",
    "relatorios": "Relatorios",
    "rotas": "Rotas",
    "escala": "Escala",
    "app_motorista": "App do motorista",
    "realtime_tracking": "Rastreamento",
    "financial_reports": "Financeiro",
    "advanced_reports": "Relatorios avancados",
    "api_access": "API",
    "custom_contract": "Contrato customizado",
    "priority_support": "Suporte prioritario",
}

PLAN_PUBLIC_COPY = {item["code"]: item for item in PLAN_FALLBACKS}


class PublicSignupPayload(BaseModel):
    name: str = Field(min_length=2, max_length=160)
    document: str = Field(min_length=11, max_length=20)
    email: str = Field(min_length=5, max_length=180)
    phone: str = Field(min_length=8, max_length=30)
    company: str = Field(min_length=2, max_length=180)
    plan_code: str = Field(min_length=1, max_length=80)
    username: str = Field(min_length=3, max_length=80)
    password: str = Field(min_length=6, max_length=128)
    message: str | None = Field(default="", max_length=600)

    @field_validator("name", "document", "email", "phone", "company", "plan_code", "username", "message", mode="before")
    @classmethod
    def strip_text(cls, value: Any) -> str:
        return str(value or "").strip()

    @field_validator("document")
    @classmethod
    def validate_document(cls, value: str) -> str:
        digits = re.sub(r"\D+", "", value)
        if len(digits) not in {11, 14}:
            raise ValueError("Informe CPF ou CNPJ valido.")
        return digits

    @field_validator("email")
    @classmethod
    def validate_email(cls, value: str) -> str:
        email = value.strip().lower()
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
            raise ValueError("Informe um e-mail valido.")
        return email

    @field_validator("plan_code")
    @classmethod
    def normalize_plan(cls, value: str) -> str:
        return value.strip().lower()

    @field_validator("username")
    @classmethod
    def validate_username(cls, value: str) -> str:
        username = value.strip()
        if not re.fullmatch(r"[A-Za-z0-9._@-]+", username):
            raise ValueError("Use somente letras, numeros, ponto, traco, underline ou @ no login.")
        return username


def company_code(value: str, lead_id: int) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or "empresa")).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-z0-9]+", "-", normalized.lower()).strip("-")[:42] or "empresa"
    return f"{slug}-{int(lead_id)}"


def public_client_ip(request: Request) -> str:
    forwarded = str(request.headers.get("x-forwarded-for") or "").strip()
    if forwarded:
        return forwarded.split(",", 1)[0].strip()
    real_ip = str(request.headers.get("x-real-ip") or "").strip()
    if real_ip:
        return real_ip
    return str(request.client.host if request.client else "unknown")


def enforce_public_signup_rate_limit(request: Request) -> str:
    ip_address = public_client_ip(request)
    now = time.time()
    attempts = [
        attempted_at
        for attempted_at in PUBLIC_SIGNUP_ATTEMPTS[ip_address]
        if now - attempted_at < PUBLIC_SIGNUP_WINDOW_SECONDS
    ]
    if len(attempts) >= PUBLIC_SIGNUP_MAX_ATTEMPTS:
        PUBLIC_SIGNUP_ATTEMPTS[ip_address] = attempts
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail="Muitas tentativas de cadastro. Aguarde alguns minutos e tente novamente.",
        )
    attempts.append(now)
    PUBLIC_SIGNUP_ATTEMPTS[ip_address] = attempts
    return ip_address


async def best_trial_plan_id(db: AsyncSession) -> int:
    result = await db.execute(
        text(
            """
            SELECT id, code
              FROM plans
             WHERE status='active'
               AND code IN ('enterprise', 'professional', 'growth', 'starter')
            """
        )
    )
    plans = {str(row.code or "").strip().lower(): int(row.id) for row in result.fetchall()}
    for code in TRIAL_PLAN_PRIORITY:
        if code in plans:
            return plans[code]
    raise HTTPException(status_code=404, detail="Plano de demonstracao nao encontrado.")


def decode_features(raw: Any) -> list[str]:
    if isinstance(raw, dict):
        return [
            FEATURE_PUBLIC_LABELS.get(str(key), str(key).replace("_", " ").title())
            for key, enabled in raw.items()
            if bool(enabled) and str(key) in FEATURE_PUBLIC_LABELS
        ]
    elif isinstance(raw, list):
        values = raw
    else:
        try:
            parsed = json.loads(str(raw or "[]"))
        except Exception:
            parsed = []
        if isinstance(parsed, dict):
            return decode_features(parsed)
        elif isinstance(parsed, list):
            values = parsed
        else:
            values = []
    return [str(item).strip() for item in values if str(item).strip()]


async def ensure_public_signup_table(db: AsyncSession) -> None:
    await db.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS public_signup_leads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                document TEXT NOT NULL,
                email TEXT NOT NULL,
                phone TEXT NOT NULL,
                company TEXT NOT NULL,
                plan_code TEXT NOT NULL,
                message TEXT,
                status TEXT DEFAULT 'novo',
                user_agent TEXT,
                ip_address TEXT,
                company_id INTEGER,
                reviewed_at TEXT,
                reviewed_by TEXT,
                review_notes TEXT,
                trial_days INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
    result = await db.execute(text("PRAGMA table_info(public_signup_leads)"))
    columns = {str(row[1]).lower() for row in result.fetchall()}
    for column, definition in {
        "company_id": "INTEGER",
        "reviewed_at": "TEXT",
        "reviewed_by": "TEXT",
        "review_notes": "TEXT",
        "trial_days": "INTEGER",
    }.items():
        if column not in columns:
            await db.execute(text(f"ALTER TABLE public_signup_leads ADD COLUMN {column} {definition}"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_public_signup_leads_document ON public_signup_leads(document)"))
    await db.execute(text("CREATE INDEX IF NOT EXISTS idx_public_signup_leads_created ON public_signup_leads(created_at)"))


@router.get("/plans")
async def public_plans(db: AsyncSession = Depends(get_db)):
    try:
        result = await db.execute(
            text(
                """
                SELECT code, name, monthly_price, description, features_json
                  FROM plans
                 WHERE COALESCE(status, 'active')='active'
                   AND code IN ('starter', 'growth', 'professional', 'enterprise')
                 ORDER BY CASE code
                    WHEN 'starter' THEN 1
                    WHEN 'growth' THEN 2
                    WHEN 'professional' THEN 3
                    WHEN 'enterprise' THEN 4
                    ELSE 99
                 END
                 LIMIT 6
                """
            )
        )
        plans = []
        for row in result.mappings().all():
            item = dict(row)
            code = str(item.get("code") or "").strip()
            public_copy = PLAN_PUBLIC_COPY.get(code, {})
            features = decode_features(item.get("features_json"))
            if not features:
                features = list(public_copy.get("features") or [])
            plans.append(
                {
                    "code": code,
                    "name": str(public_copy.get("name") or item.get("name") or item.get("code") or "").strip(),
                    "monthly_price": float(item.get("monthly_price") or 0),
                    "description": str(public_copy.get("description") or item.get("description") or "").strip(),
                    "summary": str(public_copy.get("summary") or "").strip(),
                    "vehicle_limit": public_copy.get("vehicle_limit"),
                    "user_limit": public_copy.get("user_limit"),
                    "price_label": "Sob consulta" if code == "enterprise" else "",
                    "features": features,
                }
            )
        if plans:
            return {"plans": plans}
    except Exception:
        pass
    return {"plans": PLAN_FALLBACKS}


@router.post("/signup", status_code=status.HTTP_201_CREATED)
async def public_signup(
    payload: PublicSignupPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
):
    await ensure_public_signup_table(db)
    ip_address = enforce_public_signup_rate_limit(request)
    try:
        result = await db.execute(
            text("SELECT 1 FROM companies WHERE document=:document LIMIT 1"),
            {"document": payload.document},
        )
        if result.scalar_one_or_none():
            raise HTTPException(status_code=409, detail="Ja existe uma empresa cadastrada com este CPF/CNPJ.")

        result = await db.execute(
            text("SELECT 1 FROM usuarios WHERE UPPER(TRIM(username))=UPPER(:username) LIMIT 1"),
            {"username": payload.username},
        )
        if result.scalar_one_or_none():
            raise HTTPException(status_code=409, detail="Este login ja esta em uso. Escolha outro.")

        plan_id = await best_trial_plan_id(db)

        await db.execute(
            text(
                """
                INSERT INTO public_signup_leads (
                    name, document, email, phone, company, plan_code, message,
                    status, user_agent, ip_address, reviewed_at, reviewed_by, trial_days
                ) VALUES (
                    :name, :document, :email, :phone, :company, :plan_code, :message,
                    'aprovado', :user_agent, :ip_address, datetime('now'), 'cadastro_publico', 30
                )
                """
            ),
            {
                "name": payload.name,
                "document": payload.document,
                "email": str(payload.email),
                "phone": payload.phone,
                "company": payload.company,
                "plan_code": payload.plan_code,
                "message": payload.message or "",
                "user_agent": request.headers.get("user-agent", "")[:300],
                "ip_address": ip_address,
            },
        )
        lead_id = (await db.execute(text("SELECT last_insert_rowid()"))).scalar_one()
        result = await db.execute(
            text(
                """
                INSERT INTO companies (
                    code, name, legal_name, document, email, phone, status, timezone, created_at, updated_at
                )
                SELECT
                    :code, :company, :company, :document, :email, :phone,
                    'active', 'America/Fortaleza', datetime('now'), datetime('now')
                WHERE NOT EXISTS (
                    SELECT 1 FROM companies WHERE document=:document
                )
                """
            ),
            {
                "code": company_code(payload.company, lead_id),
                "company": payload.company,
                "document": payload.document,
                "email": str(payload.email),
                "phone": payload.phone,
            },
        )
        if result.rowcount != 1:
            raise HTTPException(status_code=409, detail="Ja existe uma empresa cadastrada com este CPF/CNPJ.")
        company_id = (await db.execute(text("SELECT last_insert_rowid()"))).scalar_one()
        await db.execute(
            text(
                """
                INSERT INTO subscriptions (
                    company_id, plan_id, status, billing_cycle,
                    current_period_start, current_period_end, next_due_date,
                    created_at, updated_at
                ) VALUES (
                    :company_id, :plan_id, 'trialing', 'monthly',
                    date('now'), date('now', '+30 day'), date('now', '+30 day'),
                    datetime('now'), datetime('now')
                )
                """
            ),
            {"company_id": company_id, "plan_id": plan_id},
        )
        subscription_id = (await db.execute(text("SELECT last_insert_rowid()"))).scalar_one()
        result = await db.execute(
            text(
                """
                INSERT INTO usuarios (
                    username, nome, senha, permissoes, cpf, telefone, is_active, company_id
                )
                SELECT
                    :username, :name, :password_hash, 'ADMIN', :document, :phone, 1, :company_id
                WHERE NOT EXISTS (
                    SELECT 1 FROM usuarios WHERE UPPER(TRIM(username))=UPPER(:username)
                )
                """
            ),
            {
                "username": payload.username,
                "name": payload.name,
                "password_hash": get_password_hash(payload.password),
                "document": payload.document,
                "phone": payload.phone,
                "company_id": company_id,
            },
        )
        if result.rowcount != 1:
            raise HTTPException(status_code=409, detail="Este login ja esta em uso. Escolha outro.")
        await db.execute(
            text("UPDATE public_signup_leads SET company_id=:company_id WHERE id=:lead_id"),
            {"company_id": company_id, "lead_id": lead_id},
        )
        await db.execute(
            text(
                """
                INSERT INTO audit_logs (
                    company_id, actor_type, action, entity_type, entity_id, severity,
                    ip_address, metadata_json, created_at
                ) VALUES (
                    :company_id, 'customer', 'demonstracao_autoativada', 'subscription',
                    :subscription_id, 'info', :ip_address, :metadata_json, datetime('now')
                )
                """
            ),
            {
                "company_id": company_id,
                "subscription_id": str(subscription_id),
                "ip_address": ip_address,
                "metadata_json": json.dumps(
                    {"lead_id": int(lead_id), "plan_code": payload.plan_code, "trial_days": 30},
                    ensure_ascii=True,
                    sort_keys=True,
                ),
            },
        )
        await db.commit()
    except HTTPException:
        await db.rollback()
        raise
    except Exception as exc:
        await db.rollback()
        raise HTTPException(
            status_code=409,
            detail="Nao foi possivel concluir o cadastro. Confira os dados e tente novamente.",
        ) from exc
    return {
        "ok": True,
        "message": "Cadastro concluido. Seu acesso ao RotaHub ja esta liberado.",
        "next_url": "/app/index.html",
    }
