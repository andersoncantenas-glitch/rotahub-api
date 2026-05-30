# backend/api/v1/endpoints/public.py
"""
Public marketing and signup endpoints.

These endpoints intentionally do not authenticate users into the operational
system. They collect commercial interest and expose plan cards for the public
landing page.
"""
from __future__ import annotations

import json
import re
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Request, status
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from backend.config.database import get_db

router = APIRouter()


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
    message: str | None = Field(default="", max_length=600)

    @field_validator("name", "document", "email", "phone", "company", "plan_code", "message", mode="before")
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
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )
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
    result = await db.execute(
        text("SELECT 1 FROM public_signup_leads WHERE document=:document AND status='novo' LIMIT 1"),
        {"document": payload.document},
    )
    duplicate = result.scalar_one_or_none()
    if duplicate:
        raise HTTPException(status_code=409, detail="Ja existe um cadastro em analise para este CPF/CNPJ.")
    await db.execute(
        text(
            """
            INSERT INTO public_signup_leads (
                name, document, email, phone, company, plan_code, message, user_agent, ip_address
            ) VALUES (
                :name, :document, :email, :phone, :company, :plan_code, :message, :user_agent, :ip_address
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
            "ip_address": str(request.client.host if request.client else ""),
        },
    )
    await db.commit()
    return {
        "ok": True,
        "message": "Cadastro recebido. A equipe RotaHub irá validar os dados e liberar o acesso.",
        "next_url": "/app/index.html",
    }
