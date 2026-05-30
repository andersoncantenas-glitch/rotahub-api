# backend/api/v1/api.py
"""
Main API router for v1 endpoints
"""
from fastapi import APIRouter

from backend.api.v1.endpoints import (
    audit,
    auth,
    billing,
    cadastros,
    centro_custos,
    compras,
    despesas,
    escala,
    health,
    home,
    importar_vendas,
    logistica,
    permissoes,
    programacao,
    public,
    recebimentos,
    relatorios,
    rotas,
    saas_admin,
    system_tools,
    users,
)

# Create main API router
api_router = APIRouter()

# Include endpoint routers
api_router.include_router(auth.router, prefix="/auth", tags=["Authentication"])
api_router.include_router(billing.router, prefix="/billing", tags=["Billing"])
api_router.include_router(public.router, prefix="/public", tags=["Public"])
api_router.include_router(health.router, prefix="/health", tags=["Health"])
api_router.include_router(home.router, prefix="/home", tags=["Home"])
api_router.include_router(users.router, prefix="/users", tags=["Users"])
api_router.include_router(audit.router, prefix="/audit-logs", tags=["Audit"])
api_router.include_router(cadastros.router, prefix="/cadastros", tags=["Cadastros"])
api_router.include_router(programacao.router, prefix="/programacao", tags=["Planejamento de Rota"])
api_router.include_router(importar_vendas.router, prefix="/importar-vendas", tags=["Importar Pedidos"])
api_router.include_router(logistica.router, prefix="/logistica", tags=["Logistica"])
api_router.include_router(rotas.router, prefix="/rotas", tags=["Rotas"])
api_router.include_router(escala.router, prefix="/escala", tags=["Escala"])
api_router.include_router(recebimentos.router, prefix="/recebimentos", tags=["Recebimentos"])
api_router.include_router(despesas.router, prefix="/despesas", tags=["Custos e Despesas"])
api_router.include_router(centro_custos.router, prefix="/centro-custos", tags=["Analise de Custos"])
api_router.include_router(compras.router, prefix="/compras", tags=["Compras"])
api_router.include_router(relatorios.router, prefix="/relatorios", tags=["Relatorios"])
api_router.include_router(system_tools.router, prefix="/system-tools", tags=["Ferramentas do Sistema"])
api_router.include_router(permissoes.router, prefix="/permissoes", tags=["Permissoes"])
api_router.include_router(saas_admin.router, prefix="/saas-admin", tags=["Admin SaaS"])

# Future endpoints to be added:
# api_router.include_router(programacao.router, prefix="/programacao", tags=["Programação"])
# api_router.include_router(rotas.router, prefix="/rotas", tags=["Rotas"])
# api_router.include_router(escala.router, prefix="/escala", tags=["Escala"])
# api_router.include_router(financeiro.router, prefix="/financeiro", tags=["Financeiro"])
# api_router.include_router(relatorios.router, prefix="/relatorios", tags=["Relatórios"])
