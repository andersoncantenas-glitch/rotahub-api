# backend/models/despesa.py
"""
Despesas linked to programacoes.
"""
from sqlalchemy import Column, Float, Integer, String

from backend.config.database import Base


class DespesaDB(Base):
    __tablename__ = "despesas"

    id = Column(Integer, primary_key=True, index=True)
    codigo_programacao = Column(String, index=True, nullable=False)
    descricao = Column(String)
    valor = Column(Float, default=0)
    data_registro = Column(String)
    tipo_despesa = Column(String, default="ROTA")
    categoria = Column(String)
    motorista = Column(String)
    veiculo = Column(String)
    observacao = Column(String)
    id_local = Column(String)
    forma_pagamento = Column(String)
    comprovante_path = Column(String)
    estabelecimento = Column(String)
    documento = Column(String)
    litros = Column(Float)
    valor_litro = Column(Float)
    desconto = Column(Float)
    combustivel = Column(String)
    odometro = Column(Float)
    lat = Column(Float)
    lon = Column(Float)
    accuracy = Column(Float)
    registrado_em = Column(String)
    motorista_codigo = Column(String)
    motorista_nome = Column(String)
    sync_key = Column(String)
    status_sync = Column(String)
    origem = Column(String)
    vinculo_prestacao_json = Column(String)
    desktop_web_json = Column(String)
    foto_despesa_ref_json = Column(String)
    company_id = Column(Integer)
