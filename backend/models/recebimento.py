# backend/models/recebimento.py
"""
Recebimentos linked to programacoes.
"""
from sqlalchemy import Column, Float, Integer, String

from backend.config.database import Base


class RecebimentoDB(Base):
    __tablename__ = "recebimentos"

    id = Column(Integer, primary_key=True, index=True)
    codigo_programacao = Column(String, index=True, nullable=False)
    cod_cliente = Column(String, index=True, nullable=False)
    pedido = Column(String)
    nome_cliente = Column(String, nullable=False)
    valor = Column(Float, default=0)
    forma_pagamento = Column(String)
    observacao = Column(String)
    num_nf = Column(String)
    data_registro = Column(String)
    company_id = Column(Integer)
