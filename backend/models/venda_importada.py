# backend/models/venda_importada.py
"""
Imported sales model used by the Importar Vendas browser flow.

The columns follow the desktop table ``vendas_importadas`` so sales can move
from import, to selection, to a linked programacao without changing vocabulary.
"""
from sqlalchemy import Column, Float, Integer, String

from backend.config.database import Base


class VendaImportadaDB(Base):
    __tablename__ = "vendas_importadas"

    id = Column(Integer, primary_key=True, index=True)
    pedido = Column(String, index=True)
    data_venda = Column(String)
    cliente = Column(String, index=True)
    nome_cliente = Column(String)
    vendedor = Column(String)
    produto_id = Column(Integer, index=True)
    produto = Column(String, index=True)
    vr_total = Column(Float, default=0)
    qnt = Column(Float, default=0)
    qnt_caixas = Column(Integer, default=0)
    cidade = Column(String)
    valor_unitario = Column(Float, default=0)
    observacao = Column(String)
    selecionada = Column(Integer, default=0)
    usada = Column(Integer, default=0)
    usada_em = Column(String)
    codigo_programacao = Column(String, index=True)
    company_id = Column(Integer)
