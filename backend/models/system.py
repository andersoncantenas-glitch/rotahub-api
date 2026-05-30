# backend/models/system.py
"""
System maintenance logs used by Backup / Exportar and Ferramentas.
"""
from datetime import datetime

from sqlalchemy import Column, Integer, String, Text

from backend.config.database import Base


def local_now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class SistemaLogDB(Base):
    __tablename__ = "sistema_logs"

    id = Column(Integer, primary_key=True, index=True)
    tipo_acao = Column(String)
    descricao = Column(String)
    usuario = Column(String)
    status = Column(String, default="OK")
    resultado_texto = Column(Text)
    executado_em = Column(String, default=local_now_text)
