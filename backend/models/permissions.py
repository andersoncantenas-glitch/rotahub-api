# backend/models/permissions.py
"""
Fine-grained permission models used by the web permissions manager.
"""
from sqlalchemy import Column, ForeignKey, Integer, String, UniqueConstraint, text

from backend.config.database import Base


class PermissaoDB(Base):
    """Permission catalog entry."""

    __tablename__ = "permissoes"

    id = Column(Integer, primary_key=True, index=True)
    modulo = Column(String, nullable=False)
    nome_permissao = Column(String, nullable=False)
    descricao = Column(String)
    ativo = Column(Integer, default=1)


class UsuarioPermissaoDB(Base):
    """Permission granted to a user."""

    __tablename__ = "usuario_permissoes"
    __table_args__ = (UniqueConstraint("usuario_id", "permissao_id", name="uq_usuario_permissao"),)

    id = Column(Integer, primary_key=True, index=True)
    usuario_id = Column(Integer, ForeignKey("usuarios.id"), nullable=False)
    permissao_id = Column(Integer, ForeignKey("permissoes.id"), nullable=False)
    concedida_em = Column(String, server_default=text("(datetime('now'))"))
    concedida_por = Column(String)
