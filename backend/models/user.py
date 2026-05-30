# backend/models/user.py
"""
User model
"""
from typing import Optional
from sqlalchemy import Boolean, Column, Integer, String
from pydantic import BaseModel as PydanticBaseModel, ConfigDict

from backend.config.database import Base


class UserDB(Base):
    """SQLAlchemy User model"""
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True, nullable=False)
    nome = Column(String, nullable=False)
    senha = Column(String, nullable=False)
    permissoes = Column(String, default="OPERADOR")
    is_active = Column(Boolean, nullable=False, default=True)
    cpf = Column(String)
    idade = Column(Integer)
    telefone = Column(String)
    company_id = Column(Integer)


class User(PydanticBaseModel):
    """Pydantic User model"""
    model_config = ConfigDict(from_attributes=True)

    id: int
    username: str
    nome: str
    senha: str
    permissoes: str = "OPERADOR"
    is_active: bool = True
    cpf: Optional[str] = None
    idade: Optional[int] = None
    telefone: Optional[str] = None
    company_id: Optional[int] = None

