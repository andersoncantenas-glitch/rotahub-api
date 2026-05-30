# backend/models/rota.py
"""
Operational route tracking models.
"""
from sqlalchemy import Column, Float, Integer, String

from backend.config.database import Base


class RotaGpsPingDB(Base):
    __tablename__ = "rota_gps_pings"

    id = Column(Integer, primary_key=True, index=True)
    codigo_programacao = Column(String, index=True, nullable=False)
    motorista = Column(String)
    lat = Column(Float)
    lon = Column(Float)
    speed = Column(Float)
    accuracy = Column(Float)
    recorded_at = Column(String)
    company_id = Column(Integer)
