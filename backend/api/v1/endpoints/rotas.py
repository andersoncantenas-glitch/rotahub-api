# backend/api/v1/endpoints/rotas.py
"""
Route monitoring endpoints mirroring the desktop RotasPage core flow.
"""
from __future__ import annotations

from datetime import datetime
import math

from fastapi import APIRouter, Depends, HTTPException, Query, Request, status
from pydantic import BaseModel, Field
from sqlalchemy import func, or_, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.utils.formatters import safe_float
from backend.api.v1.endpoints.programacao import get_programacao_by_codigo, upper_text
from backend.api.v1.endpoints.users import require_admin_user
from backend.config.database import get_db
from backend.models.programacao import ProgramacaoDB
from backend.models.rota import RotaGpsPingDB
from backend.models.user import User
from backend.services.audit import client_ip_from_request, record_audit_log

router = APIRouter()

BLOCKED_STATUSES = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}


class RotaMonitoramentoResponse(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    veiculo: str = ""
    status: str = ""
    lat: float | None = None
    lon: float | None = None
    speed: float | None = None
    accuracy: float | None = None
    recorded_at: str = ""


class RotaGpsPayload(BaseModel):
    lat: float = Field(ge=-90, le=90)
    lon: float = Field(ge=-180, le=180)
    speed: float | None = Field(default=None, ge=0)
    accuracy: float | None = Field(default=None, ge=0)
    recorded_at: str | None = Field(default=None, max_length=40)


class RotaGpsResponse(BaseModel):
    id: int
    codigo_programacao: str
    lat: float
    lon: float
    speed: float | None = None
    accuracy: float | None = None
    recorded_at: str


class RotaGpsPoint(BaseModel):
    lat: float
    lon: float
    speed: float | None = None
    accuracy: float | None = None
    recorded_at: str = ""


class RotaParada(BaseModel):
    lat: float
    lon: float
    inicio: str = ""
    fim: str = ""
    minutos: float = 0
    pontos: int = 0


class RotaRastreamentoResponse(BaseModel):
    codigo_programacao: str
    motorista: str = ""
    veiculo: str = ""
    status: str = ""
    pontos: list[RotaGpsPoint] = Field(default_factory=list)
    paradas: list[RotaParada] = Field(default_factory=list)
    total_pontos: int = 0
    inicio: str = ""
    fim: str = ""
    ultima_atualizacao: str = ""
    km_estimado: float = 0
    velocidade_media: float = 0
    velocidade_maxima: float = 0
    tempo_rastreado_min: float = 0


def parse_recorded_at(value: str | None) -> datetime | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"):
        try:
            return datetime.strptime(raw[:26], fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None


def distance_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return radius * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))


def route_metrics(points: list[RotaGpsPoint]) -> tuple[float, float, float, float, list[RotaParada]]:
    if not points:
        return 0.0, 0.0, 0.0, 0.0, []
    total_km = 0.0
    speeds = []
    dated: list[tuple[RotaGpsPoint, datetime | None]] = [(p, parse_recorded_at(p.recorded_at)) for p in points]
    for index, (point, _) in enumerate(dated):
        if point.speed is not None:
            speeds.append(safe_float(point.speed, 0.0))
        if index > 0:
            prev = dated[index - 1][0]
            segment = distance_km(prev.lat, prev.lon, point.lat, point.lon)
            if segment < 5:
                total_km += segment
    valid_dates = [dt for _, dt in dated if dt is not None]
    minutes = 0.0
    if len(valid_dates) >= 2:
        minutes = max((max(valid_dates) - min(valid_dates)).total_seconds() / 60, 0.0)
    avg_speed = (total_km / (minutes / 60)) if minutes > 0 and total_km > 0 else (sum(speeds) / len(speeds) if speeds else 0.0)
    max_speed = max(speeds) if speeds else 0.0
    stops: list[RotaParada] = []
    cluster: list[tuple[RotaGpsPoint, datetime | None]] = []

    def flush_cluster() -> None:
        nonlocal cluster
        if len(cluster) < 2:
            cluster = []
            return
        dates = [dt for _, dt in cluster if dt is not None]
        if len(dates) < 2:
            cluster = []
            return
        duration = (max(dates) - min(dates)).total_seconds() / 60
        if duration >= 3:
            stops.append(
                RotaParada(
                    lat=round(sum(p.lat for p, _ in cluster) / len(cluster), 6),
                    lon=round(sum(p.lon for p, _ in cluster) / len(cluster), 6),
                    inicio=min(dates).strftime("%Y-%m-%d %H:%M:%S"),
                    fim=max(dates).strftime("%Y-%m-%d %H:%M:%S"),
                    minutos=round(duration, 1),
                    pontos=len(cluster),
                )
            )
        cluster = []

    for point, dt in dated:
        stopped = safe_float(point.speed, 0.0) <= 3
        if not stopped:
            flush_cluster()
            continue
        if cluster:
            prev = cluster[-1][0]
            if distance_km(prev.lat, prev.lon, point.lat, point.lon) > 0.08:
                flush_cluster()
        cluster.append((point, dt))
    flush_cluster()
    return round(total_km, 2), round(avg_speed, 2), round(max_speed, 2), round(minutes, 1), stops


def route_status(programacao: ProgramacaoDB) -> str:
    return upper_text(programacao.status_operacional or programacao.status)


def is_active_route(programacao: ProgramacaoDB) -> bool:
    status_value = upper_text(programacao.status)
    operational = upper_text(programacao.status_operacional)
    if status_value in BLOCKED_STATUSES or operational in BLOCKED_STATUSES:
        return False
    if safe_float(programacao.finalizada_no_app, 0) == 1:
        return False
    if upper_text(programacao.data_chegada) or upper_text(programacao.hora_chegada):
        return False
    if safe_float(programacao.km_final, 0.0) > 0:
        return False
    return True


def gps_response(ping: RotaGpsPingDB) -> RotaGpsResponse:
    return RotaGpsResponse(
        id=ping.id,
        codigo_programacao=upper_text(ping.codigo_programacao),
        lat=safe_float(ping.lat, 0.0),
        lon=safe_float(ping.lon, 0.0),
        speed=None if ping.speed is None else safe_float(ping.speed, 0.0),
        accuracy=None if ping.accuracy is None else safe_float(ping.accuracy, 0.0),
        recorded_at=ping.recorded_at or "",
    )


@router.get("/monitoramento", response_model=list[RotaMonitoramentoResponse])
async def monitoramento_rotas(
    limit: int = 300,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    latest = (
        select(
            RotaGpsPingDB.codigo_programacao.label("codigo_programacao"),
            func.max(RotaGpsPingDB.id).label("max_id"),
        )
        .group_by(RotaGpsPingDB.codigo_programacao)
        .subquery()
    )
    stmt = (
        select(ProgramacaoDB, RotaGpsPingDB)
        .outerjoin(latest, latest.c.codigo_programacao == ProgramacaoDB.codigo_programacao)
        .outerjoin(RotaGpsPingDB, RotaGpsPingDB.id == latest.c.max_id)
        .where(
            or_(ProgramacaoDB.status.is_(None), func.upper(func.coalesce(ProgramacaoDB.status, "")).not_in(BLOCKED_STATUSES)),
            or_(
                ProgramacaoDB.status_operacional.is_(None),
                func.upper(func.coalesce(ProgramacaoDB.status_operacional, "")).not_in(BLOCKED_STATUSES),
            ),
            func.coalesce(ProgramacaoDB.finalizada_no_app, 0) == 0,
            func.trim(func.coalesce(ProgramacaoDB.data_chegada, "")) == "",
            func.trim(func.coalesce(ProgramacaoDB.hora_chegada, "")) == "",
            func.coalesce(ProgramacaoDB.km_final, 0) == 0,
        )
        .order_by(ProgramacaoDB.id.desc())
        .limit(max(min(limit, 500), 1))
    )
    result = await db.execute(stmt)
    out = []
    for programacao, ping in result.all():
        if not is_active_route(programacao):
            continue
        out.append(
            RotaMonitoramentoResponse(
                codigo_programacao=upper_text(programacao.codigo_programacao),
                motorista=upper_text(programacao.motorista),
                veiculo=upper_text(programacao.veiculo),
                status=route_status(programacao),
                lat=None if ping is None or ping.lat is None else safe_float(ping.lat, 0.0),
                lon=None if ping is None or ping.lon is None else safe_float(ping.lon, 0.0),
                speed=None if ping is None or ping.speed is None else safe_float(ping.speed, 0.0),
                accuracy=None if ping is None or ping.accuracy is None else safe_float(ping.accuracy, 0.0),
                recorded_at="" if ping is None else (ping.recorded_at or ""),
            )
        )
    return out


@router.get("/{codigo_programacao}/rastreamento", response_model=RotaRastreamentoResponse)
async def rastreamento_rota(
    codigo_programacao: str,
    limit: int = Query(1000, ge=1, le=5000),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    del current_user
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    codigo = upper_text(programacao.codigo_programacao)
    result = await db.execute(
        select(RotaGpsPingDB)
        .where(func.upper(func.coalesce(RotaGpsPingDB.codigo_programacao, "")) == codigo)
        .order_by(RotaGpsPingDB.id.desc())
        .limit(limit)
    )
    rows = list(reversed(result.scalars().all()))
    points = [
        RotaGpsPoint(
            lat=safe_float(row.lat, 0.0),
            lon=safe_float(row.lon, 0.0),
            speed=None if row.speed is None else safe_float(row.speed, 0.0),
            accuracy=None if row.accuracy is None else safe_float(row.accuracy, 0.0),
            recorded_at=row.recorded_at or "",
        )
        for row in rows
        if row.lat is not None and row.lon is not None
    ]
    km, avg_speed, max_speed, minutes, stops = route_metrics(points)
    return RotaRastreamentoResponse(
        codigo_programacao=codigo,
        motorista=upper_text(programacao.motorista),
        veiculo=upper_text(programacao.veiculo),
        status=route_status(programacao),
        pontos=points,
        paradas=stops,
        total_pontos=len(points),
        inicio=points[0].recorded_at if points else "",
        fim=points[-1].recorded_at if points else "",
        ultima_atualizacao=points[-1].recorded_at if points else "",
        km_estimado=km,
        velocidade_media=avg_speed,
        velocidade_maxima=max_speed,
        tempo_rastreado_min=minutes,
    )


@router.post("/{codigo_programacao}/gps", response_model=RotaGpsResponse, status_code=status.HTTP_201_CREATED)
async def registrar_gps_rota(
    codigo_programacao: str,
    payload: RotaGpsPayload,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(require_admin_user),
):
    programacao = await get_programacao_by_codigo(db, codigo_programacao)
    if not programacao:
        raise HTTPException(status_code=404, detail="Programacao nao encontrada")
    if not is_active_route(programacao):
        raise HTTPException(status_code=409, detail="Programacao nao esta ativa para rastreamento.")

    ping = RotaGpsPingDB(
        codigo_programacao=upper_text(programacao.codigo_programacao),
        motorista=upper_text(programacao.motorista) or upper_text(current_user.nome or current_user.username),
        lat=safe_float(payload.lat, 0.0),
        lon=safe_float(payload.lon, 0.0),
        speed=None if payload.speed is None else safe_float(payload.speed, 0.0),
        accuracy=None if payload.accuracy is None else safe_float(payload.accuracy, 0.0),
        recorded_at=payload.recorded_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    db.add(ping)
    record_audit_log(
        db,
        action="rota_gps_registrado",
        actor_user=current_user,
        entity_type="programacao",
        entity_id=upper_text(programacao.codigo_programacao),
        ip_address=client_ip_from_request(request),
        metadata={"lat": ping.lat, "lon": ping.lon, "speed": ping.speed, "accuracy": ping.accuracy},
    )
    await db.commit()
    await db.refresh(ping)
    return gps_response(ping)
