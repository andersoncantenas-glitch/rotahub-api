from __future__ import annotations

import json

from app.repositories import plan_repository
from app.services.saas_result import error_message, service_result


def _decode_features(plan: dict | None) -> dict | None:
    if not plan:
        return plan
    out = dict(plan)
    raw = str(out.get("features_json") or "{}")
    try:
        out["features"] = json.loads(raw)
    except Exception:
        out["features"] = {}
    return out


def list_plans(*, include_inactive: bool = False, limit: int | None = 100) -> dict:
    try:
        plans = [_decode_features(plan) for plan in plan_repository.list_plans(include_inactive=include_inactive, limit=limit)]
        return service_result(ok=True, data=plans)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao listar planos."))


def get_plan(plan_id: int) -> dict:
    try:
        plan = _decode_features(plan_repository.get_plan(plan_id))
        if not plan:
            return service_result(ok=False, data=None, error="Plano nao encontrado.")
        return service_result(ok=True, data=plan)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao buscar plano."))


def get_plan_by_code(code: str) -> dict:
    try:
        plan = _decode_features(plan_repository.get_plan_by_code(code))
        if not plan:
            return service_result(ok=False, data=None, error="Plano nao encontrado.")
        return service_result(ok=True, data=plan)
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao buscar plano."))


def can_use_feature(plan: dict, feature_name: str) -> bool:
    features = plan.get("features") if isinstance(plan, dict) else None
    if not isinstance(features, dict):
        try:
            features = json.loads(str((plan or {}).get("features_json") or "{}"))
        except Exception:
            features = {}
    return bool(features.get(str(feature_name or "").strip()))
