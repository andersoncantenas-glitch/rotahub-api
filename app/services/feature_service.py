from __future__ import annotations

import json

from app.repositories import subscription_repository
from app.services.saas_result import error_message, service_result


def list_company_features(company_id: int) -> dict:
    try:
        subscription = subscription_repository.get_active_subscription(company_id)
        if not subscription:
            return service_result(ok=False, data=None, error="Assinatura ativa nao encontrada.")
        features = _decode_features(subscription.get("plan_features_json"))
        return service_result(
            ok=True,
            data={
                "company_id": int(company_id),
                "plan_code": subscription.get("plan_code"),
                "plan_name": subscription.get("plan_name"),
                "features": features,
            },
        )
    except Exception as exc:
        return service_result(ok=False, data=None, error=error_message(exc, "Falha ao consultar recursos do plano."))


def can_use_feature(company_id: int, feature_name: str) -> dict:
    result = list_company_features(company_id)
    if not bool(result.get("ok", False)):
        return result
    data = result.get("data") or {}
    features = data.get("features") if isinstance(data.get("features"), dict) else {}
    feature_key = str(feature_name or "").strip()
    allowed = bool(features.get(feature_key))
    out = dict(data)
    out.update({"feature": feature_key, "allowed": allowed})
    return service_result(ok=True, data=out)


def _decode_features(raw) -> dict:
    if isinstance(raw, dict):
        return dict(raw)
    try:
        parsed = json.loads(str(raw or "{}"))
    except Exception:
        parsed = {}
    return parsed if isinstance(parsed, dict) else {}
