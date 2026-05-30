# backend/api/v1/endpoints/auth.py
"""
Authentication endpoints
"""
from datetime import timedelta
import json
from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from pydantic import BaseModel

from backend.config.database import get_db
from backend.config.settings import settings
from backend.models.user import User
from backend.services.auth import authenticate_user, create_access_token, get_current_user

router = APIRouter()

# OAuth2 scheme
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/v1/auth/login")


class Token(BaseModel):
    access_token: str
    token_type: str
    expires_in: int


class TokenData(BaseModel):
    username: str | None = None


class PlanContext(BaseModel):
    company_id: int
    plan_code: str | None = None
    plan_name: str | None = None
    subscription_status: str | None = None
    next_due_date: str | None = None
    vehicle_limit: int | None = None
    user_limit: int | None = None
    features: dict[str, bool] = {}


def issue_user_token(user: User) -> Token:
    access_token_expires = timedelta(minutes=settings.JWT_ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": user.username, "user_id": user.id, "company_id": user.company_id or 1},
        expires_delta=access_token_expires,
    )
    return Token(
        access_token=access_token,
        token_type="bearer",
        expires_in=int(access_token_expires.total_seconds()),
    )


@router.post("/login", response_model=Token)
async def login_for_access_token(
    form_data: OAuth2PasswordRequestForm = Depends(),
    db: AsyncSession = Depends(get_db)
):
    """Authenticate user and return access token"""
    user = await authenticate_user(db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Bearer"},
        )

    return issue_user_token(user)


@router.post("/refresh", response_model=Token)
async def refresh_access_token(current_user: User = Depends(get_current_user)):
    """Issue a fresh access token for a valid authenticated session."""
    return issue_user_token(current_user)


@router.get("/plan-context", response_model=PlanContext)
async def get_plan_context(
    current_user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    company_id = int(current_user.company_id or 1)
    result = await db.execute(
        text(
            """
            SELECT s.status AS subscription_status, s.next_due_date,
                   p.code AS plan_code, p.name AS plan_name,
                   p.vehicle_limit, p.user_limit, p.features_json
              FROM subscriptions s
              JOIN plans p ON p.id=s.plan_id
             WHERE s.company_id=:company_id
               AND s.status IN ('active', 'trialing', 'past_due')
             ORDER BY s.id DESC
             LIMIT 1
            """
        ),
        {"company_id": company_id},
    )
    row = result.mappings().first()
    if not row:
        return PlanContext(company_id=company_id)
    try:
        features = json.loads(str(row.get("features_json") or "{}"))
    except Exception:
        features = {}
    return PlanContext(
        company_id=company_id,
        plan_code=row.get("plan_code"),
        plan_name=row.get("plan_name"),
        subscription_status=row.get("subscription_status"),
        next_due_date=row.get("next_due_date"),
        vehicle_limit=row.get("vehicle_limit"),
        user_limit=row.get("user_limit"),
        features=features if isinstance(features, dict) else {},
    )


@router.post("/logout")
async def logout(current_user: User = Depends(get_current_user)):
    """Confirm logout for a valid authenticated session.

    JWTs are stateless in this deployment; the browser completes logout by
    removing the stored token.
    """
    return {"message": "Logged out successfully", "username": current_user.username}
