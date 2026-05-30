# backend/services/auth.py
"""
Authentication service
"""
import logging
from datetime import datetime, timedelta, timezone
from typing import Optional
from jose import JWTError, jwt
from passlib.context import CryptContext
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.config.database import get_db
from backend.config.settings import settings
from backend.models.user import User, UserDB
from app.security.passwords import verify_password_pbkdf2

logger = logging.getLogger(__name__)

# Password hashing
pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

# OAuth2 scheme
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/v1/auth/login")


def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Verify a password against its hash"""
    stored = str(hashed_password or "")
    if stored.startswith("pbkdf2_sha256$"):
        return verify_password_pbkdf2(plain_password, stored)
    if not stored.startswith("$"):
        return stored == str(plain_password or "")
    try:
        return pwd_context.verify(plain_password, stored)
    except Exception:
        logger.debug("Password verification failed", exc_info=True)
        return False


def get_password_hash(password: str) -> str:
    """Hash a password"""
    return pwd_context.hash(password)


def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    """Create JWT access token"""
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.now(timezone.utc) + expires_delta
    else:
        expire = datetime.now(timezone.utc) + timedelta(minutes=15)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, settings.JWT_SECRET_KEY, algorithm=settings.JWT_ALGORITHM)
    return encoded_jwt


async def authenticate_user(db: AsyncSession, username: str, password: str) -> Optional[User]:
    """Authenticate user with username and password"""
    username_norm = str(username or "").strip().upper()
    stmt = select(UserDB).where(
        (UserDB.username.is_not(None) & (UserDB.username != "") & (UserDB.username.ilike(username_norm)))
        | (UserDB.nome.is_not(None) & (UserDB.nome != "") & (UserDB.nome.ilike(username_norm)))
    ).order_by(UserDB.id.asc()).limit(1)
    result = await db.execute(stmt)
    user_db = result.scalar_one_or_none()

    if not user_db or not user_db.is_active or not verify_password(password, user_db.senha):
        return None

    if not user_db.username:
        user_db.username = str(user_db.nome or username_norm or "ADMIN").strip().upper()
        await db.flush()

    company_id = int(user_db.company_id or 1)
    db.info["company_id"] = company_id
    db.sync_session.info["company_id"] = company_id
    return User.model_validate(user_db)


async def get_user_by_id(db: AsyncSession, user_id: int) -> Optional[User]:
    """Get user by ID from the database"""
    user_db = await db.get(UserDB, user_id)
    return User.model_validate(user_db) if user_db else None


async def get_current_user(token: str = Depends(oauth2_scheme), db: AsyncSession = Depends(get_db)) -> User:
    """Get current user from JWT token"""
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, settings.JWT_SECRET_KEY, algorithms=[settings.JWT_ALGORITHM])
        username: str = payload.get("sub")
        user_id: int = payload.get("user_id")
        if username is None or user_id is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception

    user = await get_user_by_id(db, user_id)
    if user is None or user.username != username or not user.is_active:
        raise credentials_exception

    company_id = int(user.company_id or 1)
    db.info["company_id"] = company_id
    db.sync_session.info["company_id"] = company_id
    return user
