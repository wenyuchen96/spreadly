from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from app.core.database import get_db
from app.models.user import User
from app.models.session import Session as SessionModel
from app.api.endpoints.auth import oauth2_scheme
from app.services.auth_service import AuthService

router = APIRouter()

@router.get("/sessions")
async def get_user_sessions(
    token: str = Depends(oauth2_scheme),
    db: Session = Depends(get_db)
):
    """Get user's Excel processing sessions"""
    auth_service = AuthService()
    user = auth_service.get_current_user(db, token)
    
    if not user:
        raise HTTPException(status_code=401, detail="Authentication required")
    
    sessions = db.query(SessionModel).filter(SessionModel.user_id == user.id).all()
    
    return {
        "user_id": user.id,
        "sessions": [
            {
                "session_token": session.session_token,
                "file_name": session.file_name,
                "status": session.processing_status,
                "created_at": session.created_at
            }
            for session in sessions
        ]
    }

@router.get("/profile")
async def get_user_profile(
    token: str = Depends(oauth2_scheme),
    db: Session = Depends(get_db)
):
    """Get user profile information"""
    auth_service = AuthService()
    user = auth_service.get_current_user(db, token)
    
    if not user:
        raise HTTPException(status_code=401, detail="Authentication required")
    
    return {
        "id": user.id,
        "email": user.email,
        "full_name": user.full_name,
        "created_at": user.created_at,
        "is_active": user.is_active
    }