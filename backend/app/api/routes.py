from fastapi import APIRouter
from app.api.endpoints import excel, auth, users

router = APIRouter()

router.include_router(auth.router, prefix="/auth", tags=["authentication"])
router.include_router(users.router, prefix="/users", tags=["users"])
router.include_router(excel.router, prefix="/excel", tags=["excel"])