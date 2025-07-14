from fastapi import APIRouter
from app.api.endpoints import excel, auth, users, model_management, tracing, incremental_model

router = APIRouter()

router.include_router(auth.router, prefix="/auth", tags=["authentication"])
router.include_router(users.router, prefix="/users", tags=["users"])
router.include_router(excel.router, prefix="/excel", tags=["excel"])
router.include_router(model_management.router, prefix="/models", tags=["model-management"])
router.include_router(tracing.router, prefix="/tracing", tags=["llm-tracing"])
router.include_router(incremental_model.router, prefix="/incremental", tags=["incremental-models"])