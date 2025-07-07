from fastapi import APIRouter, Depends, UploadFile, File, HTTPException, Request
from sqlalchemy.orm import Session
from typing import List, Dict, Any
from app.core.database import get_db
from app.services.excel_service import ExcelService
from app.services.ai_service_simple import AIService
from app.models.session import Session as SessionModel
from app.models.spreadsheet import Spreadsheet
import uuid

router = APIRouter()

@router.get("/test")
async def test_connection():
    """Simple test endpoint to verify frontend-backend connection"""
    return {
        "status": "success",
        "message": "🎉 REAL backend connection successful!",
        "timestamp": "2025-01-07T00:00:00Z",
        "ai_powered": True
    }

@router.post("/upload")
async def upload_excel(
    request: Request,
    file: UploadFile = File(None),
    db: Session = Depends(get_db)
):
    """Upload and process Excel file or data"""
    session_token = str(uuid.uuid4())
    
    # Handle both file upload and JSON data
    if file and file.filename:
        # Traditional file upload
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files are allowed")
        
        file_name = file.filename
        # Create session record
        session = SessionModel(
            session_token=session_token,
            file_name=file_name,
            processing_status="processing"
        )
        db.add(session)
        db.commit()
        
        try:
            # Process Excel file
            excel_service = ExcelService()
            spreadsheet_data = await excel_service.process_file(file, session.id)
            
            # Update session status
            session.processing_status = "completed"
            session.analysis_result = spreadsheet_data.get("summary", "")
            db.commit()
            
            return {
                "session_token": session_token,
                "message": "File uploaded and processed successfully",
                "data": spreadsheet_data
            }
        except Exception as e:
            session.processing_status = "failed"
            db.commit()
            raise HTTPException(status_code=500, detail=str(e))
    else:
        # Handle JSON data from frontend
        try:
            request_body = await request.json()
            if not request_body:
                raise HTTPException(status_code=400, detail="No data provided")
            
            file_name = request_body.get("file_name", "spreadsheet.xlsx")
            excel_data = request_body.get("data", [])
            
            if not excel_data:
                raise HTTPException(status_code=400, detail="No Excel data provided")
            
            # Create session record
            session = SessionModel(
                session_token=session_token,
                file_name=file_name,
                processing_status="processing"
            )
            db.add(session)
            db.commit()
            
            # Process data directly
            excel_service = ExcelService()
            spreadsheet_data = await excel_service.process_data(excel_data, session.id, file_name)
            
            # Update session status
            session.processing_status = "completed"
            session.analysis_result = spreadsheet_data.get("summary", "")
            db.commit()
            
            return {
                "session_token": session_token,
                "message": "Data processed successfully",
                "data": spreadsheet_data
            }
        except Exception as e:
            if 'session' in locals():
                session.processing_status = "failed"
                db.commit()
            raise HTTPException(status_code=500, detail=str(e))

@router.get("/analyze/{session_token}")
async def analyze_spreadsheet(
    session_token: str,
    db: Session = Depends(get_db)
):
    """Get AI-powered analysis of spreadsheet"""
    session = db.query(SessionModel).filter(SessionModel.session_token == session_token).first()
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    spreadsheet = db.query(Spreadsheet).filter(Spreadsheet.session_id == session.id).first()
    if not spreadsheet:
        raise HTTPException(status_code=404, detail="Spreadsheet not found")
    
    ai_service = AIService()
    analysis = await ai_service.analyze_spreadsheet(spreadsheet)
    
    return {
        "session_token": session_token,
        "analysis": analysis,
        "insights": spreadsheet.ai_insights
    }

@router.post("/query")
async def query_spreadsheet(
    query_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """Natural language query on spreadsheet data"""
    session_token = query_data.get("session_token")
    query = query_data.get("query")
    
    if not session_token or not query:
        raise HTTPException(status_code=400, detail="Session token and query are required")
    
    session = db.query(SessionModel).filter(SessionModel.session_token == session_token).first()
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    ai_service = AIService()
    result = await ai_service.process_natural_language_query(session.id, query)
    
    return {
        "session_token": session_token,
        "query": query,
        "result": result
    }

@router.get("/formulas")
async def generate_formulas(
    description: str,
    context: str = None,
    db: Session = Depends(get_db)
):
    """Generate Excel formulas from natural language description"""
    ai_service = AIService()
    formulas = await ai_service.generate_formulas(description, context)
    
    return {
        "description": description,
        "formulas": formulas
    }

@router.post("/search")
async def search_patterns(
    search_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """Vector search for similar spreadsheet patterns"""
    query = search_data.get("query")
    pattern_type = search_data.get("type", "all")
    
    if not query:
        raise HTTPException(status_code=400, detail="Search query is required")
    
    ai_service = AIService()
    patterns = await ai_service.search_similar_patterns(query, pattern_type)
    
    return {
        "query": query,
        "patterns": patterns
    }