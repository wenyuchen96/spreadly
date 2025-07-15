from fastapi import APIRouter, Depends, UploadFile, File, HTTPException, Request
from sqlalchemy.orm import Session
from typing import List, Dict, Any
from app.core.database import get_db
from app.services.excel_service import ExcelService
from app.services.ai_service_simple import AIService
from app.services.model_vector_store import get_vector_store
from app.services.model_curator import get_model_curator
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
            
            # Convert any numpy/pandas objects to JSON-serializable format
            serializable_data = {}
            if spreadsheet_data:
                for key, value in spreadsheet_data.items():
                    if hasattr(value, 'tolist'):  # numpy array
                        serializable_data[key] = value.tolist()
                    elif hasattr(value, 'to_dict'):  # pandas DataFrame
                        serializable_data[key] = value.to_dict()
                    elif hasattr(value, 'item'):  # numpy scalar
                        serializable_data[key] = value.item()
                    else:
                        serializable_data[key] = str(value)  # Fallback to string
            
            return {
                "session_token": session_token,
                "message": "Data processed successfully",
                "data": serializable_data
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
    """Natural language query on spreadsheet data with comprehensive workbook context"""
    session_token = query_data.get("session_token")
    query = query_data.get("query")
    workbook_context = query_data.get("workbook_context")  # New comprehensive context
    
    if not session_token or not query:
        raise HTTPException(status_code=400, detail="Session token and query are required")
    
    session = db.query(SessionModel).filter(SessionModel.session_token == session_token).first()
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    print(f"🔍 Backend: Processing query with workbook context: {bool(workbook_context)}")
    if workbook_context:
        print(f"📊 Backend: Context includes {len(workbook_context.get('sheets', []))} sheets, {len(workbook_context.get('tables', []))} tables")
    
    ai_service = AIService()
    result = await ai_service.process_natural_language_query(session.id, query, workbook_context)
    
    # Check if result is raw JavaScript code (for financial models)
    if isinstance(result, str) and 'Excel.run' in result:
        print("🔍 Backend: Detected raw JavaScript financial model response")
        return {
            "session_token": session_token,
            "query": query,
            "result": result  # Return raw JavaScript code directly
        }
    
    # Handle different response types properly
    if isinstance(result, str):
        # Plain text response from AI - wrap it in expected structure
        return {
            "session_token": session_token,
            "query": query,
            "result": {
                "answer": result
            }
        }
    else:
        # Structured response (dict) - return as-is
        return {
            "session_token": session_token,
            "query": query,
            "result": result
        }

@router.post("/web-search")
async def web_search_query(
    search_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """Perform web search enhanced query for current market data, trends, and research"""
    query = search_data.get("query")
    session_token = search_data.get("session_token", None)
    
    if not query:
        raise HTTPException(status_code=400, detail="Query is required")
    
    print(f"🌐 Backend: Processing web search query: {query}")
    
    ai_service = AIService()
    
    # Force web search for this endpoint by temporarily modifying the query
    web_enhanced_query = f"Please search the web for current information about: {query}. Provide a clear, direct answer without JSON formatting."
    
    try:
        result = await ai_service.process_natural_language_query(
            session_id=None if not session_token else session_token, 
            query=web_enhanced_query, 
            workbook_context=None
        )
        
        # Handle web search response properly
        if isinstance(result, str):
            formatted_result = {"answer": result}
        else:
            formatted_result = result
            
        return {
            "query": query,
            "result": formatted_result,
            "web_search_enabled": True,
            "session_token": session_token
        }
    except Exception as e:
        print(f"🚨 Web search error: {e}")
        raise HTTPException(status_code=500, detail=f"Web search failed: {str(e)}")

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

@router.get("/rag/status")
async def rag_status():
    """Get RAG system status and statistics"""
    try:
        vector_store = get_vector_store()
        stats = vector_store.get_stats()
        
        return {
            "rag_enabled": vector_store.is_available(),
            "vector_store_stats": stats,
            "status": "operational" if vector_store.is_available() else "unavailable"
        }
    except Exception as e:
        return {
            "rag_enabled": False,
            "error": str(e),
            "status": "error"
        }

@router.post("/rag/initialize")
async def initialize_rag_library():
    """Initialize RAG library with professional model templates"""
    try:
        vector_store = get_vector_store()
        
        if not vector_store.is_available():
            raise HTTPException(
                status_code=503, 
                detail="Vector store not available. Make sure RAG dependencies are installed."
            )
        
        model_curator = get_model_curator()
        results = await model_curator.initialize_model_library()
        
        return {
            "status": "success",
            "message": f"Initialized RAG library with {results['total_added']} models",
            "details": results
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error initializing RAG library: {str(e)}")

@router.delete("/rag/reset")
async def reset_rag_library():
    """Reset RAG library (development use only)"""
    try:
        vector_store = get_vector_store()
        
        if not vector_store.is_available():
            raise HTTPException(
                status_code=503, 
                detail="Vector store not available"
            )
        
        await vector_store.reset_store()
        
        return {
            "status": "success",
            "message": "RAG library reset successfully"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error resetting RAG library: {str(e)}")