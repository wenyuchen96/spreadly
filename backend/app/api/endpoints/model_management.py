"""
API endpoints for managing financial models in the RAG system
"""

from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from typing import List, Dict, Any, Optional
import tempfile
import os
from pathlib import Path

from app.services.model_vector_store import get_vector_store
from app.models.financial_model import FinancialModel, ModelType, Industry, ComplexityLevel

router = APIRouter()

@router.post("/models/upload-xlsx")
async def upload_xlsx_model(
    file: UploadFile = File(...),
    model_type: str = Form(...),
    industry: str = Form(default="general"),
    complexity: str = Form(default="intermediate"),
    name: Optional[str] = Form(default=None),
    description: Optional[str] = Form(default=None)
):
    """
    Upload an XLSX file and convert it to a RAG model
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
    
    # Validate enum values
    try:
        model_type_enum = ModelType(model_type.lower())
        industry_enum = Industry(industry.lower())
        complexity_enum = ComplexityLevel(complexity.lower())
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"Invalid enum value: {str(e)}")
    
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        content = await file.read()
        tmp_file.write(content)
        tmp_file_path = tmp_file.name
    
    try:
        # Import converter here to avoid circular imports
        from tools.xlsx_to_model_converter import XLSXToModelConverter
        
        converter = XLSXToModelConverter()
        
        # Generate model ID
        file_stem = Path(file.filename).stem
        model_id = f"uploaded_{file_stem}_{model_type}"
        
        # Convert XLSX to model
        model = converter.convert_xlsx_to_model(
            tmp_file_path,
            model_id,
            model_type_enum,
            industry_enum,
            complexity_enum
        )
        
        # Override name/description if provided
        if name:
            model.name = name
        if description:
            model.description = description
        
        # Add to vector store
        success = await vector_store.add_model(model)
        
        if not success:
            raise HTTPException(status_code=500, detail="Failed to add model to vector store")
        
        return {
            "status": "success",
            "message": f"Model '{model.name}' uploaded successfully",
            "model_id": model.id,
            "model_details": {
                "name": model.name,
                "type": model.model_type,
                "industry": model.industry,
                "complexity": model.complexity,
                "keywords": model.keywords
            }
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")
    
    finally:
        # Clean up temporary file
        try:
            os.unlink(tmp_file_path)
        except:
            pass

@router.post("/models/bulk-upload")
async def bulk_upload_models(
    files: List[UploadFile] = File(...),
    auto_detect: bool = Form(default=True)
):
    """
    Upload multiple XLSX files at once
    """
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    results = {
        "total_files": len(files),
        "successful": 0,
        "failed": 0,
        "results": [],
        "errors": []
    }
    
    from tools.xlsx_to_model_converter import XLSXToModelConverter
    converter = XLSXToModelConverter()
    
    for file in files:
        if not file.filename.endswith(('.xlsx', '.xls')):
            results["failed"] += 1
            results["errors"].append(f"Skipped {file.filename}: not an Excel file")
            continue
        
        # Save file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            content = await file.read()
            tmp_file.write(content)
            tmp_file_path = tmp_file.name
        
        try:
            # Auto-detect or use defaults
            if auto_detect:
                # Use the bulk loader's detection logic
                from tools.bulk_model_loader import BulkModelLoader
                loader = BulkModelLoader()
                model_type, industry, complexity = loader._detect_from_filename(file.filename)
            else:
                model_type = ModelType.DCF
                industry = Industry.GENERAL
                complexity = ComplexityLevel.INTERMEDIATE
            
            # Convert and add model
            file_stem = Path(file.filename).stem
            model_id = f"bulk_uploaded_{file_stem}_{model_type}"
            
            model = converter.convert_xlsx_to_model(
                tmp_file_path,
                model_id,
                model_type,
                industry,
                complexity
            )
            
            success = await vector_store.add_model(model)
            
            if success:
                results["successful"] += 1
                results["results"].append({
                    "filename": file.filename,
                    "model_id": model.id,
                    "status": "success",
                    "detected_type": model_type,
                    "detected_industry": industry
                })
            else:
                results["failed"] += 1
                results["errors"].append(f"Vector store failed for {file.filename}")
            
        except Exception as e:
            results["failed"] += 1
            results["errors"].append(f"Error processing {file.filename}: {str(e)}")
        
        finally:
            try:
                os.unlink(tmp_file_path)
            except:
                pass
    
    return results

@router.get("/models/list")
async def list_models():
    """
    List all models in the vector store
    """
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    try:
        # Get all models from ChromaDB
        collection = vector_store.collection
        results = collection.get(include=['metadatas', 'documents'])
        
        models = []
        for i, doc_id in enumerate(results['ids']):
            metadata = results['metadatas'][i]
            models.append({
                "id": doc_id,
                "name": metadata.get('name', 'Unknown'),
                "model_type": metadata.get('model_type', 'unknown'),
                "industry": metadata.get('industry', 'unknown'),
                "complexity": metadata.get('complexity', 'unknown'),
                "user_rating": metadata.get('user_rating', 0),
                "usage_count": metadata.get('usage_count', 0),
                "created_at": metadata.get('created_at', 'unknown')
            })
        
        return {
            "total_models": len(models),
            "models": models
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error retrieving models: {str(e)}")

@router.delete("/models/{model_id}")
async def delete_model(model_id: str):
    """
    Delete a specific model from the vector store
    """
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    try:
        collection = vector_store.collection
        collection.delete(ids=[model_id])
        
        return {
            "status": "success",
            "message": f"Model {model_id} deleted successfully"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error deleting model: {str(e)}")

@router.get("/models/stats")
async def get_model_stats():
    """
    Get statistics about the model collection
    """
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    return vector_store.get_stats()

@router.get("/models/search")
async def search_models(
    query: str,
    model_type: Optional[str] = None,
    industry: Optional[str] = None,
    limit: int = 5
):
    """
    Search models using semantic similarity
    """
    vector_store = get_vector_store()
    if not vector_store.is_available():
        raise HTTPException(status_code=503, detail="Vector store not available")
    
    from app.models.financial_model import ModelSearchQuery
    
    # Convert string enums to enum objects
    model_type_enum = ModelType(model_type) if model_type else None
    industry_enum = Industry(industry) if industry else None
    
    search_query = ModelSearchQuery(
        query_text=query,
        model_type=model_type_enum,
        industry=industry_enum,
        limit=limit
    )
    
    try:
        results = await vector_store.search_models(search_query)
        
        return {
            "query": query,
            "total_results": len(results.results),
            "results": [
                {
                    "model_id": result.model.id,
                    "name": result.model.name,
                    "similarity_score": result.similarity_score,
                    "model_type": result.model.model_type,
                    "industry": result.model.industry,
                    "description": result.model.description[:200] + "..." if len(result.model.description) > 200 else result.model.description
                }
                for result in results.results
            ]
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error searching models: {str(e)}")