"""
Model Vector Store Service for RAG Implementation
Handles embedding generation, storage, and similarity search for financial models
"""

import os
import json
import asyncio
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
import logging

try:
    import chromadb
    from chromadb.config import Settings
    from sentence_transformers import SentenceTransformer
    DEPENDENCIES_AVAILABLE = True
except ImportError:
    DEPENDENCIES_AVAILABLE = False
    logging.warning("RAG dependencies not available. Run: pip install chromadb sentence-transformers")

from app.models.financial_model import (
    FinancialModel, 
    ModelSearchQuery, 
    ModelSearchResult, 
    ModelSearchResponse,
    ModelType,
    Industry,
    ComplexityLevel
)


class ModelVectorStore:
    """
    Vector store for financial model templates using ChromaDB and sentence-transformers
    """
    
    def __init__(self, persist_directory: str = "./chroma_db"):
        self.persist_directory = persist_directory
        self.collection_name = "financial_models"
        self.embedding_model_name = "all-MiniLM-L6-v2"  # Lightweight, fast model
        
        if not DEPENDENCIES_AVAILABLE:
            logging.warning("RAG dependencies not available. Vector store will not function.")
            self.client = None
            self.collection = None
            self.embeddings = None
            return
            
        # Initialize embedding model
        self.embeddings = SentenceTransformer(self.embedding_model_name)
        
        # Initialize ChromaDB
        self.client = chromadb.PersistentClient(
            path=persist_directory,
            settings=Settings(allow_reset=True)
        )
        
        # Get or create collection
        self.collection = self.client.get_or_create_collection(
            name=self.collection_name,
            metadata={"description": "Financial model templates for RAG"}
        )
        
        logging.info(f"ModelVectorStore initialized with {self.collection.count()} models")
    
    def is_available(self) -> bool:
        """Check if vector store is available"""
        return DEPENDENCIES_AVAILABLE and self.client is not None
    
    async def add_model(self, model: FinancialModel) -> bool:
        """Add a financial model to the vector store"""
        if not self.is_available():
            logging.warning("Vector store not available, skipping model addition")
            return False
            
        try:
            # Create searchable text from model
            searchable_text = self._create_searchable_text(model)
            
            # Generate embedding
            embedding = self.embeddings.encode(searchable_text).tolist()
            
            # Prepare metadata for ChromaDB (must be JSON serializable)
            metadata = {
                "model_type": model.model_type.value if hasattr(model.model_type, 'value') else str(model.model_type),
                "industry": model.industry.value if hasattr(model.industry, 'value') else str(model.industry),
                "complexity": model.complexity.value if hasattr(model.complexity, 'value') else str(model.complexity),
                "user_rating": model.performance.user_rating,
                "execution_success_rate": model.performance.execution_success_rate,
                "usage_count": model.performance.usage_count,
                "created_at": model.created_at.isoformat(),
                "components": json.dumps(model.metadata.components),
                "excel_functions": json.dumps(model.metadata.excel_functions),
                "keywords": json.dumps(model.keywords),
                "tags": json.dumps(model.tags)
            }
            
            # Add to collection
            self.collection.add(
                ids=[model.id],
                embeddings=[embedding],
                documents=[searchable_text],
                metadatas=[metadata]
            )
            
            logging.info(f"Added model {model.id} to vector store")
            return True
            
        except Exception as e:
            logging.error(f"Error adding model {model.id}: {e}")
            return False
    
    async def search_models(self, query: ModelSearchQuery) -> ModelSearchResponse:
        """Search for similar financial models"""
        start_time = datetime.utcnow()
        
        if not self.is_available():
            return ModelSearchResponse(
                query=query,
                results=[],
                total_found=0,
                search_time_ms=0.0,
                retrieval_strategy="vector_store_unavailable"
            )
        
        try:
            # Generate query embedding
            query_embedding = self.embeddings.encode(query.query_text).tolist()
            
            # Build where clause for metadata filtering - ChromaDB format
            where_conditions = []
            if query.model_type:
                where_conditions.append({"model_type": {"$eq": query.model_type.value if hasattr(query.model_type, 'value') else str(query.model_type)}})
            if query.industry:
                where_conditions.append({"industry": {"$eq": query.industry.value if hasattr(query.industry, 'value') else str(query.industry)}})
            if query.complexity:
                where_conditions.append({"complexity": {"$eq": query.complexity.value if hasattr(query.complexity, 'value') else str(query.complexity)}})
            if query.min_rating:
                where_conditions.append({"user_rating": {"$gte": query.min_rating}})
            
            # Combine conditions with $and if multiple conditions exist
            if len(where_conditions) > 1:
                where_clause = {"$and": where_conditions}
            elif len(where_conditions) == 1:
                where_clause = where_conditions[0]
            else:
                where_clause = None
            
            # Perform similarity search
            results = self.collection.query(
                query_embeddings=[query_embedding],
                n_results=query.limit,
                where=where_clause,
                include=["documents", "metadatas", "distances"]
            )
            
            # Convert results to ModelSearchResult objects
            search_results = []
            for i in range(len(results['ids'][0])):
                # Reconstruct FinancialModel from stored data
                model_id = results['ids'][0][i]
                metadata = results['metadatas'][0][i]
                distance = results['distances'][0][i]
                similarity_score = max(0.0, 1.0 - distance)  # Convert distance to similarity, ensuring non-negative
                
                # Create a minimal FinancialModel for results
                # In production, you might want to store full models or reconstruct them
                model = FinancialModel(
                    id=model_id,
                    name=f"Model {model_id}",
                    description=results['documents'][0][i][:200] + "...",
                    model_type=metadata['model_type'],
                    industry=metadata['industry'],
                    complexity=metadata['complexity'],
                    excel_code="# Model code would be retrieved from full storage",
                    business_description=results['documents'][0][i],
                    sample_inputs={},
                    expected_outputs={},
                    metadata={
                        "components": json.loads(metadata.get('components', '[]')),
                        "excel_functions": json.loads(metadata.get('excel_functions', '[]')),
                        "formatting_features": [],
                        "business_assumptions": [],
                        "time_horizon_years": None,
                        "currencies": ["USD"],
                        "regions": ["global"]
                    },
                    performance={
                        "execution_success_rate": metadata.get('execution_success_rate', 0.0),
                        "user_rating": metadata.get('user_rating', 0.0),
                        "usage_count": metadata.get('usage_count', 0),
                        "last_used": None,
                        "error_count": 0,
                        "modification_frequency": 0.0
                    },
                    created_by="system",
                    keywords=json.loads(metadata.get('keywords', '[]')),
                    tags=json.loads(metadata.get('tags', '[]'))
                )
                
                search_results.append(ModelSearchResult(
                    model=model,
                    similarity_score=similarity_score,
                    relevance_explanation=f"Similarity: {similarity_score:.3f}, Type: {metadata['model_type']}"
                ))
            
            end_time = datetime.utcnow()
            search_time_ms = (end_time - start_time).total_seconds() * 1000
            
            return ModelSearchResponse(
                query=query,
                results=search_results,
                total_found=len(search_results),
                search_time_ms=search_time_ms,
                retrieval_strategy="semantic_similarity_with_metadata_filtering"
            )
            
        except Exception as e:
            logging.error(f"Error searching models: {e}")
            return ModelSearchResponse(
                query=query,
                results=[],
                total_found=0,
                search_time_ms=0.0,
                retrieval_strategy="error_fallback"
            )
    
    async def update_model_performance(self, model_id: str, success: bool, user_rating: Optional[float] = None):
        """Update model performance metrics"""
        if not self.is_available():
            return
            
        try:
            # Get current model
            result = self.collection.get(ids=[model_id], include=["metadatas"])
            if not result['ids']:
                logging.warning(f"Model {model_id} not found for performance update")
                return
                
            metadata = result['metadatas'][0]
            
            # Update metrics
            current_usage = metadata.get('usage_count', 0)
            current_errors = metadata.get('error_count', 0)
            current_success_rate = metadata.get('execution_success_rate', 0.0)
            
            new_usage = current_usage + 1
            new_errors = current_errors + (0 if success else 1)
            new_success_rate = (current_usage * current_success_rate + (1 if success else 0)) / new_usage
            
            # Update metadata
            updated_metadata = metadata.copy()
            updated_metadata['usage_count'] = new_usage
            updated_metadata['error_count'] = new_errors
            updated_metadata['execution_success_rate'] = new_success_rate
            
            if user_rating is not None:
                # Simple average rating update (in production, might want weighted average)
                current_rating = metadata.get('user_rating', 0.0)
                updated_metadata['user_rating'] = (current_rating + user_rating) / 2
            
            # Update in collection
            self.collection.update(
                ids=[model_id],
                metadatas=[updated_metadata]
            )
            
            logging.info(f"Updated performance for model {model_id}")
            
        except Exception as e:
            logging.error(f"Error updating model performance: {e}")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get vector store statistics"""
        if not self.is_available():
            return {"status": "unavailable", "total_models": 0}
            
        try:
            count = self.collection.count()
            return {
                "status": "available",
                "total_models": count,
                "embedding_model": self.embedding_model_name,
                "collection_name": self.collection_name
            }
        except Exception as e:
            return {"status": "error", "error": str(e)}
    
    def _create_searchable_text(self, model: FinancialModel) -> str:
        """Create searchable text representation of a model"""
        text_parts = [
            model.name,
            model.description,
            model.business_description,
            model.model_type,
            model.industry,
            model.complexity,
            " ".join(model.keywords),
            " ".join(model.tags),
            " ".join(model.metadata.components),
            " ".join(model.metadata.excel_functions)
        ]
        
        return " ".join(filter(None, text_parts))
    
    async def reset_store(self):
        """Reset the vector store (useful for development)"""
        if not self.is_available():
            return
            
        try:
            self.client.reset()
            self.collection = self.client.get_or_create_collection(
                name=self.collection_name,
                metadata={"description": "Financial model templates for RAG"}
            )
            logging.info("Vector store reset successfully")
        except Exception as e:
            logging.error(f"Error resetting vector store: {e}")


# Singleton instance
_vector_store_instance = None

def get_vector_store() -> ModelVectorStore:
    """Get singleton vector store instance"""
    global _vector_store_instance
    if _vector_store_instance is None:
        _vector_store_instance = ModelVectorStore()
    return _vector_store_instance