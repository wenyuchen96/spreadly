#!/usr/bin/env python3
"""
Test script to check RAG vector store status and DCF model availability
"""

import asyncio
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app.services.model_vector_store import get_vector_store
from app.services.model_curator import get_model_curator
from app.models.financial_model import ModelSearchQuery, ModelType

async def test_rag_system():
    print("🔍 Testing RAG System...")
    
    # Initialize vector store
    vector_store = get_vector_store()
    print(f"✅ Vector store available: {vector_store.is_available()}")
    
    if not vector_store.is_available():
        print("❌ Vector store not available - missing dependencies?")
        return
    
    # Get current stats
    stats = vector_store.get_stats()
    print(f"📊 Current stats: {stats}")
    
    # Reset and reinitialize to fix enum string issues
    print("🔄 Resetting vector store to fix enum storage...")
    await vector_store.reset_store()
    
    print("📚 Reinitializing with corrected templates...")
    curator = get_model_curator()
    results = await curator.initialize_model_library()
    print(f"✅ Initialized: {results}")
    
    # Get updated stats
    stats = vector_store.get_stats()
    print(f"📊 Updated stats: {stats}")
    
    # Test searches with different levels of filtering
    from app.models.financial_model import Industry, ComplexityLevel
    
    print("\n🔍 Testing broad search (no filters)...")
    broad_query = ModelSearchQuery(
        query_text="dcf model",
        model_type=None,
        industry=None, 
        complexity=None,
        min_rating=0.0,
        limit=5
    )
    
    broad_results = await vector_store.search_models(broad_query)
    print(f"🎯 Broad search results: {len(broad_results.results)} models found")
    
    for i, result in enumerate(broad_results.results, 1):
        print(f"  {i}. {result.model.name} (similarity: {result.similarity_score:.3f})")
        print(f"     Type: {result.model.model_type}, Industry: {result.model.industry}")
    
    print("\n🔍 Testing DCF-only filter...")
    dcf_query = ModelSearchQuery(
        query_text="create a dcf model for technology company",
        model_type=ModelType.DCF,
        industry=None,  # Remove industry filter
        complexity=None,  # Remove complexity filter
        min_rating=0.0,
        limit=3
    )
    
    dcf_results = await vector_store.search_models(dcf_query)
    print(f"🎯 DCF-only search results: {len(dcf_results.results)} models found")
    
    for i, result in enumerate(dcf_results.results, 1):
        print(f"  {i}. {result.model.name} (similarity: {result.similarity_score:.3f})")
        print(f"     Type: {result.model.model_type}, Industry: {result.model.industry}")
        print(f"     Keywords: {result.model.keywords[:5]}")
    
    print(f"\n✅ Test complete! Vector store has {stats.get('total_models', 0)} models")

if __name__ == "__main__":
    asyncio.run(test_rag_system())