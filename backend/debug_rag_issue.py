#!/usr/bin/env python3
"""
Debug why RAG isn't being triggered in the live service
"""

import asyncio
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app.services.ai_service_simple import AIService

async def debug_rag_issue():
    print("🔍 Debugging RAG Issue...")
    
    # Initialize AI service exactly like the live service does
    ai_service = AIService()
    
    print(f"📊 AI Service State:")
    print(f"   - RAG enabled: {ai_service.rag_enabled}")
    print(f"   - Vector store exists: {ai_service.vector_store is not None}")
    if ai_service.vector_store:
        print(f"   - Vector store available: {ai_service.vector_store.is_available()}")
        stats = ai_service.vector_store.get_stats()
        print(f"   - Vector store stats: {stats}")
    
    # Test the exact query from the trace
    query = "create a DCF model for a tech company in sheet3"
    print(f"\n🔍 Testing query: '{query}'")
    
    # Check detection logic
    model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv', 'irr', 'scenario analysis', 'monte carlo', 'sensitivity analysis']
    wants_model = any(keyword in query.lower() for keyword in model_keywords)
    print(f"📊 wants_model: {wants_model}")
    
    # Check all conditions for RAG trigger
    rag_conditions = {
        "wants_model": wants_model,
        "rag_enabled": ai_service.rag_enabled,
        "vector_store_exists": ai_service.vector_store is not None,
        "vector_store_available": ai_service.vector_store.is_available() if ai_service.vector_store else False
    }
    
    print(f"📊 RAG trigger conditions:")
    for condition, value in rag_conditions.items():
        print(f"   - {condition}: {value}")
    
    all_conditions_met = all(rag_conditions.values())
    print(f"📊 All conditions met: {all_conditions_met}")
    
    if all_conditions_met:
        print("\n🔍 Testing RAG search manually...")
        try:
            # Test the search that should happen
            from app.models.financial_model import ModelSearchQuery, ModelType, Industry, ComplexityLevel
            
            model_type = ModelType.DCF
            industry = Industry.TECHNOLOGY
            complexity = ComplexityLevel.INTERMEDIATE
            
            search_query = ModelSearchQuery(
                query_text=query,
                model_type=model_type,
                industry=industry,
                complexity=None,  # Don't filter by complexity
                min_rating=0.0,
                limit=3
            )
            
            print(f"📊 Search query: {search_query}")
            search_response = await ai_service.vector_store.search_models(search_query)
            print(f"🎯 Search results: {len(search_response.results)} models found")
            
            for i, result in enumerate(search_response.results, 1):
                print(f"  {i}. {result.model.name} (similarity: {result.similarity_score:.3f})")
            
        except Exception as e:
            print(f"❌ Search error: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("\n❌ RAG conditions not met - that's why it's not triggering")

if __name__ == "__main__":
    asyncio.run(debug_rag_issue())