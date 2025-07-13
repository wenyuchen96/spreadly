#!/usr/bin/env python3
"""
Test end-to-end RAG flow simulation
"""

import asyncio
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app.services.ai_service_simple import AIService

async def test_end_to_end_rag():
    print("🔍 Testing End-to-End RAG Flow...")
    
    # Initialize AI service
    ai_service = AIService()
    print(f"✅ AI Service initialized. RAG enabled: {ai_service.rag_enabled}")
    print(f"✅ Vector store available: {ai_service.vector_store.is_available() if ai_service.vector_store else False}")
    
    # Test queries that should trigger RAG
    test_queries = [
        "create a dcf model for a technology company",
        "I need a DCF valuation model",
        "build a discounted cash flow model",
        "make an NPV analysis model"
    ]
    
    for query in test_queries:
        print(f"\n🔍 Testing query: '{query}'")
        
        # Check if it would trigger model search
        model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv']
        wants_model = any(keyword in query.lower() for keyword in model_keywords)
        print(f"📊 Would trigger RAG: {wants_model}")
        
        if wants_model and ai_service.rag_enabled and ai_service.vector_store:
            # Simulate the RAG search that would happen
            from app.models.financial_model import ModelSearchQuery, ModelType, Industry, ComplexityLevel
            
            # Detect model characteristics (from ai_service_simple.py logic)
            query_lower = query.lower()
            
            model_type = None
            if any(word in query_lower for word in ['dcf', 'discounted cash flow']):
                model_type = ModelType.DCF
            elif any(word in query_lower for word in ['npv', 'net present value']):
                model_type = ModelType.NPV
                
            industry = Industry.GENERAL
            if any(word in query_lower for word in ['tech', 'technology']):
                industry = Industry.TECHNOLOGY
                
            complexity = ComplexityLevel.INTERMEDIATE
            
            print(f"📊 Detected - Type: {model_type}, Industry: {industry}, Complexity: {complexity}")
            
            # Perform search
            if model_type:
                search_query = ModelSearchQuery(
                    query_text=query,
                    model_type=model_type,
                    industry=industry if industry != Industry.GENERAL else None,
                    complexity=None,  # Don't filter by complexity for broader results
                    min_rating=0.0,
                    limit=3
                )
                
                search_response = await ai_service.vector_store.search_models(search_query)
                print(f"🎯 RAG Results: {len(search_response.results)} models retrieved")
                
                for i, result in enumerate(search_response.results, 1):
                    print(f"  {i}. {result.model.name} (similarity: {result.similarity_score:.3f})")
                    print(f"     Keywords: {result.model.keywords[:3]}")
            else:
                print("📊 No specific model type detected, would do general search")
    
    print(f"\n✅ End-to-end test complete!")

if __name__ == "__main__":
    asyncio.run(test_end_to_end_rag())