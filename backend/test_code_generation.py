#!/usr/bin/env python3
"""
Test script to verify code generation produces clean, executable JavaScript
"""

import asyncio
import sys
import os

# Add the backend directory to the path
sys.path.append('/Users/wenyuc/Dev/spreadly/backend')

from app.services.ai_service_simple import AIService
from app.api.endpoints.incremental_model import clean_generated_code

async def test_code_generation():
    """Test basic code generation"""
    
    print("🧪 Testing AI code generation...")
    
    # Initialize AI service
    ai_service = AIService()
    
    if not ai_service.client:
        print("❌ AI service not available (no API key)")
        return
    
    # Test basic chunk generation
    try:
        print("🔧 Generating test chunk...")
        
        chunk_code = await ai_service.generate_incremental_chunk(
            session_id=0,
            model_type="dcf",
            build_context="Generate simple headers for DCF model",
            workbook_context={"sheets": [{"name": "Sheet1", "data": [["DCF Model"]]}]},
            previous_errors=None
        )
        
        print(f"📝 Raw generated code ({len(chunk_code)} chars):")
        print("-" * 50)
        print(chunk_code)
        print("-" * 50)
        
        # Clean the code
        cleaned_code = clean_generated_code(chunk_code)
        
        print(f"✨ Cleaned code ({len(cleaned_code)} chars):")
        print("-" * 50)
        print(cleaned_code)
        print("-" * 50)
        
        # Basic validation
        if cleaned_code.strip().startswith('await Excel.run'):
            print("✅ Code starts correctly")
        else:
            print("❌ Code does not start with 'await Excel.run'")
            
        if cleaned_code.strip().endswith('});'):
            print("✅ Code ends correctly")
        else:
            print("❌ Code does not end with '});'")
            
        if 'Looking at' in cleaned_code or 'I can see' in cleaned_code:
            print("❌ Code contains explanatory text")
        else:
            print("✅ Code is clean of explanatory text")
            
        # Check for 2D array usage
        if '.values = [[' in cleaned_code:
            print("✅ Uses 2D arrays correctly")
        else:
            print("❌ May not be using 2D arrays")
        
        print("🎉 Test completed!")
        
    except Exception as e:
        print(f"❌ Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_code_generation())