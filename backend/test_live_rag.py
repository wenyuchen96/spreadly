#!/usr/bin/env python3
"""
Test if the live backend service has our RAG fixes
"""

import requests
import json

def test_live_rag():
    print("🔍 Testing if live backend has RAG fixes...")
    
    # Test basic health
    base_url = "https://2df4fc01760f.ngrok-free.app"
    headers = {"ngrok-skip-browser-warning": "true"}
    
    # Check RAG status
    try:
        response = requests.get(f"{base_url}/api/excel/rag/status", headers=headers)
        print(f"📊 RAG Status: {response.status_code}")
        if response.status_code == 200:
            data = response.json()
            print(f"   - RAG Enabled: {data.get('rag_enabled')}")
            print(f"   - Models: {data.get('vector_store_stats', {}).get('total_models')}")
            print(f"   - Status: {data.get('status')}")
        else:
            print(f"   - Error: {response.text}")
    except Exception as e:
        print(f"❌ RAG Status Error: {e}")
    
    # Test a simple DCF query via the Excel query endpoint
    print("\n🔍 Testing DCF query via Excel endpoint...")
    try:
        query_data = {
            "session_token": "test-session-123",
            "query": "create a simple dcf model"
        }
        
        response = requests.post(
            f"{base_url}/api/excel/query", 
            headers={**headers, "Content-Type": "application/json"},
            json=query_data,
            timeout=30
        )
        
        print(f"📊 Query Response: {response.status_code}")
        if response.status_code == 200:
            print("✅ Query succeeded - check traces.jsonl for RAG details")
        else:
            print(f"❌ Query failed: {response.text[:200]}...")
            
    except Exception as e:
        print(f"❌ Query Error: {e}")

if __name__ == "__main__":
    test_live_rag()