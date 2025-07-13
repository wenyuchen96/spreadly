"""
API endpoints for LLM tracing and monitoring
"""

from fastapi import APIRouter, HTTPException
from typing import List, Dict, Any, Optional
import json
from datetime import datetime, timedelta

from app.core.tracing import local_storage

router = APIRouter()

@router.get("/traces/recent")
async def get_recent_traces(limit: int = 50):
    """Get recent LLM traces"""
    try:
        traces = local_storage.get_recent_traces(limit)
        return {
            "total_traces": len(traces),
            "traces": traces
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error retrieving traces: {str(e)}")

@router.get("/traces/stats")
async def get_trace_stats():
    """Get statistics about LLM usage"""
    try:
        traces = local_storage.get_recent_traces(1000)  # Get more for stats
        
        if not traces:
            return {
                "total_calls": 0,
                "avg_duration": 0,
                "success_rate": 0,
                "models_used": [],
                "operations": {}
            }
        
        # Calculate statistics
        total_calls = len(traces)
        successful_calls = sum(1 for trace in traces if trace.get('success', False))
        avg_duration = sum(trace.get('duration_seconds', 0) for trace in traces) / total_calls
        
        # Count operations
        operations = {}
        models_used = set()
        total_tokens = 0
        
        for trace in traces:
            operation = trace.get('operation', 'unknown')
            operations[operation] = operations.get(operation, 0) + 1
            
            model = trace.get('model', 'unknown')
            models_used.add(model)
            
            tokens = trace.get('tokens_used')
            if tokens:
                total_tokens += tokens
        
        return {
            "total_calls": total_calls,
            "successful_calls": successful_calls,
            "success_rate": successful_calls / total_calls if total_calls > 0 else 0,
            "avg_duration_seconds": avg_duration,
            "total_tokens_used": total_tokens,
            "models_used": list(models_used),
            "operations": operations,
            "last_24h_calls": len([
                trace for trace in traces 
                if datetime.fromisoformat(trace['timestamp'].replace('Z', '+00:00')) > 
                   datetime.now() - timedelta(hours=24)
            ])
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error calculating stats: {str(e)}")

@router.get("/traces/operations/{operation}")
async def get_traces_by_operation(operation: str, limit: int = 20):
    """Get traces filtered by operation type"""
    try:
        all_traces = local_storage.get_recent_traces(500)
        filtered_traces = [
            trace for trace in all_traces 
            if trace.get('operation') == operation
        ][-limit:]
        
        return {
            "operation": operation,
            "total_traces": len(filtered_traces),
            "traces": filtered_traces
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error retrieving traces: {str(e)}")

@router.get("/traces/errors")
async def get_error_traces(limit: int = 20):
    """Get traces that resulted in errors"""
    try:
        all_traces = local_storage.get_recent_traces(500)
        error_traces = [
            trace for trace in all_traces 
            if not trace.get('success', True)
        ][-limit:]
        
        # Group errors by type
        error_summary = {}
        for trace in error_traces:
            error = trace.get('error', 'Unknown error')
            error_summary[error] = error_summary.get(error, 0) + 1
        
        return {
            "total_errors": len(error_traces),
            "error_summary": error_summary,
            "recent_errors": error_traces
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error retrieving error traces: {str(e)}")

@router.get("/traces/performance")
async def get_performance_metrics():
    """Get performance metrics for LLM calls"""
    try:
        traces = local_storage.get_recent_traces(200)
        
        if not traces:
            return {"message": "No traces available"}
        
        # Calculate performance metrics
        durations = [trace.get('duration_seconds', 0) for trace in traces if trace.get('duration_seconds')]
        token_counts = [trace.get('tokens_used', 0) for trace in traces if trace.get('tokens_used')]
        
        performance = {}
        
        if durations:
            durations.sort()
            performance['response_time'] = {
                "min": min(durations),
                "max": max(durations),
                "avg": sum(durations) / len(durations),
                "p50": durations[len(durations) // 2],
                "p95": durations[int(len(durations) * 0.95)],
                "p99": durations[int(len(durations) * 0.99)]
            }
        
        if token_counts:
            token_counts.sort()
            performance['token_usage'] = {
                "min": min(token_counts),
                "max": max(token_counts),
                "avg": sum(token_counts) / len(token_counts),
                "total": sum(token_counts)
            }
        
        # Group by operation for detailed metrics
        by_operation = {}
        for trace in traces:
            operation = trace.get('operation', 'unknown')
            if operation not in by_operation:
                by_operation[operation] = {
                    'count': 0,
                    'avg_duration': 0,
                    'success_count': 0
                }
            
            by_operation[operation]['count'] += 1
            by_operation[operation]['avg_duration'] += trace.get('duration_seconds', 0)
            if trace.get('success', False):
                by_operation[operation]['success_count'] += 1
        
        # Calculate averages
        for operation, stats in by_operation.items():
            if stats['count'] > 0:
                stats['avg_duration'] = stats['avg_duration'] / stats['count']
                stats['success_rate'] = stats['success_count'] / stats['count']
        
        performance['by_operation'] = by_operation
        
        return performance
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error calculating performance: {str(e)}")

@router.delete("/traces/clear")
async def clear_traces():
    """Clear all stored traces (development only)"""
    try:
        import os
        if os.path.exists(local_storage.storage_file):
            os.remove(local_storage.storage_file)
        
        return {"message": "All traces cleared successfully"}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error clearing traces: {str(e)}")

@router.get("/traces/live")
async def get_live_trace_info():
    """Get live information about current tracing status"""
    return {
        "tracing_enabled": True,
        "storage_file": local_storage.storage_file,
        "opentelemetry_enabled": True,
        "anthropic_model": "claude-3-5-sonnet-20241022",
        "rag_enabled": True,
        "vector_store": "chromadb",
        "embedding_model": "all-MiniLM-L6-v2"
    }