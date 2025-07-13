"""
Comprehensive tracing setup for AI/LLM operations
"""

import os
import time
import json
import logging
from typing import Dict, Any, Optional, List
from datetime import datetime
from functools import wraps
from contextlib import contextmanager

from opentelemetry import trace
from opentelemetry.exporter.otlp.proto.grpc.trace_exporter import OTLPSpanExporter
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.sdk.trace.export import BatchSpanProcessor, ConsoleSpanExporter
from opentelemetry.sdk.resources import Resource
from opentelemetry.instrumentation.fastapi import FastAPIInstrumentor
from opentelemetry.instrumentation.requests import RequestsInstrumentor

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class LLMTracer:
    """Enhanced tracing for LLM operations"""
    
    def __init__(self, service_name: str = "spreadly-ai-service"):
        self.service_name = service_name
        self.tracer = None
        self.setup_tracing()
    
    def setup_tracing(self):
        """Initialize OpenTelemetry tracing"""
        
        # Create resource
        resource = Resource.create({
            "service.name": self.service_name,
            "service.version": "1.0.0",
            "deployment.environment": os.getenv("ENVIRONMENT", "development")
        })
        
        # Set up tracer provider
        trace.set_tracer_provider(TracerProvider(resource=resource))
        tracer_provider = trace.get_tracer_provider()
        
        # Add console exporter for development
        console_exporter = ConsoleSpanExporter()
        console_processor = BatchSpanProcessor(console_exporter)
        tracer_provider.add_span_processor(console_processor)
        
        # Add OTLP exporter if endpoint is configured
        otlp_endpoint = os.getenv("OTEL_EXPORTER_OTLP_ENDPOINT")
        if otlp_endpoint:
            logger.info(f"Setting up OTLP exporter to {otlp_endpoint}")
            otlp_exporter = OTLPSpanExporter(endpoint=otlp_endpoint)
            otlp_processor = BatchSpanProcessor(otlp_exporter)
            tracer_provider.add_span_processor(otlp_processor)
        
        # Get tracer
        self.tracer = trace.get_tracer(__name__)
        
        # Note: FastAPI instrumentation will be done in main.py
        # Instrument requests only here
        RequestsInstrumentor().instrument()
        
        logger.info("🔍 LLM Tracing initialized successfully")
    
    @contextmanager
    def trace_llm_call(
        self, 
        operation: str,
        model_name: str,
        provider: str = "anthropic",
        **kwargs
    ):
        """Context manager for tracing LLM API calls"""
        
        with self.tracer.start_as_current_span(
            f"llm.{operation}",
            attributes={
                "llm.provider": provider,
                "llm.model": model_name,
                "llm.operation": operation,
                **{f"llm.{k}": v for k, v in kwargs.items()}
            }
        ) as span:
            start_time = time.time()
            
            try:
                yield span
                span.set_status(trace.Status(trace.StatusCode.OK))
                
            except Exception as e:
                span.set_status(
                    trace.Status(
                        trace.StatusCode.ERROR, 
                        description=str(e)
                    )
                )
                span.set_attribute("error.type", type(e).__name__)
                span.set_attribute("error.message", str(e))
                raise
                
            finally:
                duration = time.time() - start_time
                span.set_attribute("llm.duration_seconds", duration)
                span.set_attribute("llm.timestamp", datetime.utcnow().isoformat())
    
    @contextmanager
    def trace_rag_operation(
        self, 
        operation: str,
        query: str,
        **kwargs
    ):
        """Context manager for tracing RAG operations"""
        
        with self.tracer.start_as_current_span(
            f"rag.{operation}",
            attributes={
                "rag.operation": operation,
                "rag.query": query[:100],  # Truncate long queries
                "rag.query_length": len(query),
                **{f"rag.{k}": v for k, v in kwargs.items()}
            }
        ) as span:
            start_time = time.time()
            
            try:
                yield span
                span.set_status(trace.Status(trace.StatusCode.OK))
                
            except Exception as e:
                span.set_status(
                    trace.Status(
                        trace.StatusCode.ERROR,
                        description=str(e)
                    )
                )
                span.set_attribute("error.type", type(e).__name__)
                span.set_attribute("error.message", str(e))
                raise
                
            finally:
                duration = time.time() - start_time
                span.set_attribute("rag.duration_seconds", duration)
    
    def trace_llm_metrics(
        self, 
        span,
        prompt_tokens: Optional[int] = None,
        completion_tokens: Optional[int] = None,
        total_tokens: Optional[int] = None,
        response_length: Optional[int] = None,
        **metrics
    ):
        """Add LLM-specific metrics to current span"""
        
        if prompt_tokens is not None:
            span.set_attribute("llm.usage.prompt_tokens", prompt_tokens)
        if completion_tokens is not None:
            span.set_attribute("llm.usage.completion_tokens", completion_tokens)
        if total_tokens is not None:
            span.set_attribute("llm.usage.total_tokens", total_tokens)
        if response_length is not None:
            span.set_attribute("llm.response.length", response_length)
        
        # Add custom metrics
        for key, value in metrics.items():
            span.set_attribute(f"llm.{key}", value)
    
    def trace_rag_metrics(
        self,
        span,
        num_retrieved: Optional[int] = None,
        similarity_scores: Optional[List[float]] = None,
        vector_store_status: Optional[str] = None,
        **metrics
    ):
        """Add RAG-specific metrics to current span"""
        
        if num_retrieved is not None:
            span.set_attribute("rag.retrieved_count", num_retrieved)
        if similarity_scores:
            span.set_attribute("rag.max_similarity", max(similarity_scores))
            span.set_attribute("rag.avg_similarity", sum(similarity_scores) / len(similarity_scores))
        if vector_store_status:
            span.set_attribute("rag.vector_store_status", vector_store_status)
        
        # Add custom metrics
        for key, value in metrics.items():
            span.set_attribute(f"rag.{key}", value)
    
    def log_trace_event(self, event_name: str, **attributes):
        """Log a custom trace event"""
        
        with self.tracer.start_as_current_span(event_name) as span:
            for key, value in attributes.items():
                span.set_attribute(key, value)
            
            logger.info(f"📊 Trace Event: {event_name}", extra=attributes)

class LocalTraceStorage:
    """Simple local storage for traces (development use)"""
    
    def __init__(self, storage_file: str = "traces.jsonl"):
        self.storage_file = storage_file
    
    def log_llm_call(
        self,
        operation: str,
        model: str,
        prompt: str,
        response: str,
        duration: float,
        tokens_used: Optional[int] = None,
        success: bool = True,
        error: Optional[str] = None,
        rag_used: bool = False,
        rag_models_retrieved: int = 0,
        rag_similarity_scores: Optional[List[float]] = None,
        **metadata
    ):
        """Log LLM call to local file"""
        
        trace_data = {
            "timestamp": datetime.utcnow().isoformat(),
            "operation": operation,
            "model": model,
            "prompt_preview": prompt[:200] + "..." if len(prompt) > 200 else prompt,
            "response_preview": response[:200] + "..." if len(response) > 200 else response,
            "prompt_length": len(prompt),
            "response_length": len(response),
            "duration_seconds": duration,
            "tokens_used": tokens_used,
            "success": success,
            "error": error,
            "rag_used": rag_used,
            "rag_models_retrieved": rag_models_retrieved,
            "rag_max_similarity": max(rag_similarity_scores) if rag_similarity_scores else None,
            "rag_avg_similarity": sum(rag_similarity_scores) / len(rag_similarity_scores) if rag_similarity_scores else None,
            **metadata
        }
        
        # Append to JSONL file
        with open(self.storage_file, "a") as f:
            f.write(json.dumps(trace_data) + "\n")
    
    def get_recent_traces(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Get recent traces from storage"""
        
        if not os.path.exists(self.storage_file):
            return []
        
        traces = []
        with open(self.storage_file, "r") as f:
            for line in f:
                try:
                    traces.append(json.loads(line.strip()))
                except json.JSONDecodeError:
                    continue
        
        return traces[-limit:]

# Global tracer instances
llm_tracer = LLMTracer()
local_storage = LocalTraceStorage()

# Decorator for automatic tracing
def trace_llm_operation(operation: str, model_name: str = None):
    """Decorator for tracing LLM operations"""
    
    def decorator(func):
        @wraps(func)
        async def async_wrapper(*args, **kwargs):
            model = model_name or getattr(args[0], 'model_name', 'unknown') if args else 'unknown'
            
            with llm_tracer.trace_llm_call(operation, model) as span:
                start_time = time.time()
                try:
                    result = await func(*args, **kwargs)
                    
                    # Log to local storage
                    local_storage.log_llm_call(
                        operation=operation,
                        model=model,
                        prompt=str(kwargs.get('prompt', 'N/A'))[:500],
                        response=str(result)[:500],
                        duration=time.time() - start_time,
                        success=True
                    )
                    
                    return result
                    
                except Exception as e:
                    local_storage.log_llm_call(
                        operation=operation,
                        model=model,
                        prompt=str(kwargs.get('prompt', 'N/A'))[:500],
                        response="",
                        duration=time.time() - start_time,
                        success=False,
                        error=str(e)
                    )
                    raise
        
        @wraps(func)
        def sync_wrapper(*args, **kwargs):
            model = model_name or getattr(args[0], 'model_name', 'unknown') if args else 'unknown'
            
            with llm_tracer.trace_llm_call(operation, model) as span:
                start_time = time.time()
                try:
                    result = func(*args, **kwargs)
                    
                    local_storage.log_llm_call(
                        operation=operation,
                        model=model,
                        prompt=str(kwargs.get('prompt', 'N/A'))[:500],
                        response=str(result)[:500],
                        duration=time.time() - start_time,
                        success=True
                    )
                    
                    return result
                    
                except Exception as e:
                    local_storage.log_llm_call(
                        operation=operation,
                        model=model,
                        prompt=str(kwargs.get('prompt', 'N/A'))[:500],
                        response="",
                        duration=time.time() - start_time,
                        success=False,
                        error=str(e)
                    )
                    raise
        
        # Return appropriate wrapper based on function type
        import asyncio
        if asyncio.iscoroutinefunction(func):
            return async_wrapper
        else:
            return sync_wrapper
    
    return decorator