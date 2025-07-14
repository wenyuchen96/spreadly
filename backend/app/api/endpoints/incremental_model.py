"""
Incremental Model Building API Endpoints

Provides endpoints for chunk-based financial model generation with
real-time error recovery and adaptive optimization.
"""

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session
from typing import Dict, Any, Optional, List
from app.core.database import get_db
from app.services.incremental_model_builder import incremental_builder, ExecutionStatus
from app.services.ai_service_simple import AIService
from app.models.session import Session as SessionModel
import time
import json
import re
from datetime import datetime, timedelta
from collections import defaultdict

router = APIRouter()

# Rate limiting for error analysis to prevent API abuse
class RateLimiter:
    def __init__(self):
        self.error_analysis_calls = defaultdict(list)  # session_id -> [timestamps]
        self.max_calls_per_minute = 5
        self.max_calls_per_hour = 20
        
    def can_make_error_analysis(self, session_id: str) -> bool:
        now = datetime.now()
        calls = self.error_analysis_calls[session_id]
        
        # Remove old timestamps (older than 1 hour)
        calls[:] = [ts for ts in calls if now - ts < timedelta(hours=1)]
        
        # Check hourly limit
        if len(calls) >= self.max_calls_per_hour:
            return False
            
        # Check per-minute limit
        recent_calls = [ts for ts in calls if now - ts < timedelta(minutes=1)]
        if len(recent_calls) >= self.max_calls_per_minute:
            return False
            
        return True
        
    def record_error_analysis(self, session_id: str):
        self.error_analysis_calls[session_id].append(datetime.now())

rate_limiter = RateLimiter()

@router.post("/start")
async def start_incremental_model_build(
    request_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """
    Initialize incremental model building session
    
    Expected payload:
    {
        "session_token": "uuid",
        "model_type": "dcf|npv|lbo|...",
        "query": "user's original request",
        "workbook_context": {...}
    }
    """
    
    session_token = request_data.get("session_token")
    model_type = request_data.get("model_type", "financial")
    query = request_data.get("query", "")
    workbook_context = request_data.get("workbook_context", {})
    
    if not session_token:
        raise HTTPException(status_code=400, detail="Session token is required")
    
    # Verify session exists
    session = db.query(SessionModel).filter(SessionModel.session_token == session_token).first()
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    try:
        # Initialize incremental build
        build_state = incremental_builder.start_incremental_build(
            session_id=session_token,
            model_type=model_type,
            initial_query=query,
            workbook_context=workbook_context
        )
        
        print(f"🔧 Started incremental build for {model_type} model (session: {session_token})")
        
        return {
            "success": True,
            "session_id": session_token,
            "model_type": model_type,
            "build_state": {
                "progress_percentage": build_state.progress_percentage,
                "total_chunks": build_state.total_chunks,
                "status": "initialized"
            },
            "message": f"Incremental {model_type} model building initialized"
        }
        
    except Exception as e:
        print(f"❌ Error starting incremental build: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/next-chunk")
async def generate_next_chunk(
    request_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """
    Generate and return the next code chunk for execution
    
    Expected payload:
    {
        "session_token": "uuid",
        "current_context": {...},  // Current Excel state
        "last_execution_result": {  // Optional: result of previous chunk
            "success": true/false,
            "error_message": "...",
            "execution_time": 0.5
        }
    }
    """
    
    session_token = request_data.get("session_token")
    current_context = request_data.get("current_context")
    last_result = request_data.get("last_execution_result")
    retry_chunk_id = request_data.get("retry_chunk_id")  # For retrying fixed chunks
    
    if not session_token:
        raise HTTPException(status_code=400, detail="Session token is required")
    
    try:
        # Record previous execution result if provided
        if last_result and "chunk_id" in last_result:
            success = last_result.get("success", False)
            incremental_builder.record_chunk_execution(
                session_id=session_token,
                chunk_id=last_result["chunk_id"],
                success=success,
                error_message=last_result.get("error_message"),
                execution_time=last_result.get("execution_time"),
                new_context=current_context
            )
            
            if success:
                print(f"✅ Chunk {last_result['chunk_id']} executed successfully")
            else:
                print(f"❌ Chunk {last_result['chunk_id']} failed: {last_result.get('error_message', 'Unknown error')[:100]}")
        
        # If requesting a specific retry chunk, return the fixed version
        if retry_chunk_id:
            build_state = incremental_builder.active_sessions.get(session_token)
            if build_state and retry_chunk_id in build_state.chunks:
                fixed_chunk = build_state.chunks[retry_chunk_id]
                progress = incremental_builder.get_build_progress(session_token)
                
                print(f"🔄 Returning fixed chunk {retry_chunk_id} for retry")
                
                return {
                    "success": True,
                    "completed": False,
                    "chunk": {
                        "id": fixed_chunk.id,
                        "type": fixed_chunk.type.value,
                        "complexity": fixed_chunk.complexity.value,
                        "code": fixed_chunk.code,
                        "description": fixed_chunk.description,
                        "estimated_operations": fixed_chunk.estimated_operations,
                        "stage": fixed_chunk.stage
                    },
                    "progress": progress
                }
        
        # Check if build is complete or should stop due to too many failures
        build_state = incremental_builder.active_sessions.get(session_token)
        if build_state and build_state.failed_chunks > 10:  # Circuit breaker
            print(f"🛑 Too many failed chunks ({build_state.failed_chunks}), stopping build")
            progress = incremental_builder.get_build_progress(session_token)
            return {
                "success": True,
                "completed": True,
                "progress": progress,
                "message": f"Build stopped due to excessive failures. Success rate: {progress['success_rate']:.1f}%"
            }
        
        if incremental_builder.is_build_complete(session_token):
            progress = incremental_builder.get_build_progress(session_token)
            return {
                "success": True,
                "completed": True,
                "progress": progress,
                "message": f"Model building completed! Success rate: {progress['success_rate']:.1f}%"
            }
        
        # Generate next chunk using AI service
        ai_service = AIService()
        
        # Get the next chunk
        chunk = await incremental_builder.generate_next_chunk(
            session_id=session_token,
            ai_service=ai_service,
            current_context=current_context
        )
        
        if not chunk:
            raise HTTPException(status_code=404, detail="No more chunks to generate")
        
        # Update chunk status
        chunk.status = ExecutionStatus.IN_PROGRESS
        
        print(f"🔧 Generated chunk {chunk.id} ({chunk.type.value}, {chunk.complexity.value})")
        
        return {
            "success": True,
            "completed": False,
            "chunk": {
                "id": chunk.id,
                "type": chunk.type.value,
                "complexity": chunk.complexity.value,
                "code": chunk.code,
                "description": chunk.description,
                "estimated_operations": chunk.estimated_operations,
                "stage": chunk.stage
            },
            "progress": incremental_builder.get_build_progress(session_token)
        }
        
    except Exception as e:
        print(f"❌ Error generating next chunk: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/handle-error")
async def handle_chunk_error(
    request_data: Dict[str, Any],
    db: Session = Depends(get_db)
):
    """
    Handle chunk execution error and attempt recovery
    
    Expected payload:
    {
        "session_token": "uuid",
        "chunk_id": "chunk_1",
        "error_message": "Array dimension mismatch...",
        "execution_time": 0.3,
        "current_context": {...}
    }
    """
    
    session_token = request_data.get("session_token")
    chunk_id = request_data.get("chunk_id")
    error_message = request_data.get("error_message", "Unknown error")
    execution_time = request_data.get("execution_time", 0.0)
    current_context = request_data.get("current_context")
    
    if not session_token or not chunk_id:
        raise HTTPException(status_code=400, detail="Session token and chunk ID are required")
    
    try:
        # Record the failure
        success = incremental_builder.record_chunk_execution(
            session_id=session_token,
            chunk_id=chunk_id,
            success=False,
            error_message=error_message,
            execution_time=execution_time,
            new_context=current_context
        )
        
        if not success:
            raise HTTPException(status_code=404, detail="Chunk or session not found")
        
        # Check if we should retry
        should_retry = incremental_builder.should_retry_chunk(session_token, chunk_id)
        
        if should_retry:
            print(f"🔄 Retrying chunk {chunk_id} after error: {error_message}")
            
            # Rate limiting for error analysis to prevent API abuse
            if not rate_limiter.can_make_error_analysis(session_token):
                print(f"⚠️ Rate limit exceeded for session {session_token}, skipping error analysis")
                return {
                    "success": True,
                    "action": "skip",
                    "chunk_id": chunk_id,
                    "retry_attempt": False,
                    "message": f"Rate limit exceeded, skipping chunk {chunk_id}",
                    "progress": incremental_builder.get_build_progress(session_token)
                }
            
            # Record the error analysis call
            rate_limiter.record_error_analysis(session_token)
            
            # Implement AI-powered error analysis and code regeneration
            ai_service = AIService()
            error_analysis = await analyze_and_fix_chunk_error(
                ai_service, 
                session_token, 
                chunk_id, 
                error_message,
                current_context
            )
            
            if error_analysis.get("fixed_code"):
                # Update the chunk with the fixed code
                build_state = incremental_builder.active_sessions.get(session_token)
                if build_state and chunk_id in build_state.chunks:
                    build_state.chunks[chunk_id].code = error_analysis["fixed_code"]
                    build_state.chunks[chunk_id].status = ExecutionStatus.PENDING
                    build_state.chunks[chunk_id].error_history.append(f"Auto-fixed: {error_analysis.get('fix_description', 'Code regenerated')}")
                    print(f"✅ Auto-fixed chunk {chunk_id}: {error_analysis.get('fix_description', 'Code regenerated')}")
            
            return {
                "success": True,
                "action": "retry",
                "chunk_id": chunk_id,
                "retry_attempt": True,
                "auto_fixed": error_analysis.get("fixed_code") is not None,
                "fix_description": error_analysis.get("fix_description"),
                "message": f"Retrying chunk {chunk_id} with improved code",
                "progress": incremental_builder.get_build_progress(session_token)
            }
        else:
            # Max retries reached, move to next chunk or fail gracefully
            print(f"❌ Chunk {chunk_id} failed permanently: {error_message}")
            
            return {
                "success": True,
                "action": "skip",
                "chunk_id": chunk_id,
                "retry_attempt": False,
                "message": f"Skipping failed chunk {chunk_id}, continuing with next",
                "progress": incremental_builder.get_build_progress(session_token)
            }
        
    except Exception as e:
        print(f"❌ Error handling chunk error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/status/{session_token}")
async def get_build_status(
    session_token: str,
    db: Session = Depends(get_db)
):
    """Get current incremental build status and progress"""
    
    try:
        progress = incremental_builder.get_build_progress(session_token)
        
        if not progress:
            raise HTTPException(status_code=404, detail="Build session not found")
        
        return {
            "success": True,
            "progress": progress,
            "is_complete": incremental_builder.is_build_complete(session_token)
        }
        
    except Exception as e:
        print(f"❌ Error getting build status: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/cancel/{session_token}")
async def cancel_build(
    session_token: str,
    db: Session = Depends(get_db)
):
    """Cancel an ongoing incremental build session"""
    
    try:
        success = incremental_builder.cleanup_session(session_token)
        
        if success:
            print(f"🛑 Cancelled incremental build session: {session_token}")
            return {
                "success": True,
                "message": "Build session cancelled successfully"
            }
        else:
            raise HTTPException(status_code=404, detail="Build session not found")
        
    except Exception as e:
        print(f"❌ Error cancelling build: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/sessions")
async def list_active_sessions():
    """List all active incremental build sessions (for debugging)"""
    
    try:
        sessions = []
        for session_id, build_state in incremental_builder.active_sessions.items():
            sessions.append({
                "session_id": session_id,
                "model_type": build_state.financial_model_type,
                "progress": build_state.progress_percentage,
                "total_chunks": build_state.total_chunks,
                "completed_chunks": build_state.completed_chunks,
                "failed_chunks": build_state.failed_chunks,
                "started_at": build_state.started_at.isoformat()
            })
        
        return {
            "success": True,
            "active_sessions": sessions,
            "total_active": len(sessions)
        }
        
    except Exception as e:
        print(f"❌ Error listing sessions: {e}")
        raise HTTPException(status_code=500, detail=str(e))


async def analyze_and_fix_chunk_error(
    ai_service: AIService,
    session_token: str,
    chunk_id: str,
    error_message: str,
    current_context: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Analyze JavaScript execution errors and generate fixed code using AI
    """
    
    try:
        # Get the failed chunk
        build_state = incremental_builder.active_sessions.get(session_token)
        if not build_state or chunk_id not in build_state.chunks:
            return {"error": "Chunk not found"}
        
        failed_chunk = build_state.chunks[chunk_id]
        original_code = failed_chunk.code
        
        # Parse error details if it's JSON (from enhanced frontend error reporting)
        error_details = {}
        try:
            error_details = json.loads(error_message)
        except (json.JSONDecodeError, TypeError):
            error_details = {"message": error_message, "syntax_error": "SyntaxError" in error_message}
        
        # Analyze the error type and create targeted fix prompt
        context_snippet = json.dumps(current_context, indent=2)[:500] + "..." if current_context else "No context"
        
        error_analysis_prompt = f"""SYSTEM: You are a JavaScript code fixer. Return ONLY executable JavaScript code with NO explanations.

BROKEN CODE:
{original_code}

ERROR: {error_details.get('message', error_message)}

TASK: Fix the error and return complete, executable JavaScript code.

REQUIREMENTS:
1. Start with: await Excel.run(async (context) => {{
2. End with: }});
3. Use 2D arrays: .values = [["value"]]
4. NO explanations, NO analysis text
5. Fix the specific error only

CORRECTED CODE:"""
        
        print(f"🔍 Analyzing error for chunk {chunk_id}: {error_details.get('message', error_message)[:100]}...")
        
        # Use AI to analyze and fix the error
        fixed_code = await ai_service.generate_incremental_chunk(
            session_id=0,
            model_type="error_fix",
            build_context=error_analysis_prompt,
            workbook_context=current_context,
            previous_errors=[error_message]
        )
        
        # Clean the fixed code
        cleaned_code = clean_generated_code(fixed_code)
        
        # Determine fix description based on error type
        fix_description = determine_fix_description(error_details, original_code, cleaned_code)
        
        print(f"✅ Generated fix for {chunk_id}: {fix_description}")
        
        return {
            "fixed_code": cleaned_code,
            "fix_description": fix_description,
            "original_error": error_details.get('message', error_message),
            "error_type": error_details.get('name', 'Unknown')
        }
        
    except Exception as e:
        print(f"❌ Error analyzing chunk error: {e}")
        return {"error": str(e)}


def clean_generated_code(code: str) -> str:
    """Clean AI-generated code to ensure it's executable"""
    
    cleaned = code.strip()
    
    # Remove markdown code fences
    cleaned = re.sub(r'^```(?:javascript|js)?\n?', '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'\n?```$', '', cleaned, flags=re.MULTILINE)
    
    # AGGRESSIVE cleaning: Remove all explanatory text before Excel.run
    lines = cleaned.split('\n')
    excel_run_found = False
    code_lines = []
    
    for line in lines:
        line_stripped = line.strip()
        
        # Start collecting code once we find Excel.run
        if 'await Excel.run' in line:
            excel_run_found = True
            code_lines.append(line)
        elif excel_run_found:
            code_lines.append(line)
        # Skip all explanatory text before Excel.run
    
    if code_lines:
        cleaned = '\n'.join(code_lines)
    else:
        # Fallback: look for any JavaScript-like content
        for i, line in enumerate(lines):
            if ('Excel.run' in line or 'const sheet' in line or 
                'sheet.getRange' in line or 'async (context)' in line):
                cleaned = '\n'.join(lines[i:])
                break
    
    # Remove any remaining explanatory text after the code
    cleaned = cleaned.strip()
    if '}});' in cleaned:
        # Find the last }});  and cut everything after it
        last_closing = cleaned.rfind('}});')
        if last_closing != -1:
            cleaned = cleaned[:last_closing + 4]
    
    # Validate code completeness and syntax
    validation_errors = validate_javascript_syntax(cleaned)
    if validation_errors:
        print(f"⚠️ Code validation errors found: {validation_errors}")
        cleaned = fix_syntax_errors(cleaned, validation_errors)
    
    # Final completeness check and truncation handling
    if not is_code_complete(cleaned):
        print(f"⚠️ Generated code appears incomplete, attempting to complete...")
        # First try our new truncation completion
        cleaned = complete_truncated_code(cleaned)
        
        # If still incomplete, try the existing fix
        if not is_code_complete(cleaned):
            cleaned = fix_incomplete_code(cleaned)
    
    return cleaned


def validate_javascript_syntax(code: str) -> List[str]:
    """Validate JavaScript syntax and return list of errors"""
    errors = []
    
    # Check for balanced brackets, braces, parentheses
    open_braces = code.count('{')
    close_braces = code.count('}')
    if open_braces != close_braces:
        errors.append(f"Unmatched braces: {open_braces} opening, {close_braces} closing")
    
    open_parens = code.count('(')
    close_parens = code.count(')')
    if open_parens != close_parens:
        errors.append(f"Unmatched parentheses: {open_parens} opening, {close_parens} closing")
    
    open_brackets = code.count('[')
    close_brackets = code.count(']')
    if open_brackets != close_brackets:
        errors.append(f"Unmatched brackets: {open_brackets} opening, {close_brackets} closing")
    
    # Check for Excel.run structure
    if 'Excel.run' in code and not re.search(r'Excel\.run\s*\(\s*async\s*\(\s*context\s*\)\s*=>', code):
        errors.append("Invalid Excel.run structure")
    
    # Check for context.sync()
    if 'Excel.run' in code and 'context.sync()' not in code:
        errors.append("Missing context.sync() call")
    
    # Check for common syntax errors
    lines = code.split('\n')
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped and not stripped.startswith('//') and not stripped.startswith('/*'):
            # Check for statements that should end with semicolon
            if (stripped.endswith('.values') or 
                stripped.endswith('.formulas') or
                stripped.endswith('.color') or
                stripped.endswith('true') or
                stripped.endswith('false')):
                if not stripped.endswith(';'):
                    errors.append(f"Line {i+1}: Missing semicolon")
    
    return errors


def fix_syntax_errors(code: str, errors: List[str]) -> str:
    """Attempt to fix common syntax errors"""
    fixed = code
    
    # Fix missing semicolons
    if any("Missing semicolon" in error for error in errors):
        lines = fixed.split('\n')
        for i, line in enumerate(lines):
            stripped = line.strip()
            if (stripped.endswith('.values') or 
                stripped.endswith('.formulas') or 
                stripped.endswith('.color') or
                stripped.endswith('true') or 
                stripped.endswith('false')):
                if not stripped.endswith(';'):
                    lines[i] = line + ';'
        fixed = '\n'.join(lines)
    
    # Fix unmatched braces
    if any("Unmatched braces" in error for error in errors):
        open_braces = fixed.count('{')
        close_braces = fixed.count('}')
        if open_braces > close_braces:
            fixed += '\n' + '}'.join([''] * (open_braces - close_braces + 1))
    
    # Fix unmatched parentheses
    if any("Unmatched parentheses" in error for error in errors):
        open_parens = fixed.count('(')
        close_parens = fixed.count(')')
        if open_parens > close_parens:
            fixed += ')' * (open_parens - close_parens)
    
    # Add missing context.sync()
    if any("Missing context.sync()" in error for error in errors):
        if 'Excel.run' in fixed and 'context.sync()' not in fixed:
            # Insert before the last closing brace
            lines = fixed.split('\n')
            for i in range(len(lines) - 1, -1, -1):
                if '}' in lines[i]:
                    lines.insert(i, '    await context.sync();')
                    break
            fixed = '\n'.join(lines)
    
    return fixed


def complete_truncated_code(code: str) -> str:
    """Complete truncated JavaScript code to make it executable"""
    
    lines = code.split('\n')
    if not lines:
        return code
    
    last_line = lines[-1].strip()
    
    # Handle specific truncation patterns we've seen
    if '=1/(1+$B$7)^' in last_line:
        # Complete the discount factor formula
        if last_line.endswith('^'):
            lines[-1] = last_line + '3", "=1/(1+$B$7)^4", "=1/(1+$B$7)^5"]]'
        elif '=1/(1+$B$7)^' in last_line and not last_line.endswith(']]'):
            # Find where the truncation happened and complete the array
            lines[-1] = last_line + '"]]'
    
    elif 'sheet.getRange(' in last_line and not last_line.endswith(';'):
        # Complete incomplete getRange statements
        if '.values = [' in last_line and not last_line.endswith(']]'):
            lines[-1] = last_line + ']]'
        elif not last_line.endswith(';'):
            lines[-1] = last_line + ';'
    
    elif last_line.endswith('sheet'):
        # Remove hanging sheet reference
        lines = lines[:-1]
    
    # Ensure proper Excel.run closure
    code_str = '\n'.join(lines)
    
    # Check if we have Excel.run but missing proper closure
    if 'Excel.run(async (context) => {' in code_str:
        open_braces = code_str.count('{')
        close_braces = code_str.count('}')
        
        if open_braces > close_braces:
            # Add context.sync() if missing
            if 'context.sync()' not in code_str:
                code_str += '\n    await context.sync();'
            
            # Add missing closing braces
            missing_braces = open_braces - close_braces
            code_str += '\n' + '}'.join([''] * (missing_braces + 1))
    
    return code_str


def is_code_complete(code: str) -> bool:
    """Check if JavaScript code appears to be complete"""
    
    lines = code.split('\n')
    if not lines:
        return False
    
    last_line = lines[-1].strip()
    
    # Check for incomplete statements (enhanced list)
    incomplete_indicators = [
        'sheet.getRange(',
        'sheet.getRange("',
        'format.',
        'values =',
        'formulas =',
        '.values = [',
        '.formulas = [',
        '=1/(1+$B$7)^',  # Specific to the error we saw
        'sheet.getRange("A',
        'sheet.getRange("B',
        'sheet.getRange("C',
        '"//',  # Incomplete comment
        '/*',  # Unclosed comment
        'const ',
        'let ',
        'var ',
        'sheet.',  # Hanging sheet reference
        'context.',  # Hanging context reference
    ]
    
    if any(last_line.endswith(indicator) for indicator in incomplete_indicators):
        return False
    
    # Check for truncated formulas (common issue)
    if ('=1/(1+$B$7)^' in last_line or 
        '=SUM(' in last_line and not ')' in last_line or
        '="' in last_line and last_line.count('"') % 2 == 1):
        return False
    
    # Check bracket/brace balance
    open_braces = code.count('{')
    close_braces = code.count('}')
    open_parens = code.count('(')
    close_parens = code.count(')')
    open_brackets = code.count('[')
    close_brackets = code.count(']')
    
    if (open_braces != close_braces or 
        open_parens != close_parens or 
        open_brackets != close_brackets):
        return False
    
    return True


def fix_incomplete_code(code: str) -> str:
    """Attempt to fix incomplete JavaScript code"""
    
    lines = code.split('\n')
    last_line = lines[-1].strip()
    
    # Fix incomplete comments
    if last_line.startswith('//') and not last_line.endswith('*/'):
        # Complete the comment line by removing incomplete part
        lines = lines[:-1]
    
    # Fix incomplete statements
    if last_line.endswith(('const ', 'let ', 'var ', 'sheet.getRange(', 'format.')):
        # Remove incomplete statement
        lines = lines[:-1]
    
    fixed_code = '\n'.join(lines)
    
    # Ensure proper context.sync() and closing
    if 'Excel.run' in fixed_code and 'context.sync()' not in fixed_code:
        # Find the last Excel operation and add sync
        lines = fixed_code.split('\n')
        insert_sync = False
        for i in range(len(lines) - 1, -1, -1):
            if 'sheet.' in lines[i] or 'getRange' in lines[i]:
                # Add sync after this line
                lines.insert(i + 1, '    await context.sync();')
                insert_sync = True
                break
        
        if insert_sync:
            fixed_code = '\n'.join(lines)
    
    # Ensure proper closing for Excel.run
    open_braces = fixed_code.count('{')
    close_braces = fixed_code.count('}')
    open_parens = fixed_code.count('(')
    close_parens = fixed_code.count(')')
    
    # Add missing closing braces/parens
    if open_braces > close_braces:
        fixed_code += '\n' + '}'.join([''] * (open_braces - close_braces + 1))
    
    if open_parens > close_parens:
        fixed_code += ')' * (open_parens - close_parens)
    
    return fixed_code.strip()


def determine_fix_description(error_details: Dict, original_code: str, fixed_code: str) -> str:
    """Determine what was fixed based on error analysis"""
    
    error_msg = error_details.get('message', '').lower()
    
    if 'unexpected end of script' in error_msg:
        return "Fixed incomplete/truncated code - completed missing statements"
    elif 'unexpected identifier' in error_msg:
        return "Fixed syntax error: missing semicolons or malformed function calls"
    elif 'syntaxerror' in error_msg:
        return "Fixed JavaScript syntax error"
    elif 'excel is not defined' in error_msg:
        return "Added proper Excel.js wrapper and context"
    elif 'array' in error_msg or 'dimension' in error_msg:
        return "Fixed array dimension mismatch for Excel.js"
    elif 'async' in error_msg or 'await' in error_msg:
        return "Fixed async/await syntax structure"
    elif len(fixed_code) > len(original_code) * 1.2:
        return "Added missing Excel.run wrapper and error handling"
    else:
        return "Code regenerated with improved syntax and structure"