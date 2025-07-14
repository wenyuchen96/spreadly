"""
Incremental Model Builder Service

Implements adaptive, chunk-based financial model generation with real-time
error recovery and intelligent optimization.
"""

from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum
import time
import json
from datetime import datetime

class ChunkComplexity(Enum):
    SIMPLE = "simple"      # Headers, basic data entry
    MEDIUM = "medium"      # Basic formulas, simple formatting  
    COMPLEX = "complex"    # Advanced formulas, complex logic
    CRITICAL = "critical"  # Key calculations, validation

class ChunkType(Enum):
    SETUP = "setup"           # Sheet creation, initial setup
    HEADERS = "headers"       # Column headers, labels
    DATA = "data"            # Static data entry
    FORMULAS = "formulas"    # Calculations, formulas
    FORMATTING = "formatting" # Colors, fonts, borders
    VALIDATION = "validation" # Data validation, checks
    FINALIZATION = "finalization" # Final touches, cleanup

class ExecutionStatus(Enum):
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    COMPLETED = "completed"
    FAILED = "failed"
    RETRYING = "retrying"

@dataclass
class CodeChunk:
    id: str
    type: ChunkType
    complexity: ChunkComplexity
    code: str
    description: str
    dependencies: List[str] = field(default_factory=list)
    stage: int = 0
    estimated_operations: int = 1
    max_retries: int = 3
    status: ExecutionStatus = ExecutionStatus.PENDING
    execution_attempts: int = 0
    error_history: List[str] = field(default_factory=list)
    execution_time: Optional[float] = None
    created_at: datetime = field(default_factory=datetime.now)

@dataclass 
class ModelBuildState:
    session_id: str
    financial_model_type: str  # Renamed to avoid Pydantic conflict
    total_chunks: int = 0
    completed_chunks: int = 0
    failed_chunks: int = 0
    current_chunk_id: Optional[str] = None
    chunks: Dict[str, CodeChunk] = field(default_factory=dict)
    execution_history: List[str] = field(default_factory=list)
    workbook_context: Dict[str, Any] = field(default_factory=dict)
    last_successful_context: Dict[str, Any] = field(default_factory=dict)
    error_patterns: List[str] = field(default_factory=list)
    started_at: datetime = field(default_factory=datetime.now)
    
    @property
    def progress_percentage(self) -> float:
        if self.total_chunks == 0:
            return 0.0
        return (self.completed_chunks / self.total_chunks) * 100
    
    @property
    def success_rate(self) -> float:
        total_attempts = self.completed_chunks + self.failed_chunks
        if total_attempts == 0:
            return 100.0
        return (self.completed_chunks / total_attempts) * 100

class ChunkGenerator:
    """Generates optimally-sized code chunks based on complexity analysis"""
    
    def __init__(self):
        self.complexity_patterns = {
            ChunkComplexity.SIMPLE: {
                'max_operations': 15,
                'keywords': ['getRange', 'values', 'basic'],
                'risk_score': 1
            },
            ChunkComplexity.MEDIUM: {
                'max_operations': 8,
                'keywords': ['formulas', 'format', 'SUM', 'AVERAGE'],
                'risk_score': 3
            },
            ChunkComplexity.COMPLEX: {
                'max_operations': 4,
                'keywords': ['IF', 'VLOOKUP', 'INDEX', 'MATCH', 'nested'],
                'risk_score': 7
            },
            ChunkComplexity.CRITICAL: {
                'max_operations': 2,
                'keywords': ['validation', 'error', 'critical', 'key'],
                'risk_score': 10
            }
        }
    
    def analyze_code_complexity(self, code: str) -> ChunkComplexity:
        """Analyze code to determine complexity level"""
        code_lower = code.lower()
        
        # Check for critical patterns first
        if any(keyword in code_lower for keyword in self.complexity_patterns[ChunkComplexity.CRITICAL]['keywords']):
            return ChunkComplexity.CRITICAL
        
        # Check for complex patterns
        if any(keyword in code_lower for keyword in self.complexity_patterns[ChunkComplexity.COMPLEX]['keywords']):
            return ChunkComplexity.COMPLEX
        
        # Check for medium patterns
        if any(keyword in code_lower for keyword in self.complexity_patterns[ChunkComplexity.MEDIUM]['keywords']):
            return ChunkComplexity.MEDIUM
        
        return ChunkComplexity.SIMPLE
    
    def determine_chunk_type(self, code: str, stage: int) -> ChunkType:
        """Determine the type of operation this chunk performs"""
        code_lower = code.lower()
        
        if stage == 0 or 'worksheets.add' in code_lower or 'getActiveWorksheet' in code_lower:
            return ChunkType.SETUP
        elif 'header' in code_lower or stage == 1:
            return ChunkType.HEADERS
        elif 'values' in code_lower and 'format' not in code_lower:
            return ChunkType.DATA
        elif 'formulas' in code_lower or '=' in code:
            return ChunkType.FORMULAS
        elif 'format' in code_lower or 'color' in code_lower or 'font' in code_lower:
            return ChunkType.FORMATTING
        elif 'validation' in code_lower or 'dataValidation' in code_lower:
            return ChunkType.VALIDATION
        else:
            return ChunkType.DATA  # Default fallback
    
    def estimate_operations(self, code: str) -> int:
        """Estimate the number of Excel operations in the code"""
        # Count potential Excel API calls
        operation_patterns = [
            'getRange',
            'values =',
            'formulas =', 
            'format.',
            'add(',
            'delete(',
            'insert(',
            'load(',
            'sync()'
        ]
        
        count = 0
        for pattern in operation_patterns:
            count += code.lower().count(pattern.lower())
        
        return max(1, count)  # At least 1 operation

class IncrementalModelBuilder:
    """Core service for incremental model building"""
    
    def __init__(self):
        self.chunk_generator = ChunkGenerator()
        self.active_sessions: Dict[str, ModelBuildState] = {}
        
    def start_incremental_build(
        self, 
        session_id: str, 
        model_type: str, 
        initial_query: str,
        workbook_context: Dict[str, Any]
    ) -> ModelBuildState:
        """Initialize a new incremental model building session"""
        
        # Create new build state
        build_state = ModelBuildState(
            session_id=session_id,
            financial_model_type=model_type,
            workbook_context=workbook_context,
            last_successful_context=workbook_context.copy()
        )
        
        # Store in active sessions
        self.active_sessions[session_id] = build_state
        
        return build_state
    
    async def generate_next_chunk(
        self, 
        session_id: str, 
        ai_service,
        current_context: Dict[str, Any] = None
    ) -> Optional[CodeChunk]:
        """Generate the next code chunk for execution"""
        
        if session_id not in self.active_sessions:
            raise ValueError(f"No active session found for {session_id}")
        
        build_state = self.active_sessions[session_id]
        
        # Update context if provided
        if current_context:
            build_state.workbook_context = current_context
        
        # Build context-aware prompt for next chunk
        chunk_prompt = self._build_chunk_prompt(build_state)
        
        # Generate actual chunk using AI service
        chunk_id = f"chunk_{len(build_state.chunks) + 1}"
        
        try:
            # Call AI service to generate the actual chunk
            print(f"🔧 Calling AI service for chunk generation...")
            chunk_code = await ai_service.generate_incremental_chunk(
                session_id=0,  # Placeholder session ID
                model_type=build_state.financial_model_type,
                build_context=chunk_prompt,
                workbook_context=build_state.workbook_context,
                previous_errors=build_state.error_patterns[-3:] if build_state.error_patterns else None
            )
            print(f"✅ AI generated chunk code ({len(chunk_code)} chars)")
        except Exception as e:
            print(f"❌ AI chunk generation failed, using fallback: {e}")
            import traceback
            traceback.print_exc()
            
            # Fallback to basic functional chunk
            chunk_code = f"""
await Excel.run(async (context) => {{
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Basic {build_state.financial_model_type} model step {len(build_state.chunks) + 1}
    sheet.getRange("A{len(build_state.chunks) + 1}").values = [["Step {len(build_state.chunks) + 1}"]];
    
    await context.sync();
}});
"""
        
        # Clean and validate the generated code
        from app.api.endpoints.incremental_model import clean_generated_code
        cleaned_chunk_code = clean_generated_code(chunk_code)
        
        chunk = CodeChunk(
            id=chunk_id,
            type=self.chunk_generator.determine_chunk_type(cleaned_chunk_code, len(build_state.chunks)),
            complexity=self.chunk_generator.analyze_code_complexity(cleaned_chunk_code),
            code=cleaned_chunk_code,
            description=f"Generated chunk {chunk_id} for {build_state.financial_model_type}",
            stage=len(build_state.chunks),
            estimated_operations=self.chunk_generator.estimate_operations(cleaned_chunk_code)
        )
        
        # Add to build state
        build_state.chunks[chunk_id] = chunk
        build_state.current_chunk_id = chunk_id
        build_state.total_chunks += 1
        
        return chunk
    
    def record_chunk_execution(
        self, 
        session_id: str, 
        chunk_id: str, 
        success: bool, 
        error_message: str = None,
        execution_time: float = None,
        new_context: Dict[str, Any] = None
    ) -> bool:
        """Record the result of chunk execution"""
        
        if session_id not in self.active_sessions:
            return False
        
        build_state = self.active_sessions[session_id]
        
        if chunk_id not in build_state.chunks:
            return False
        
        chunk = build_state.chunks[chunk_id]
        chunk.execution_attempts += 1
        chunk.execution_time = execution_time
        
        if success:
            chunk.status = ExecutionStatus.COMPLETED
            build_state.completed_chunks += 1
            build_state.execution_history.append(f"✅ {chunk_id}: {chunk.description}")
            
            # Update successful context
            if new_context:
                build_state.last_successful_context = new_context
                
        else:
            chunk.status = ExecutionStatus.FAILED
            build_state.failed_chunks += 1
            
            if error_message:
                chunk.error_history.append(error_message)
                build_state.error_patterns.append(error_message)
                build_state.execution_history.append(f"❌ {chunk_id}: {error_message}")
        
        return True
    
    def should_retry_chunk(self, session_id: str, chunk_id: str) -> bool:
        """Determine if a failed chunk should be retried"""
        
        if session_id not in self.active_sessions:
            return False
        
        build_state = self.active_sessions[session_id]
        
        if chunk_id not in build_state.chunks:
            return False
        
        chunk = build_state.chunks[chunk_id]
        
        return (chunk.status == ExecutionStatus.FAILED and 
                chunk.execution_attempts < chunk.max_retries)
    
    def is_build_complete(self, session_id: str) -> bool:
        """Check if the model building is complete"""
        
        if session_id not in self.active_sessions:
            return False
        
        build_state = self.active_sessions[session_id]
        
        # Build is complete when we have enough successful chunks for a complete DCF model
        if build_state.completed_chunks >= 25:  # Complete DCF model typically needs 25+ chunks
            print(f"✅ DCF model complete: {build_state.completed_chunks} chunks successfully executed")
            return True
        
        # Also complete if we've hit the formatting stage and have good progress
        current_stage = self._determine_build_stage(build_state.completed_chunks)
        if current_stage >= 6 and build_state.completed_chunks >= 20:
            print(f"✅ DCF model complete: Stage {current_stage} reached with {build_state.completed_chunks} chunks")
            return True
        
        # Stop if we've tried too many chunks without progress (stuck in loop)
        if build_state.total_chunks > 50:
            print(f"🛑 Stopping build: Too many chunks generated ({build_state.total_chunks})")
            return True
        
        # Check for repetitive failures
        if build_state.failed_chunks > 15:
            print(f"🛑 Stopping build: Too many failures ({build_state.failed_chunks})")
            return True
        
        return False
    
    def get_build_progress(self, session_id: str) -> Optional[Dict[str, Any]]:
        """Get current build progress and statistics"""
        
        if session_id not in self.active_sessions:
            return None
        
        build_state = self.active_sessions[session_id]
        
        return {
            'session_id': session_id,
            'model_type': build_state.financial_model_type,
            'progress_percentage': build_state.progress_percentage,
            'success_rate': build_state.success_rate,
            'total_chunks': build_state.total_chunks,
            'completed_chunks': build_state.completed_chunks,
            'failed_chunks': build_state.failed_chunks,
            'current_chunk_id': build_state.current_chunk_id,
            'execution_history': build_state.execution_history[-10:],  # Last 10 entries
            'error_patterns': list(set(build_state.error_patterns)),  # Unique errors
            'elapsed_time': (datetime.now() - build_state.started_at).total_seconds()
        }
    
    def _build_chunk_prompt(self, build_state: ModelBuildState) -> str:
        """Build a context-aware prompt for the next chunk generation"""
        
        # Determine current stage based on completed chunks
        total_completed = build_state.completed_chunks
        current_stage = self._determine_build_stage(total_completed)
        next_stage_description = self._get_next_stage_description(current_stage, build_state.financial_model_type)
        
        # Get completed chunk types to avoid repetition
        completed_types = []
        completed_descriptions = []
        for chunk in build_state.chunks.values():
            if chunk.status == ExecutionStatus.COMPLETED:
                completed_types.append(f"{chunk.type.value} ({chunk.complexity.value})")
                completed_descriptions.append(chunk.description)
        
        recent_completions = completed_descriptions[-3:] if completed_descriptions else ["None yet"]
        
        progress_context = f"""
        DCF MODEL BUILDING PROGRESS - STAGE {current_stage}/6
        
        Model Type: {build_state.financial_model_type.upper()}
        Progress: {build_state.completed_chunks} chunks completed successfully
        Success Rate: {build_state.success_rate:.1f}%
        
        COMPLETED STAGES:
        {chr(10).join(recent_completions)}
        
        CURRENT STAGE TARGET: {next_stage_description}
        
        AVOID REPEATING THESE TYPES:
        {', '.join(completed_types[-5:])}
        
        CURRENT WORKBOOK STATE:
        {self._format_workbook_context(build_state.workbook_context)}
        
        PROGRESSION REQUIREMENTS:
        1. DO NOT repeat similar chunks - move to the next logical stage
        2. Build a COMPLETE DCF model with: Assumptions → Revenue → Expenses → Cash Flow → Valuation
        3. Each chunk should advance the model construction
        4. Focus on {next_stage_description}
        
        PREVIOUS ERRORS TO AVOID:
        {chr(10).join(list(set(build_state.error_patterns))[-2:]) if build_state.error_patterns else "None"}
        """
        
        return progress_context
    
    def _determine_build_stage(self, completed_chunks: int) -> int:
        """Determine what stage of DCF building we're in"""
        if completed_chunks < 3:
            return 1  # Initial setup and headers
        elif completed_chunks < 8:
            return 2  # Assumptions section
        elif completed_chunks < 15:
            return 3  # Revenue projections
        elif completed_chunks < 22:
            return 4  # Expenses and cash flow
        elif completed_chunks < 28:
            return 5  # Valuation calculations
        else:
            return 6  # Formatting and finalization
    
    def _get_next_stage_description(self, stage: int, model_type: str) -> str:
        """Get description of what should be built in the next stage"""
        stage_descriptions = {
            1: "Create main model headers and initial setup",
            2: "Build detailed assumptions section with input cells", 
            3: "Add revenue projections and growth calculations",
            4: "Implement operating expenses and cash flow calculations",
            5: "Create DCF valuation formulas (NPV, terminal value)",
            6: "Apply professional formatting and final touches"
        }
        
        return stage_descriptions.get(stage, "Complete the DCF model")
    
    def _format_workbook_context(self, context: Dict[str, Any]) -> str:
        """Format workbook context for better readability"""
        if not context or not context.get('sheets'):
            return "Empty workbook"
        
        sheet_info = []
        for sheet in context.get('sheets', [])[:2]:  # Max 2 sheets
            name = sheet.get('name', 'Unknown')
            data = sheet.get('data', [])
            if data and len(data) > 0:
                row_count = len([row for row in data if row and any(cell for cell in row if cell)])
                sheet_info.append(f"Sheet '{name}': {row_count} rows with content")
            else:
                sheet_info.append(f"Sheet '{name}': empty")
        
        return '; '.join(sheet_info)
    
    def cleanup_session(self, session_id: str) -> bool:
        """Clean up completed or abandoned sessions"""
        
        if session_id in self.active_sessions:
            del self.active_sessions[session_id]
            return True
        
        return False

# Global instance
incremental_builder = IncrementalModelBuilder()