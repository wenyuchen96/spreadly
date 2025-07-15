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
    token_usage: Dict[str, int] = field(default_factory=dict)
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
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    
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
            print(f"🔍 Updated workbook context - type: {type(current_context)}")
            print(f"🔍 Raw context keys: {list(current_context.keys()) if isinstance(current_context, dict) else 'Not a dict'}")
            
            if isinstance(current_context, dict) and 'sheets' in current_context:
                sheets = current_context.get('sheets', [])
                print(f"🔍 Context has {len(sheets)} sheets")
                
                for i, sheet in enumerate(sheets[:2]):  # Log first 2 sheets
                    if isinstance(sheet, dict):
                        sheet_name = sheet.get('name', 'Unknown')
                        data = sheet.get('data', [])
                        
                        print(f"🔍 Sheet {i+1}: '{sheet_name}'")
                        print(f"🔍   - data type: {type(data)}")
                        print(f"🔍   - data length: {len(data) if data is not None else 'None'}")
                        print(f"🔍   - data is empty list: {data == []}")
                        print(f"🔍   - data is None: {data is None}")
                        
                        if data and len(data) > 0:
                            print(f"🔍   - first row: {data[0]}")
                            print(f"🔍   - sample data structure: {data[:2]}")
                        else:
                            print(f"🔍   - NO DATA FOUND - this is the core issue!")
                            # Let's see what other fields the sheet has
                            print(f"🔍   - sheet keys: {list(sheet.keys())}")
                            if 'usedRange' in sheet and sheet['usedRange']:
                                print(f"🔍   - usedRange: {sheet['usedRange']}")
                    else:
                        print(f"🔍 Sheet {i+1}: Not a dict - {type(sheet)} = {sheet}")
            else:
                print(f"🔍 Context format (no sheets): {current_context}")
        else:
            print("🔍 No current_context provided to next-chunk")
        
        # Build context-aware prompt for next chunk
        chunk_prompt = self._build_chunk_prompt(build_state)
        
        # Generate actual chunk using AI service
        chunk_id = f"chunk_{len(build_state.chunks) + 1}"
        
        try:
            # Call AI service to generate the actual chunk
            print(f"🔧 Calling AI service for chunk generation...")
            chunk_result = await ai_service.generate_incremental_chunk(
                session_id=0,  # Placeholder session ID
                model_type=build_state.financial_model_type,
                build_context=chunk_prompt,
                workbook_context=build_state.workbook_context,
                previous_errors=build_state.error_patterns[-3:] if build_state.error_patterns else None
            )
            
            # Extract code and token information
            chunk_code = chunk_result.get("code", "") if isinstance(chunk_result, dict) else chunk_result
            token_usage = chunk_result.get("token_usage", {}) if isinstance(chunk_result, dict) else {}
            print(f"✅ AI generated chunk code ({len(chunk_code)} chars)")
            if token_usage:
                print(f"🔢 Token usage: {token_usage.get('input_tokens', 0)} input + {token_usage.get('output_tokens', 0)} output = {token_usage.get('total_tokens', 0)} total")
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
            estimated_operations=self.chunk_generator.estimate_operations(cleaned_chunk_code),
            token_usage=token_usage
        )
        
        # Update build state with token information
        if token_usage:
            build_state.total_input_tokens += token_usage.get('input_tokens', 0)
            build_state.total_output_tokens += token_usage.get('output_tokens', 0)
        
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
        current_stage = self._determine_build_stage(build_state.completed_chunks, build_state.financial_model_type)
        model_lower = build_state.financial_model_type.lower()
        
        # Different completion criteria based on model type
        if 'three' in model_lower or '3' in model_lower or 'integrated' in model_lower:
            # Three-statement models need more chunks
            completion_stage = 8
            min_chunks = 25
        elif 'dcf' in model_lower or 'discounted' in model_lower:
            # DCF models now have 9 stages
            completion_stage = 8
            min_chunks = 20
        else:
            # Other models
            completion_stage = 6
            min_chunks = 15
            
        if current_stage >= completion_stage and build_state.completed_chunks >= min_chunks:
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
            'elapsed_time': (datetime.now() - build_state.started_at).total_seconds(),
            'token_usage': {
                'input_tokens': build_state.total_input_tokens,
                'output_tokens': build_state.total_output_tokens,
                'total_tokens': build_state.total_input_tokens + build_state.total_output_tokens
            }
        }
    
    def _build_chunk_prompt(self, build_state: ModelBuildState) -> str:
        """Build a context-aware prompt for the next chunk generation"""
        
        # Determine current stage based on completed chunks
        total_completed = build_state.completed_chunks
        current_stage = self._determine_build_stage(total_completed, build_state.financial_model_type)
        next_stage_description = self._get_next_stage_description(current_stage, build_state.financial_model_type)
        
        # Get completed chunk types to avoid repetition
        completed_types = []
        completed_descriptions = []
        for chunk in build_state.chunks.values():
            if chunk.status == ExecutionStatus.COMPLETED:
                completed_types.append(f"{chunk.type.value} ({chunk.complexity.value})")
                completed_descriptions.append(chunk.description)
        
        recent_completions = completed_descriptions[-3:] if completed_descriptions else ["None yet"]
        
        # Extract content placement guidance from workbook context
        placement_guidance = self._extract_placement_guidance(build_state.workbook_context)
        
        # Determine total stages based on model type
        model_lower = build_state.financial_model_type.lower()
        if 'three' in model_lower or '3' in model_lower or 'integrated' in model_lower:
            total_stages = 9
            progression_requirements = """PROGRESSION REQUIREMENTS:
        1. DO NOT repeat similar chunks - move to the next logical stage
        2. Build a COMPLETE THREE-STATEMENT MODEL with: Assumptions → Income Statement → Balance Sheet → Cash Flow → Integration
        3. Each chunk should advance the model construction
        4. Ensure proper statement integration (Net Income → RE, Cash flow ties, etc.)"""
        elif 'dcf' in model_lower or 'discounted' in model_lower:
            total_stages = 9
            progression_requirements = """PROGRESSION REQUIREMENTS:
        1. DO NOT repeat similar chunks - move to the next logical stage
        2. Build a COMPLETE DCF MODEL with: Assumptions → P&L → Working Capital → Free Cash Flow → Terminal Value → Valuation → WACC → Sensitivity
        3. Each chunk should advance the DCF model construction
        4. Focus on proper DCF methodology (EBIAT, UFCF, WACC, Terminal Value)"""
        else:
            total_stages = 6
            progression_requirements = """PROGRESSION REQUIREMENTS:
        1. DO NOT repeat similar chunks - move to the next logical stage
        2. Build a COMPLETE financial model with logical progression
        3. Each chunk should advance the model construction"""
        
        progress_context = f"""
        {build_state.financial_model_type.upper()} MODEL BUILDING PROGRESS - STAGE {current_stage}/{total_stages}
        
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
        
        CONTENT PLACEMENT GUIDANCE:
        {placement_guidance}
        
        {progression_requirements}
        4. Focus on {next_stage_description}
        5. STRICTLY AVOID overwriting existing data - use empty ranges only
        
        PREVIOUS ERRORS TO AVOID:
        {chr(10).join(list(set(build_state.error_patterns))[-2:]) if build_state.error_patterns else "None"}
        """
        
        return progress_context
    
    def _determine_build_stage(self, completed_chunks: int, model_type: str = "dcf") -> int:
        """Determine what stage of model building we're in"""
        model_lower = model_type.lower()
        
        if 'three' in model_lower or '3' in model_lower or 'integrated' in model_lower:
            # Three-statement model has more stages
            if completed_chunks < 2:
                return 1  # Headers and setup
            elif completed_chunks < 5:
                return 2  # Assumptions
            elif completed_chunks < 8:
                return 3  # Income Statement - Revenue
            elif completed_chunks < 12:
                return 4  # Income Statement - Expenses
            elif completed_chunks < 16:
                return 5  # Balance Sheet - Assets
            elif completed_chunks < 20:
                return 6  # Balance Sheet - Liabilities
            elif completed_chunks < 25:
                return 7  # Cash Flow Statement
            elif completed_chunks < 30:
                return 8  # Integration and checks
            else:
                return 9  # Formatting and metrics
        else:
            # DCF model stages (more detailed for comprehensive DCF)
            if completed_chunks < 3:
                return 1  # Headers and assumptions
            elif completed_chunks < 6:
                return 2  # P&L projections (Revenue, EBIT, EBIAT)
            elif completed_chunks < 9:
                return 3  # Working capital and CapEx schedules
            elif completed_chunks < 12:
                return 4  # Unlevered Free Cash Flow calculation
            elif completed_chunks < 15:
                return 5  # Terminal Value calculations
            elif completed_chunks < 18:
                return 6  # DCF valuation (PV, EV, Equity Value)
            elif completed_chunks < 21:
                return 7  # WACC and Cost of Equity
            elif completed_chunks < 24:
                return 8  # Sensitivity analysis
            else:
                return 9  # Professional formatting and checks
    
    def _get_next_stage_description(self, stage: int, model_type: str) -> str:
        """Get description of what should be built in the next stage"""
        model_lower = model_type.lower()
        
        if 'three' in model_lower or '3' in model_lower or 'integrated' in model_lower:
            stage_descriptions = {
                1: "Create model title, periods, and main headers for all three statements",
                2: "Build comprehensive assumptions section (growth rates, margins, working capital days, tax rate)",
                3: "Build Income Statement revenue section with growth formulas",
                4: "Add Income Statement expenses (COGS, OpEx) and profitability calculations",
                5: "Create Balance Sheet current assets (cash, AR, inventory) with formulas",
                6: "Add Balance Sheet fixed assets, liabilities, and equity sections",
                7: "Build Cash Flow Statement with operating, investing, and financing activities",
                8: "Add integration checks and ensure all statements are properly linked",
                9: "Apply professional formatting and add key financial metrics/ratios"
            }
        else:
            # DCF model stages (comprehensive DCF methodology)
            stage_descriptions = {
                1: "Create DCF model headers and core assumptions (growth rates, margins, WACC inputs)",
                2: "Build P&L projections with Revenue, EBIT, and EBIAT calculations",
                3: "Add working capital schedules and CapEx projections",
                4: "Calculate Unlevered Free Cash Flow (EBIAT + D&A +/- WC - CapEx)",
                5: "Build Terminal Value using Perpetuity Growth and/or Exit Multiple methods",
                6: "Create DCF valuation with Present Values, Enterprise Value, and Equity Value per share",
                7: "Implement WACC calculation and Cost of Equity (CAPM methodology)",
                8: "Add sensitivity analysis tables (WACC vs Terminal Growth)",
                9: "Apply professional formatting and validation checks"
            }
        
        return stage_descriptions.get(stage, f"Complete the {model_type} model")
    
    def _format_workbook_context(self, context: Dict[str, Any]) -> str:
        """Format workbook context with detailed content analysis"""
        if not context or not context.get('sheets'):
            return "Empty workbook"
        
        sheet_analyses = []
        for sheet in context.get('sheets', [])[:2]:  # Max 2 sheets  
            name = sheet.get('name', 'Unknown')
            data = sheet.get('data', [])
            
            if not data or len(data) == 0:
                sheet_analyses.append(f"Sheet '{name}': completely empty")
                continue
                
            # Analyze content structure and placement
            analysis = self._analyze_sheet_content(data, name)
            sheet_analyses.append(analysis)
        
        return '\n'.join(sheet_analyses)
    
    def _analyze_sheet_content(self, data: List[List], sheet_name: str) -> str:
        """Analyze sheet content to understand layout and identify content types"""
        if not data:
            return f"Sheet '{sheet_name}': empty"
            
        # Find populated regions
        populated_ranges = []
        content_blocks = []
        
        max_row = len(data)
        max_col = max(len(row) for row in data) if data else 0
        
        # Analyze each row for content
        content_start = None
        current_block = []
        
        for row_idx, row in enumerate(data):
            if not row:
                continue
                
            # Check if row has content
            has_content = any(cell and str(cell).strip() for cell in row)
            
            if has_content:
                if content_start is None:
                    content_start = row_idx
                    
                # Analyze row content type
                row_content = [str(cell).strip() for cell in row if cell and str(cell).strip()]
                content_type = self._classify_row_content(row_content, row_idx)
                
                current_block.append({
                    'row': row_idx + 1,  # Excel 1-indexed
                    'content': row_content[:5],  # First 5 non-empty cells
                    'type': content_type
                })
                
                populated_ranges.append(f"Row {row_idx + 1}")
            else:
                # End of content block
                if current_block:
                    content_blocks.append(current_block)
                    current_block = []
                    content_start = None
        
        # Add final block if exists
        if current_block:
            content_blocks.append(current_block)
            
        # Generate detailed analysis
        analysis_parts = [f"Sheet '{sheet_name}':"]
        
        if not content_blocks:
            analysis_parts.append("  - No content detected")
        else:
            # Analyze content blocks
            for i, block in enumerate(content_blocks):
                start_row = block[0]['row']
                end_row = block[-1]['row']
                
                # Classify block type
                block_types = [item['type'] for item in block]
                primary_type = max(set(block_types), key=block_types.count)
                
                # Sample content from block
                sample_content = []
                for item in block[:3]:  # First 3 rows of block
                    sample_content.extend(item['content'][:3])  # First 3 cells
                
                range_desc = f"A{start_row}:Z{end_row}" if end_row > start_row else f"Row {start_row}"
                analysis_parts.append(f"  - {range_desc}: {primary_type} ({', '.join(sample_content[:5])})")
        
        # Find empty regions for suggestions
        if content_blocks:
            last_row = content_blocks[-1][-1]['row']
            if last_row < 30:  # If there's space below
                analysis_parts.append(f"  - Available space: Row {last_row + 2}+ is empty")
        else:
            analysis_parts.append("  - Available space: Entire sheet is empty")
            
        return '\n'.join(analysis_parts)
    
    def _extract_placement_guidance(self, context: Dict[str, Any]) -> str:
        """Extract specific placement guidance to avoid content collisions"""
        if not context or not context.get('sheets'):
            return "✅ PLACEMENT: Use any range starting from A1"
        
        guidance_parts = []
        
        for sheet in context.get('sheets', [])[:1]:  # Focus on active sheet
            name = sheet.get('name', 'Unknown')
            data = sheet.get('data', [])
            
            if not data or len(data) == 0:
                guidance_parts.append(f"✅ PLACEMENT FOR '{name}': Entire sheet is empty - use any range starting from A1")
                continue
            
            # Generate EXPLICIT cell-by-cell forbidden list
            forbidden_cells = []
            occupied_ranges = []
            last_content_row = 0
            
            for row_idx, row in enumerate(data):
                if row and any(cell and str(cell).strip() for cell in row):
                    last_content_row = row_idx + 1  # Excel 1-indexed
                    
                    # Find occupied columns and generate specific cell addresses
                    occupied_cols = []
                    for col_idx, cell in enumerate(row):
                        if cell and str(cell).strip():
                            col_letter = chr(65 + col_idx)  # A=65
                            occupied_cols.append(col_letter)
                            forbidden_cells.append(f"{col_letter}{row_idx + 1}")
                    
                    if occupied_cols:
                        occupied_ranges.append(f"Row {row_idx + 1}: {occupied_cols[0]}-{occupied_cols[-1]}")
            
            if occupied_ranges:
                # Suggest placement below existing content
                suggested_start_row = last_content_row + 2  # Leave one empty row
                
                guidance_parts.append(f"🚫 FORBIDDEN CELLS IN '{name}' (DO NOT USE THESE SPECIFIC CELLS):")
                guidance_parts.append(f"   NEVER use: {', '.join(forbidden_cells[:20])}")  # Show first 20 forbidden cells
                if len(forbidden_cells) > 20:
                    guidance_parts.append(f"   ... and {len(forbidden_cells) - 20} more occupied cells")
                
                guidance_parts.append(f"🚫 OCCUPIED RANGES IN '{name}':")
                guidance_parts.extend([f"   - {range_desc}" for range_desc in occupied_ranges[:5]])
                
                guidance_parts.append(f"✅ SAFE PLACEMENT ZONE: START AT ROW {suggested_start_row} OR LATER")
                guidance_parts.append(f"✅ RECOMMENDED CELLS: A{suggested_start_row}, B{suggested_start_row}, C{suggested_start_row}, etc.")
                guidance_parts.append(f"✅ SAFE RANGES: A{suggested_start_row}:Z{suggested_start_row + 10}")
                
                # Look for gaps between content blocks
                content_gaps = self._find_content_gaps(data)
                if content_gaps:
                    guidance_parts.append(f"✅ ALTERNATIVE SAFE ZONES: {', '.join(content_gaps)}")
            else:
                guidance_parts.append(f"✅ PLACEMENT FOR '{name}': No content detected - use any range")
        
        return '\n'.join(guidance_parts) if guidance_parts else "✅ PLACEMENT: Use any range starting from A1"
    
    def _find_content_gaps(self, data: List[List]) -> List[str]:
        """Find empty ranges between content blocks"""
        if not data:
            return []
        
        gaps = []
        empty_start = None
        
        for row_idx, row in enumerate(data):
            has_content = row and any(cell and str(cell).strip() for cell in row)
            
            if not has_content:
                if empty_start is None:
                    empty_start = row_idx + 1  # Excel 1-indexed
            else:
                if empty_start is not None:
                    # Found end of empty block
                    if row_idx - empty_start + 1 >= 3:  # At least 3 empty rows
                        gaps.append(f"Rows {empty_start}-{row_idx}")
                    empty_start = None
        
        # Check for trailing empty space
        if empty_start is not None and len(data) - empty_start + 1 >= 3:
            gaps.append(f"Rows {empty_start}+")
        
        return gaps[:3]  # Return up to 3 gaps
    
    def _classify_row_content(self, row_content: List[str], row_index: int) -> str:
        """Classify what type of content a row contains"""
        if not row_content:
            return "empty"
            
        # Join content for analysis
        content_text = ' '.join(row_content).lower()
        
        # Header patterns
        if row_index < 3 and any(term in content_text for term in ['assumptions', 'inputs', 'model', 'dcf', 'financial']):
            return "model headers"
        
        # Assumption patterns
        if any(term in content_text for term in ['growth', 'rate', 'margin', 'tax', 'discount', 'wacc', 'assumption']):
            return "assumptions"
            
        # Financial statement patterns
        if any(term in content_text for term in ['revenue', 'sales', 'income', 'expense', 'ebitda', 'ebit']):
            return "P&L projections"
            
        # Cash flow patterns  
        if any(term in content_text for term in ['cash flow', 'capex', 'working capital', 'fcf', 'unlevered']):
            return "cash flow"
            
        # Valuation patterns
        if any(term in content_text for term in ['terminal', 'value', 'pv', 'npv', 'enterprise', 'equity']):
            return "valuation"
            
        # Year headers
        if any(term in content_text for term in ['year', '2024', '2025', '2026', '2027', '2028']):
            return "year headers"
            
        # Check if mostly formulas (contains = signs)
        if any('=' in item for item in row_content):
            return "formulas"
            
        # Check if mostly numbers
        numeric_count = sum(1 for item in row_content if self._is_numeric(item))
        if numeric_count > len(row_content) / 2:
            return "data values"
            
        return "labels/text"
    
    def _is_numeric(self, value: str) -> bool:
        """Check if a string represents a number"""
        try:
            # Remove common formatting
            clean_value = value.replace(',', '').replace('$', '').replace('%', '').strip()
            float(clean_value)
            return True
        except (ValueError, AttributeError):
            return False
    
    def cleanup_session(self, session_id: str) -> bool:
        """Clean up completed or abandoned sessions"""
        
        if session_id in self.active_sessions:
            del self.active_sessions[session_id]
            return True
        
        return False

# Global instance
incremental_builder = IncrementalModelBuilder()