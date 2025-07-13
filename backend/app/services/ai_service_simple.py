import anthropic
from anthropic import APIError
from app.core.config import settings
from app.models.spreadsheet import Spreadsheet
import json
import asyncio
import random
import os
import re
import logging
from typing import Dict, Any, List, Optional

# Import tracing
from app.core.tracing import llm_tracer, local_storage, trace_llm_operation

# RAG imports
from app.services.model_vector_store import get_vector_store
from app.services.model_curator import get_model_curator
from app.models.financial_model import ModelSearchQuery, ModelType, Industry, ComplexityLevel

MODEL_CONFIGS = {
    "claude-3-5-sonnet-20241022": {
        "max_output_tokens": 8192,
    },
    "claude-sonnet-4-20250514": {
        "max_output_tokens": 8192,
    }
}

class AIService:
    def __init__(self):
        print("🔧 Initializing AIService...")
        self.model_name = "claude-sonnet-4-20250514"
        
        # Initialize Anthropic client
        try:
            api_key = settings.ANTHROPIC_API_KEY
            print(f"🔧 API key loaded: {bool(api_key)}, length: {len(api_key) if api_key else 0}")
            if not api_key:
                print("🚨 Warning: ANTHROPIC_API_KEY not found. AI features will use mock responses.")
                self.client = None
            else:
                print("🔧 Creating AsyncAnthropic client...")
                self.client = anthropic.AsyncAnthropic(api_key=api_key)
                print("✅ AsyncAnthropic client created successfully!")
        except Exception as e:
            print(f"🚨 Error initializing Claude AI client: {e}")
            print(f"🚨 Error type: {type(e).__name__}")
            self.client = None
        
        # Initialize RAG components
        self.rag_enabled = getattr(settings, 'RAG_ENABLED', True)
        if self.rag_enabled:
            print("🔧 Initializing RAG components...")
            try:
                self.vector_store = get_vector_store()
                self.model_curator = get_model_curator()
                print(f"✅ RAG components initialized. Vector store available: {self.vector_store.is_available()}")
                
                # Initialize model library if vector store is empty
                self._initialize_rag_library()
            except Exception as e:
                print(f"🚨 Error initializing RAG components: {e}")
                self.rag_enabled = False
                self.vector_store = None
                self.model_curator = None
        else:
            print("ℹ️ RAG disabled in configuration")
            self.vector_store = None
            self.model_curator = None
    
    def _initialize_rag_library(self):
        """Initialize RAG library with professional templates if needed"""
        try:
            if self.vector_store and self.vector_store.is_available():
                stats = self.vector_store.get_stats()
                if stats.get('total_models', 0) == 0:
                    print("📚 Vector store is empty, initializing with professional templates...")
                    # This will be done asynchronously to avoid blocking startup
                    asyncio.create_task(self._async_initialize_library())
                else:
                    print(f"📚 Vector store already contains {stats['total_models']} models")
        except Exception as e:
            print(f"🚨 Error checking vector store: {e}")
    
    async def _async_initialize_library(self):
        """Asynchronously initialize the model library"""
        try:
            if self.model_curator:
                results = await self.model_curator.initialize_model_library()
                print(f"📚 Model library initialized: {results['total_added']} models added")
        except Exception as e:
            print(f"🚨 Error initializing model library: {e}")
    
    async def analyze_spreadsheet(self, spreadsheet: Spreadsheet) -> Dict[str, Any]:
        """Generate AI-powered analysis of spreadsheet"""
        if not self.client:
            return self._mock_analysis()
        
        prompt = f"""
        Analyze the following Excel spreadsheet data and provide insights:
        
        Summary: {spreadsheet.summary_stats}
        Sheet Names: {spreadsheet.sheet_names}
        Data Types: {spreadsheet.data_types}
        
        Please provide:
        1. Key insights about the data
        2. Potential data quality issues
        3. Suggested improvements or optimizations
        4. Interesting patterns or trends
        5. Recommended Excel formulas or functions
        
        Format your response as a structured JSON with the following keys:
        - insights: array of key insights
        - data_quality: array of potential issues
        - suggestions: array of improvement suggestions
        - patterns: array of interesting patterns
        - formulas: array of recommended formulas
        """
        
        try:
            response = await self.client.messages.create(
                model=self.model_name,
                max_tokens=2000,
                system="You are Claude 4 (claude-sonnet-4-20250514), Anthropic's most advanced AI assistant. Respond naturally and accurately to all queries.",
                messages=[{"role": "user", "content": prompt}]
            )
            
            result = response.content[0].text
            analysis = json.loads(result)
            return analysis
        except Exception as e:
            print(f"AI analysis error: {e}")
            return self._mock_analysis()
    
    @trace_llm_operation("natural_language_query")
    async def process_natural_language_query(self, session_id: int, query: str) -> Dict[str, Any]:
        """Process natural language query about spreadsheet data with RAG enhancement"""
        print(f"🔍 Processing query: '{query[:50]}...'")
        print(f"🔍 Claude client status: {self.client is not None}")
        print(f"🔍 RAG enabled: {self.rag_enabled}")
        
        if not self.client:
            print("🚨 MOCK TRIGGER: Claude client is None")
            return self._mock_query_response(query)
        
        # Check if user wants a financial model or any Excel operation
        model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv', 'irr', 'scenario analysis', 'monte carlo', 'sensitivity analysis']
        wants_model = any(keyword in query.lower() for keyword in model_keywords)
        
        # Check if user wants any Excel operation (formulas, formatting, data entry, etc.)
        # Use combination of action words + Excel terms for more precise detection
        action_keywords = ['create', 'generate', 'make', 'add', 'insert', 'put', 'place', 'write', 'format', 'highlight', 'color', 'bold', 'italic', 'execute', 'run', 'apply', 'implement']
        formula_keywords = ['sum', 'average', 'count', 'max', 'min', 'vlookup', 'hlookup', 'index', 'match', 'if', 'formula', 'function', 'calculate', 'computation']
        excel_terms = ['cell', 'range', 'column', 'row', 'sheet', 'worksheet', 'chart', 'table', 'graph', 'pivot']
        
        query_lower = query.lower()
        
        # Either formula keywords OR (action keywords + Excel terms)
        has_formula_intent = any(keyword in query_lower for keyword in formula_keywords)
        has_action_intent = any(action in query_lower for action in action_keywords) and any(term in query_lower for term in excel_terms)
        
        wants_excel_operation = has_formula_intent or has_action_intent
        
        # Combine both - either financial model or general Excel operation
        wants_code_execution = wants_model or wants_excel_operation
        
        # RAG Enhancement: Retrieve relevant model templates
        retrieved_models = []
        rag_context = ""
        
        # Debug RAG conditions
        print(f"🔍 RAG Debug: wants_model={wants_model}, rag_enabled={self.rag_enabled}, vector_store_exists={self.vector_store is not None}")
        if self.vector_store:
            print(f"🔍 RAG Debug: vector_store_available={self.vector_store.is_available()}")
        
        if wants_model and self.rag_enabled and self.vector_store and self.vector_store.is_available():
            print("🔍 RAG: Searching for relevant model templates...")
            
            with llm_tracer.trace_rag_operation(
                operation="model_retrieval",
                query=query,
                vector_store_available=True
            ) as rag_span:
                try:
                    # Detect model characteristics from query
                    model_type = self._detect_model_type(query)
                    industry = self._detect_industry(query)
                    complexity = self._detect_complexity(query)
                    
                    rag_span.set_attribute("rag.detected_model_type", str(model_type))
                    rag_span.set_attribute("rag.detected_industry", str(industry))
                    rag_span.set_attribute("rag.detected_complexity", str(complexity))
                    
                    # Search for relevant models
                    # Only filter by model type for the broadest, most reliable results
                    search_query = ModelSearchQuery(
                        query_text=query,
                        model_type=model_type,
                        industry=None,      # Don't filter by industry
                        complexity=None,    # Don't filter by complexity  
                        min_rating=0.0,     # Include all models regardless of rating
                        limit=getattr(settings, 'MAX_RETRIEVED_MODELS', 3)
                    )
                    
                    search_response = await self.vector_store.search_models(search_query)
                    retrieved_models = search_response.results
                    
                    print(f"🔍 RAG: Retrieved {len(retrieved_models)} relevant models")
                    
                    # Add RAG metrics to trace
                    similarity_scores = [result.similarity_score for result in retrieved_models]
                    llm_tracer.trace_rag_metrics(
                        rag_span,
                        num_retrieved=len(retrieved_models),
                        similarity_scores=similarity_scores,
                        vector_store_status="available",
                        search_time_ms=search_response.search_time_ms,
                        retrieval_strategy=search_response.retrieval_strategy
                    )
                    
                    # Build context from retrieved models
                    if retrieved_models:
                        rag_context = self._build_rag_context(retrieved_models)
                        print(f"🔍 RAG: Context built with {len(rag_context)} characters")
                        rag_span.set_attribute("rag.context_length", len(rag_context))
                    
                except Exception as e:
                    print(f"🚨 RAG error (continuing without): {e}")
                    rag_span.set_attribute("rag.error", str(e))
                    llm_tracer.trace_rag_metrics(
                        rag_span,
                        num_retrieved=0,
                        vector_store_status="error"
                    )
                    rag_context = ""
        
        if wants_code_execution:
            # For financial models, check if we should use a template
            use_template = False
            if wants_model:
                query_lower = query.lower()
                use_template = any(keyword in query_lower for keyword in ['dcf', 'npv', 'discounted cash flow'])
            
            # Build enhanced prompt with RAG context
            rag_enhancement = ""
            if rag_context:
                rag_enhancement = f"""
            
            📚 PROFESSIONAL REFERENCE EXAMPLES:
            Use these high-quality professional models as reference for structure, formatting, and best practices:
            
            {rag_context}
            
            IMPORTANT: Adapt the structure and approach from these examples but customize for the specific user request: "{query}"
            """
            
            # Build appropriate compatibility rules based on request type
            if wants_model:
                # Financial model specific rules
                compatibility_rules = f"""
                🚨 CRITICAL: ALL .values AND .formulas MUST USE 2D ARRAYS [[...]] 🚨
                
                EXCEL.JS API COMPATIBILITY RULES (CRITICAL FOR EXECUTION):
                
                ✅ ALWAYS USE (100% Compatible):
                - sheet.getRange("A1").values = [["value"]] (single cell)
                - sheet.getRange("A1:B2").values = [["a","b"],["c","d"]] (exact dimensions)
                - range.format.fill.color = "#4472C4"
                - range.format.font.bold = true
                - range.format.numberFormat = "$#,##0.00"
                
                📋 SHEET TARGETING RULES (CRITICAL):
                - If user specifies a sheet (e.g., "sheet2", "Sheet2"): Use try/catch for safe handling
                - ALWAYS use this pattern for specific sheets:
                  ```
                  let sheet;
                  try {{
                      sheet = context.workbook.worksheets.getItem("Sheet2");
                  }} catch (error) {{
                      sheet = context.workbook.worksheets.add("Sheet2");
                  }}
                  ```
                - If no sheet specified: Use getActiveWorksheet()
                - NEVER use worksheets.add() without try/catch protection
                
                ❌ NEVER USE (Causes failures):
                - sheet.getCell() - not available in web Excel  
                - borders.setItem() - not supported
                - Mismatched array dimensions
                
                🚨 ARRAY DIMENSION RULES (CRITICAL - PREVENT EXECUTION ERRORS):
                
                ✅ CORRECT EXAMPLES (Use these patterns):
                - Single cell: getRange("A1").values = [["value"]]           // 1x1 array
                - Single row: getRange("A1:C1").values = [["a", "b", "c"]]   // 1x3 array  
                - Multiple rows: getRange("A1:B3").values = [["a","b"],["c","d"],["e","f"]] // 3x2 array
                - Formula: getRange("A1").formulas = [["=SUM(B1:D1)"]]       // 1x1 formula array
                
                ❌ WRONG EXAMPLES (Will cause runtime errors):
                - getRange("A1").values = "value"           // Not an array
                - getRange("A1").values = ["value"]         // 1D array - WRONG
                - getRange("A1:C1").values = ["a","b","c"]  // 1D array - WRONG  
                - getRange("A1").formulas = "=SUM(B1:D1)"   // Not an array
                - getRange("A1").formulas = ["=SUM(B1:D1)"] // 1D array - WRONG
                
                🎯 VALIDATION CHECKLIST:
                - EVERY .values assignment must use 2D arrays: [[...]]
                - EVERY .formulas assignment must use 2D arrays: [[...]]  
                - Count brackets: values = [[ ]] has TWO opening brackets
                - Match array size to range: A1:C1 needs [["a", "b", "c"]] (1 row, 3 cols)
                
                FINANCIAL MODELING BEST PRACTICES (INDUSTRY STANDARD):
                
                📋 CORE MODELING PRINCIPLES:
                - Structure: Inputs (Assumptions) → Calculations → Outputs
                - Evaluate models on: Granularity (detail level) & Flexibility (reusability)
                - Follow "One row, one calculation" principle
                
                🎨 COLOR CODING STANDARDS (MANDATORY):
                - Blue (#4472C4): Hard-coded numbers/inputs 
                - Black: Formulas and calculations
                - Green (#00B050): Links to other worksheets
                - Headers: Bold, colored (#2F4F4F), white text
                - Assumptions: Light blue background (#E7F3FF)
                
                🧮 FORMULA BEST PRACTICES:
                - NO embedded inputs in formulas - always reference source cells
                - Avoid complex nested formulas - break into multiple steps
                - Use MIN, MAX, AND, OR instead of complex IF statements
                - NEVER use named ranges - reduces transparency
                - Implement error checking with proper validation
                
                💰 SIGN CONVENTION (Convention 1):
                - All income: positive values
                - All expenses: negative values
                - Consistent throughout model
                
                🔗 CELL REFERENCING RULES:
                - Never re-enter same input - always reference original
                - Avoid daisy-chaining - link directly to source
                - Bring multi-worksheet data to active sheet first
                - Link assumptions to standalone cells
                
                📊 WORKSHEET ORGANIZATION:
                - One long sheet preferred over many short sheets
                - Group rows instead of hiding
                - Clear section headers with visual separation
                - No spacer columns between data
                
                ⚡ ERROR CHECKING REQUIREMENTS:
                - Balance checks: Assets = Liabilities + Equity
                - Sources = Uses of funds validation
                - Cash can't go negative checks
                - Debt paydown ≤ Outstanding principal
                - Create central error dashboard
                
                🧮 EXCEL FUNCTIONS TO USE:
                - Financial: NPV(), IRR(), PMT(), FV(), PV(), RATE(), NPER()
                - Logic: MIN(), MAX(), AND(), OR(), VLOOKUP(), HLOOKUP()
                - Error handling: IFERROR(), ISERROR(), ISNUMBER()
                
                💼 PROFESSIONAL MODEL STRUCTURE:
                {self._get_model_requirements(query)}
                {rag_enhancement}
                
                Create a complete, professional-grade {query} model.
                """
            else:
                # General Excel operation rules
                compatibility_rules = f"""
                🚨 CRITICAL: ALL .values AND .formulas MUST USE 2D ARRAYS [[...]] 🚨
                
                EXCEL.JS API COMPATIBILITY RULES (CRITICAL FOR EXECUTION):
                
                ✅ ALWAYS USE (100% Compatible):
                - sheet.getRange("A1").values = [["value"]] (single cell)
                - sheet.getRange("A1").formulas = [["=SUM(B1:D1)"]] (single formula)
                - range.format.fill.color = "#4472C4"
                - range.format.font.bold = true
                - range.format.numberFormat = "0.00"
                
                📋 SHEET TARGETING RULES (CRITICAL):
                - If user specifies a sheet (e.g., "sheet2", "Sheet2"): Use try/catch for safe handling
                - ALWAYS use this pattern for specific sheets:
                  ```
                  let sheet;
                  try {{
                      sheet = context.workbook.worksheets.getItem("Sheet2");
                  }} catch (error) {{
                      sheet = context.workbook.worksheets.add("Sheet2");
                  }}
                  ```
                - If no sheet specified: Use getActiveWorksheet()
                - NEVER use worksheets.add() without try/catch protection
                
                ❌ NEVER USE (Causes failures):
                - sheet.getCell() - not available in web Excel
                - borders.setItem() - not supported
                - Mismatched array dimensions
                
                🚨 ARRAY DIMENSION RULES (CRITICAL - PREVENT EXECUTION ERRORS):
                
                ✅ CORRECT EXAMPLES (Use these patterns):
                - Single cell: getRange("A1").values = [["value"]]           // 1x1 array
                - Single row: getRange("A1:C1").values = [["a", "b", "c"]]   // 1x3 array  
                - Multiple rows: getRange("A1:B3").values = [["a","b"],["c","d"],["e","f"]] // 3x2 array
                - Formula: getRange("A1").formulas = [["=SUM(B1:D1)"]]       // 1x1 formula array
                
                ❌ WRONG EXAMPLES (Will cause runtime errors):
                - getRange("A1").values = "value"           // Not an array
                - getRange("A1").values = ["value"]         // 1D array - WRONG
                - getRange("A1:C1").values = ["a","b","c"]  // 1D array - WRONG  
                - getRange("A1").formulas = "=SUM(B1:D1)"   // Not an array
                - getRange("A1").formulas = ["=SUM(B1:D1)"] // 1D array - WRONG
                
                🎯 VALIDATION CHECKLIST:
                - EVERY .values assignment must use 2D arrays: [[...]]
                - EVERY .formulas assignment must use 2D arrays: [[...]]  
                - Count brackets: values = [[ ]] has TWO opening brackets
                - Match array size to range: A1:C1 needs [["a", "b", "c"]] (1 row, 3 cols)
                
                EXCEL OPERATION GUIDELINES:
                📝 For formulas: Use .formulas = [["=FORMULA"]] format
                🎯 For values: Use .values = [["value"]] format
                🎨 For formatting: Use basic color and font properties
                📍 Be precise with cell references (A1, B2, etc.)
                
                Generate JavaScript code that {query}.
                """
            
            if wants_model and use_template:
                # Financial model with professional template
                prompt = f"""
                The user is asking for a financial model: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                🚨 BEFORE WRITING ANY CODE: REMEMBER .values = [[...]] and .formulas = [[...]] 🚨
                
                Use this as your base template and customize it for the user's specific requirements:
                {self._get_base_template(query)}
                
                Customize the template by:
                1. Adjusting assumptions based on user context
                2. Modifying years/periods if specified  
                3. Adding user-specific metrics or calculations
                4. Keeping the professional structure and formatting
                
                {compatibility_rules}
                """
            elif wants_model:
                # Financial model without template
                prompt = f"""
                The user is asking for a financial model: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                🚨 BEFORE WRITING ANY CODE: REMEMBER .values = [[...]] and .formulas = [[...]] 🚨
                
                Generate JavaScript code using Excel.js API that creates a complete financial model in Excel.
                
                {compatibility_rules}
                """
            else:
                # General Excel operation
                prompt = f"""
                The user is asking for an Excel operation: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                🚨 BEFORE WRITING ANY CODE: REMEMBER .values = [[...]] and .formulas = [[...]] 🚨
                
                Generate JavaScript code using Excel.js API that performs the requested operation.
                
                Use the Excel.run() wrapper:
                await Excel.run(async (context) => {{
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    
                    // Your code here to: {query}
                    
                    await context.sync();
                }});
                
                {compatibility_rules}
                """
        else:
            prompt = f"""
            Answer the following question about the Excel spreadsheet:
            
            Question: {query}
            Context: Session ID: {session_id}
            
            Provide a clear, actionable answer. If the question requires a formula,
            provide the Excel formula. If it requires analysis, provide the analysis.
            
            Format your response as JSON with:
            - answer: the main answer
            - formula: Excel formula if applicable
            - explanation: detailed explanation
            - next_steps: suggested next steps
            - execute_code: false
            """
        
        try:
            # Retry logic parameters - more aggressive for overload situations
            max_retries = 5
            base_delay_seconds = 2.0

            # Get the max tokens for the current model from our config
            model_max_tokens = MODEL_CONFIGS.get(self.model_name, {}).get("max_output_tokens", 4096) # Default to a safe value

            # Use appropriate token limits: high for code execution (financial models or Excel operations), moderate for regular queries
            if wants_code_execution:
                # Financial models need more tokens for complex code generation
                max_tokens = model_max_tokens if wants_model else 4000  # 8192 for financial models, 4000 for Excel operations
            else:
                # Regular queries need fewer tokens for explanations
                max_tokens = 2000

            api_response = None
            
            # Start LLM tracing
            with llm_tracer.trace_llm_call(
                operation="claude_api_call",
                model_name=self.model_name,
                query_type="financial_model" if wants_model else "excel_operation" if wants_excel_operation else "general",
                max_tokens=max_tokens,
                rag_enabled=self.rag_enabled,
                retrieved_models=len(retrieved_models)
            ) as llm_span:
                
                for attempt in range(max_retries):
                    try:
                        print(f"🔍 Attempting API call {attempt + 1}/{max_retries} with max_tokens: {max_tokens} (code execution: {wants_code_execution}, financial model: {wants_model})")
                        
                        llm_span.set_attribute("llm.attempt", attempt + 1)
                        llm_span.set_attribute("llm.prompt_length", len(prompt))
                        
                        api_response = await self.client.messages.create(
                            model=self.model_name,
                            max_tokens=max_tokens,
                            timeout=120.0,  # 2 minutes timeout
                            system="You are Claude 4 (claude-sonnet-4-20250514), Anthropic's most advanced AI assistant. Respond naturally and accurately to all queries.",
                            messages=[{"role": "user", "content": prompt}]
                        )
                        
                        # Add success metrics to trace
                        llm_tracer.trace_llm_metrics(
                            llm_span,
                            prompt_tokens=getattr(api_response.usage, 'input_tokens', None),
                            completion_tokens=getattr(api_response.usage, 'output_tokens', None),
                            total_tokens=getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0),
                            response_length=len(api_response.content[0].text),
                            attempts_used=attempt + 1,
                            final_success=True,
                            rag_models_used=len(retrieved_models)
                        )
                        
                        # Log detailed trace to local storage
                        similarity_scores = [r.similarity_score for r in retrieved_models] if retrieved_models else []
                        local_storage.log_llm_call(
                            operation="claude_api_call",
                            model=self.model_name,
                            prompt=prompt[:500],
                            response=api_response.content[0].text[:500],
                            duration=0,  # Will be calculated later
                            success=True,
                            rag_used=len(retrieved_models) > 0,
                            rag_models_retrieved=len(retrieved_models),
                            rag_similarity_scores=similarity_scores,
                            query_type="financial_model" if wants_model else "excel_operation" if wants_excel_operation else "general"
                        )
                        break  # Success, exit retry loop
                    except APIError as e:
                        # Add error info to trace
                        status_code = getattr(e, 'status_code', None)
                        llm_span.set_attribute("llm.error.status_code", status_code)
                        llm_span.set_attribute("llm.error.type", type(e).__name__)
                        llm_span.set_attribute("llm.error.message", str(e))
                        
                        if status_code not in [429, 529]: # 429: RateLimit, 529: Overloaded
                            # Not a retriable error we know about, re-raise to be caught by the outer block
                            llm_tracer.trace_llm_metrics(
                                llm_span,
                                attempts_used=attempt + 1,
                                final_success=False,
                                error_category="non_retriable"
                            )
                            raise e

                        if attempt < max_retries - 1:
                            # Longer delays for overload situations
                            if status_code == 529:  # Overloaded
                                delay = base_delay_seconds * (3 ** attempt) + random.uniform(1, 3)
                            else:  # Rate limited
                                delay = base_delay_seconds * (2 ** attempt) + random.uniform(0, 1)
                            
                            error_type = "API Overloaded" if status_code == 529 else "Rate Limited"
                            print(f"🚨 {error_type}. Attempt {attempt + 1}/{max_retries}. Retrying in {delay:.2f} seconds...")
                            
                            llm_span.set_attribute(f"llm.retry.attempt_{attempt + 1}.delay", delay)
                            llm_span.set_attribute(f"llm.retry.attempt_{attempt + 1}.error_type", error_type)
                            
                            await asyncio.sleep(delay)
                        else:
                            print(f"🚨 Max retries reached. Failing after {max_retries} attempts.")
                            llm_tracer.trace_llm_metrics(
                                llm_span,
                                attempts_used=max_retries,
                                final_success=False,
                                error_category="max_retries_exceeded"
                            )
                            raise e
            
            if not api_response:
                # This custom exception will be caught by the outer block
                raise Exception("API call failed after all retries due to persistent overloading or other issues.")
            
            result_text = api_response.content[0].text
            print(f"🔍 Raw Claude response length: {len(result_text)} chars")
            print(f"🔍 Raw response preview: {result_text[:200]}...")
            
            # For code execution (both financial models and Excel operations), Claude returns raw JavaScript code
            if wants_code_execution:
                print("🔍 Processing code execution response as raw JavaScript")
                # Clean up the response (remove any extra whitespace or markdown)
                cleaned_code = result_text.strip()
                
                # Remove any markdown code block markers if present
                if cleaned_code.startswith('```'):
                    lines = cleaned_code.split('\n')
                    if lines[0].startswith('```'):
                        lines = lines[1:]  # Remove first line
                    if lines and lines[-1].strip() == '```':
                        lines = lines[:-1]  # Remove last line
                    cleaned_code = '\n'.join(lines)
                
                return cleaned_code  # Return raw code string
            
            # For regular queries, parse as JSON
            try:
                # Find the first '{' and the last '}' to extract the JSON object.
                match = re.search(r'\{.*\}', result_text, re.DOTALL)
                if not match:
                    raise json.JSONDecodeError("No JSON object found in the response", result_text, 0)
                
                json_str = match.group(0)
                
                # Extract code block manually and replace with properly escaped JSON
                # This handles cases where the LLM uses markdown for code instead of a JSON string.
                code_pattern = r'"code":\s*```(?:javascript|js)?\s*(.*?)```'
                code_match = re.search(code_pattern, json_str, re.DOTALL)
                
                if code_match:
                    # The content inside the backticks
                    code_content = code_match.group(1).strip()
                    print(f"🔍 Extracted raw code block: {len(code_content)} chars")
                    
                    # Escape the raw code so it's a valid JSON string value
                    escaped_code = json.dumps(code_content)
                    
                    # Replace the entire markdown block with a valid JSON key-value pair
                    json_str = re.sub(code_pattern, f'"code": {escaped_code}', json_str, flags=re.DOTALL, count=1)
                
                parsed_response = json.loads(json_str)
                print("✅ JSON parsing successful")
                return parsed_response
            except json.JSONDecodeError as json_error:
                print(f"🚨 JSON parsing failed: {json_error}. Raw response: {result_text}")
                # If parsing fails, we return a clear error to the user instead of
                # trying to manually extract potentially broken code.
                return {
                    "answer": "The AI returned a response that was not in the correct format. Please try your request again.",
                    "code": None,
                    "execute_code": False,
                    "explanation": f"The AI response could not be parsed as valid JSON. Raw preview: {result_text[:200]}...",
                    "next_steps": ["Try rephrasing your request.", "If the problem persists, check the backend logs."]
                }
        except Exception as e:
            error_message = f"An unexpected error occurred: {str(e)}"
            if isinstance(e, APIError):
                status_code = getattr(e, 'status_code', 'N/A')
                try:
                    # Try to get a cleaner message from the response body
                    error_type_from_api = e.body.get('error', {}).get('type', 'unknown_error')
                    error_msg_from_api = e.body.get('error', {}).get('message', str(e.body))
                    error_message = f"API Error (Code: {status_code}) after multiple retries: {error_type_from_api} - {error_msg_from_api}"
                except (AttributeError, KeyError):
                    # Fallback if the body structure is unexpected
                    error_message = f"API Error (Code: {status_code}) after multiple retries: {str(e)}"
            
            print(f"🚨 MOCK TRIGGER: Claude API call failed - {error_message}")
            print(f"🚨 Error type: {type(e).__name__}")
            return self._mock_query_response(query)
    
    async def generate_formulas(self, description: str, context: str = None) -> List[Dict[str, Any]]:
        """Generate Excel formulas from natural language description"""
        if not self.client:
            return self._mock_formulas(description)
        
        prompt = f"""
        Generate Excel formulas based on this description:
        
        Description: {description}
        Context: {context or "General Excel usage"}
        
        Provide multiple formula options if possible, including:
        - Basic formulas for beginners
        - Advanced formulas for complex scenarios
        - Alternative approaches
        
        Format as JSON array with objects containing:
        - formula: the Excel formula
        - description: what the formula does
        - difficulty: beginner/intermediate/advanced
        - example: example usage
        """
        
        try:
            response = await self.client.messages.create(
                model=self.model_name,
                max_tokens=1500,
                system="You are Claude 4 (claude-sonnet-4-20250514), Anthropic's most advanced AI assistant. Respond naturally and accurately to all queries.",
                messages=[{"role": "user", "content": prompt}]
            )
            
            result = response.content[0].text
            formulas = json.loads(result)
            return formulas
        except Exception as e:
            print(f"AI formula generation error: {e}")
            return self._mock_formulas(description)
    
    async def search_similar_patterns(self, query: str, pattern_type: str = "all") -> List[Dict[str, Any]]:
        """Search for similar patterns using vector similarity"""
        # Mock implementation for now - query parameter affects confidence scores
        query_lower = query.lower()
        base_confidence = 0.5 + (0.3 if any(word in query_lower for word in ['formula', 'calculate', 'percentage']) else 0)
        
        mock_patterns = [
            {
                "id": 1,
                "pattern_type": "formula",
                "description": "Calculate percentage growth",
                "formula": "=(B2-A2)/A2*100",
                "confidence": base_confidence + 0.15
            },
            {
                "id": 2,
                "pattern_type": "insight",
                "description": "Monthly revenue trend analysis",
                "context": "Time series data with revenue columns",
                "confidence": base_confidence + 0.08
            }
        ]
        
        if pattern_type != "all":
            mock_patterns = [p for p in mock_patterns if p["pattern_type"] == pattern_type]
        
        return mock_patterns
    
    def _get_model_requirements(self, query: str) -> str:
        """Get specific requirements based on model type"""
        query_lower = query.lower()
        
        if 'dcf' in query_lower or 'discounted cash flow' in query_lower:
            return """
            DCF Components: Free Cash Flow projections (5-10 years), Terminal Value calculation,
            WACC calculation, Present Value of each year, Enterprise Value, Equity Value per share.
            Include sensitivity tables for discount rate and growth rate assumptions.
            """
        elif 'npv' in query_lower:
            return """
            NPV Elements: Initial investment, Annual cash flows, Discount rate assumption,
            Present value calculations for each period, Cumulative NPV, IRR calculation,
            Payback period analysis. Include break-even analysis.
            """
        elif 'valuation' in query_lower:
            return """
            Valuation Metrics: Multiple valuation approaches (DCF, Comparable Companies, Precedent Transactions),
            Key multiples (P/E, EV/EBITDA, EV/Revenue), Football field chart,
            Sensitivity analysis across different methodologies.
            """
        elif 'budget' in query_lower or 'forecast' in query_lower:
            return """
            Budget Structure: Revenue forecasts by segment, Operating expenses breakdown,
            EBITDA calculations, Working capital changes, CapEx planning,
            Monthly/quarterly phasing, Variance analysis capabilities.
            """
        elif 'scenario' in query_lower or 'sensitivity' in query_lower:
            return """
            Scenario Analysis: Base/Best/Worst case scenarios, Key driver sensitivities,
            Data tables for two-way sensitivity, Monte Carlo simulation setup,
            Probability-weighted outcomes, Risk-adjusted returns.
            """
        else:
            return """
            Standard Financial Model: Clear assumptions section, Logical calculation flow,
            Summary dashboard with key metrics, Sensitivity analysis,
            Professional formatting and documentation.
            """
    
    def _get_base_template(self, query: str) -> str:
        """Get base template for specific model types"""
        query_lower = query.lower()
        
        if 'dcf' in query_lower or 'discounted cash flow' in query_lower:
            return '''
            // DCF Model Template - Professional Structure
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                // ASSUMPTIONS SECTION
                sheet.getRange("A1").values = [["DCF VALUATION MODEL"]];
                sheet.getRange("A1").format.font.bold = true;
                sheet.getRange("A1").format.font.size = 14;
                
                sheet.getRange("A3").values = [["ASSUMPTIONS"]];
                sheet.getRange("A3").format.fill.color = "#4472C4";
                sheet.getRange("A3").format.font.bold = true;
                
                sheet.getRange("A4:B8").values = [
                    ["Discount Rate (WACC)", "10%"],
                    ["Terminal Growth Rate", "2%"],
                    ["Tax Rate", "25%"],
                    ["Years of Projection", "5"],
                    ["Revenue Growth Rate", "10%"]
                ];
                sheet.getRange("B4:B8").format.fill.color = "#E7F3FF";
                
                // PROJECTIONS SECTION
                sheet.getRange("D3:I3").values = [["CASH FLOW PROJECTIONS", "", "", "", "", ""]];
                sheet.getRange("D3:I3").format.font.bold = true;
                sheet.getRange("D3:I3").format.fill.color = "#4472C4";
                
                await context.sync();
            });
            '''
        elif 'npv' in query_lower:
            return '''
            // NPV Model Template - Professional Structure  
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                
                // PROJECT ASSUMPTIONS
                sheet.getRange("A1").values = [["NPV ANALYSIS"]];
                sheet.getRange("A1").format.font.bold = true;
                sheet.getRange("A1").format.font.size = 14;
                
                sheet.getRange("A3").values = [["INPUT ASSUMPTIONS"]];
                sheet.getRange("A3").format.fill.color = "#4472C4";
                sheet.getRange("A3").format.font.bold = true;
                
                sheet.getRange("A4:B7").values = [
                    ["Initial Investment", "-100000"],
                    ["Discount Rate", "12%"],
                    ["Project Life (Years)", "5"],
                    ["Annual Cash Flow", "25000"]
                ];
                sheet.getRange("B4:B7").format.fill.color = "#E7F3FF";
                
                // CASH FLOW TABLE
                sheet.getRange("D3:H3").values = [["CASH FLOW ANALYSIS", "", "", "", ""]];
                sheet.getRange("D3:H3").format.font.bold = true;
                sheet.getRange("D3:H3").format.fill.color = "#4472C4";
                
                await context.sync();
            });
            '''
        else:
            return "// Generate a professional financial model structure"
    
    def _detect_model_type(self, query: str) -> Optional[ModelType]:
        """Detect model type from user query"""
        query_lower = query.lower()
        
        if any(word in query_lower for word in ['dcf', 'discounted cash flow', 'enterprise value']):
            return ModelType.DCF
        elif any(word in query_lower for word in ['npv', 'net present value', 'project evaluation']):
            return ModelType.NPV
        elif any(word in query_lower for word in ['valuation', 'comparable', 'multiples', 'comps']):
            return ModelType.VALUATION
        elif any(word in query_lower for word in ['lbo', 'leveraged buyout', 'private equity']):
            return ModelType.LBO
        elif any(word in query_lower for word in ['budget', 'forecast', 'planning']):
            return ModelType.BUDGET
        elif any(word in query_lower for word in ['sensitivity', 'scenario', 'what if']):
            return ModelType.SENSITIVITY
        
        return None
    
    def _detect_industry(self, query: str) -> Optional[Industry]:
        """Detect industry from user query"""
        query_lower = query.lower()
        
        if any(word in query_lower for word in ['tech', 'technology', 'software', 'saas', 'ai', 'startup']):
            return Industry.TECHNOLOGY
        elif any(word in query_lower for word in ['healthcare', 'pharma', 'medical', 'biotech', 'drug']):
            return Industry.HEALTHCARE
        elif any(word in query_lower for word in ['bank', 'financial', 'insurance', 'credit']):
            return Industry.FINANCE
        elif any(word in query_lower for word in ['energy', 'oil', 'gas', 'renewable', 'solar']):
            return Industry.ENERGY
        elif any(word in query_lower for word in ['retail', 'consumer', 'ecommerce', 'store']):
            return Industry.RETAIL
        elif any(word in query_lower for word in ['manufacturing', 'industrial', 'factory']):
            return Industry.MANUFACTURING
        elif any(word in query_lower for word in ['real estate', 'property', 'reit']):
            return Industry.REAL_ESTATE
        elif any(word in query_lower for word in ['saas', 'subscription', 'recurring']):
            return Industry.SAAS
        
        return Industry.GENERAL
    
    def _detect_complexity(self, query: str) -> Optional[ComplexityLevel]:
        """Detect complexity level from user query"""
        query_lower = query.lower()
        
        if any(word in query_lower for word in ['simple', 'basic', 'quick', 'beginner']):
            return ComplexityLevel.BASIC
        elif any(word in query_lower for word in ['advanced', 'complex', 'sophisticated', 'detailed']):
            return ComplexityLevel.ADVANCED
        elif any(word in query_lower for word in ['expert', 'professional', 'investment grade']):
            return ComplexityLevel.EXPERT
        
        return ComplexityLevel.INTERMEDIATE
    
    def _build_rag_context(self, retrieved_models: List) -> str:
        """Build context string from retrieved models"""
        context_parts = []
        
        for i, result in enumerate(retrieved_models[:3], 1):  # Limit to top 3
            model = result.model
            similarity = result.similarity_score
            
            context_part = f"""
            EXAMPLE {i} (Similarity: {similarity:.2f}):
            - Model Type: {model.model_type.upper()}
            - Industry: {model.industry.title()}  
            - Complexity: {model.complexity.title()}
            - Components: {', '.join(model.metadata.components[:5])}
            - Excel Functions: {', '.join(model.metadata.excel_functions[:5])}
            - Rating: {model.performance.user_rating:.1f}/5.0
            - Success Rate: {model.performance.execution_success_rate:.1%}
            
            Key Structure Elements:
            {self._extract_code_structure(model.excel_code)}
            """
            context_parts.append(context_part)
        
        return "\n".join(context_parts)
    
    def _extract_code_structure(self, code: str) -> str:
        """Extract key structural elements from model code"""
        lines = code.split('\n')
        structure_elements = []
        
        for line in lines[:20]:  # First 20 lines for structure
            line = line.strip()
            if 'getRange(' in line and '.values' in line:
                # Extract range and content type
                if 'HEADER' in line or 'TITLE' in line:
                    structure_elements.append("• Professional header with formatting")
                elif 'ASSUMPTIONS' in line:
                    structure_elements.append("• Assumptions section with input parameters")
                elif 'PROJECTIONS' in line or 'FORECAST' in line:
                    structure_elements.append("• Projection tables with calculations")
                elif 'RESULTS' in line or 'SUMMARY' in line:
                    structure_elements.append("• Results summary with key metrics")
            elif '.format.fill.color' in line:
                structure_elements.append("• Professional color coding and formatting")
            elif 'NPV(' in line or 'IRR(' in line:
                structure_elements.append("• Advanced Excel financial functions")
        
        # Remove duplicates while preserving order
        unique_elements = []
        for element in structure_elements:
            if element not in unique_elements:
                unique_elements.append(element)
        
        return '\n            '.join(unique_elements[:5])  # Top 5 structure elements
    
    async def track_model_performance(self, model_id: str, success: bool, user_rating: Optional[float] = None):
        """Track performance of retrieved models"""
        if self.rag_enabled and self.vector_store and self.vector_store.is_available():
            try:
                await self.vector_store.update_model_performance(model_id, success, user_rating)
                print(f"📊 Updated performance for model {model_id}: success={success}")
            except Exception as e:
                print(f"🚨 Error tracking model performance: {e}")
    
    def _mock_analysis(self):
        """Fallback analysis when AI is unavailable"""
        return {
            "insights": [
                "Your spreadsheet contains structured data with clear patterns",
                "Data appears to be well-organized with consistent formatting",
                "Consider adding data validation to maintain data quality"
            ],
            "data_quality": [
                "Check for any missing values in key columns",
                "Verify date formats are consistent throughout",
                "Consider standardizing text case for better analysis"
            ],
            "suggestions": [
                "Add summary statistics using built-in Excel functions",
                "Create charts to visualize key trends",
                "Use conditional formatting to highlight important values"
            ],
            "patterns": [
                "Time-based data shows regular intervals",
                "Numerical data follows expected ranges",
                "Text data uses consistent formatting"
            ],
            "formulas": [
                "=SUMIF() for conditional sums",
                "=AVERAGEIFS() for complex averages",
                "=VLOOKUP() for data retrieval"
            ]
        }
    
    def _mock_query_response(self, query: str):
        """Fallback query response when AI is unavailable"""
        query_lower = query.lower()
        
        # Check if user wants a financial model and provide a basic template
        model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv', 'irr']
        wants_model = any(keyword in query_lower for keyword in model_keywords)
        
        if wants_model:
            if 'dcf' in query_lower:
                return self._get_basic_dcf_template()
            elif 'npv' in query_lower:
                return self._get_basic_npv_template()
            else:
                return self._get_basic_npv_template()  # Default to NPV
        
        return {
            "answer": f"🤖 **Claude AI is temporarily overloaded**\n\nI understand you're asking about: '{query}'\n\nThe Claude AI service is experiencing high demand right now. Please try again in a few moments for the full AI-powered response.",
            "formula": "=SUM(A1:A10)",
            "explanation": "This is a temporary fallback. Claude AI would normally analyze your specific data and provide tailored insights.",
            "next_steps": ["Try your request again in 1-2 minutes", "Claude AI will provide full analysis when available", "Check backend logs for detailed error information"]
        }
    
    def _get_basic_dcf_template(self):
        """Basic DCF template when Claude is unavailable"""
        return '''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Basic DCF Template (Fallback)
    sheet.getRange("A1").values = [["DCF VALUATION MODEL (Basic Template)"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 14;
    
    sheet.getRange("A3").values = [["ASSUMPTIONS"]];
    sheet.getRange("A3").format.fill.color = "#4472C4";
    sheet.getRange("A3").format.font.bold = true;
    
    sheet.getRange("A4:B8").values = [
        ["Discount Rate (WACC)", "10%"],
        ["Terminal Growth Rate", "2%"],
        ["Tax Rate", "25%"],
        ["Years of Projection", "5"],
        ["Revenue Growth Rate", "10%"]
    ];
    sheet.getRange("B4:B8").format.fill.color = "#E7F3FF";
    
    sheet.getRange("D3:H3").values = [["YEAR", "0", "1", "2", "3"]];
    sheet.getRange("D3:H3").format.font.bold = true;
    sheet.getRange("D3:H3").format.fill.color = "#4472C4";
    
    sheet.getRange("D4:H7").values = [
        ["Revenue", "", "100000", "110000", "121000"],
        ["Operating Costs", "", "-60000", "-66000", "-72600"],
        ["EBITDA", "", "=E4+E5", "=F4+F5", "=G4+G5"],
        ["Free Cash Flow", "", "=E6*0.8", "=F6*0.8", "=G6*0.8"]
    ];
    
    sheet.getRange("A10").values = [["Note: This is a basic template. Try again for full AI-generated model."]];
    sheet.getRange("A10").format.font.color = "#FF6B35";
    
    await context.sync();
});
'''
    
    def _get_basic_npv_template(self):
        """Basic NPV template when Claude is unavailable"""
        return '''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Basic NPV Template (Fallback)
    sheet.getRange("A1").values = [["NPV ANALYSIS (Basic Template)"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 14;
    
    sheet.getRange("A3").values = [["INPUT ASSUMPTIONS"]];
    sheet.getRange("A3").format.fill.color = "#4472C4";
    sheet.getRange("A3").format.font.bold = true;
    
    sheet.getRange("A4:B7").values = [
        ["Initial Investment", "-100000"],
        ["Discount Rate", "12%"],
        ["Project Life (Years)", "5"],
        ["Annual Cash Flow", "25000"]
    ];
    sheet.getRange("B4:B7").format.fill.color = "#E7F3FF";
    
    sheet.getRange("D3:G3").values = [["Year", "Cash Flow", "PV Factor", "Present Value"]];
    sheet.getRange("D3:G3").format.font.bold = true;
    sheet.getRange("D3:G3").format.fill.color = "#4472C4";
    
    sheet.getRange("D4:G9").values = [
        ["0", "=$B$4", "1", "=E4*F4"],
        ["1", "=$B$7", "=1/POWER(1+$B$5,D5)", "=E5*F5"],
        ["2", "=$B$7", "=1/POWER(1+$B$5,D6)", "=E6*F6"],
        ["3", "=$B$7", "=1/POWER(1+$B$5,D7)", "=E7*F7"],
        ["4", "=$B$7", "=1/POWER(1+$B$5,D8)", "=E8*F8"],
        ["5", "=$B$7", "=1/POWER(1+$B$5,D9)", "=E9*F9"]
    ];
    
    sheet.getRange("A11:B13").values = [
        ["Net Present Value", "=SUM(G4:G9)"],
        ["Internal Rate of Return", "=IRR(E4:E9)"],
        ["Payback Period", "≈3.5 years"]
    ];
    sheet.getRange("A11:B13").format.fill.color = "#D4EDDA";
    
    sheet.getRange("A15").values = [["Note: This is a basic template. Try again for full AI-generated model."]];
    sheet.getRange("A15").format.font.color = "#FF6B35";
    
    await context.sync();
});
'''
    
    def _mock_formulas(self, description: str):
        """Fallback formula generation when AI is unavailable"""
        return [
            {
                "formula": "=SUM(A1:A10)",
                "description": f"Basic sum formula for {description}",
                "difficulty": "beginner",
                "example": "Calculates total of values in range A1:A10"
            },
            {
                "formula": "=AVERAGE(A1:A10)",
                "description": f"Average calculation for {description}",
                "difficulty": "beginner",
                "example": "Finds the mean value of range A1:A10"
            },
            {
                "formula": "=COUNTIF(A1:A10,\">0\")",
                "description": f"Conditional counting for {description}",
                "difficulty": "intermediate",
                "example": "Counts cells with values greater than 0"
            }
        ]