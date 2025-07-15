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
        self.model_name = "claude-sonnet-4-20250514"  # Use Claude 4 as requested
        
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
    
    def _should_use_web_search(self, query: str) -> bool:
        """Determine if web search should be enabled for this query"""
        query_lower = query.lower()
        
        # Web search indicators - questions that likely need current information
        web_search_indicators = [
            'current', 'latest', 'recent', 'today', 'now', 'this year', '2024', '2025',
            'news', 'update', 'trending', 'market price', 'stock price', 'exchange rate',
            'inflation rate', 'interest rate', 'gdp', 'unemployment', 'market data',
            'compare companies', 'competitor analysis', 'industry trends', 'regulations',
            'earnings report', 'financial results', 'market cap', 'valuation',
            'what happened', 'when did', 'who is', 'where is', 'how much is',
            'search for', 'find information', 'look up', 'research',
            'benchmark', 'industry average', 'market standard', 'best practices'
        ]
        
        # Check if query contains web search indicators
        for indicator in web_search_indicators:
            if indicator in query_lower:
                return True
        
        # Don't use web search for Excel-specific operations or code generation
        excel_indicators = [
            'formula', 'cell', 'range', 'sheet', 'workbook', 'pivot',
            'chart', 'graph', 'format', 'calculate', 'sum', 'count',
            'vlookup', 'hlookup', 'macro', 'excel.js', 'javascript'
        ]
        
        for indicator in excel_indicators:
            if indicator in query_lower:
                return False
        
        # Default to no web search for most queries to avoid unnecessary costs
        return False
    
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
    async def process_natural_language_query(self, session_id: int, query: str, workbook_context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Process natural language query about spreadsheet data with RAG enhancement and comprehensive workbook context"""
        print(f"🔍 Processing query: '{query[:50]}...'")
        print(f"🔍 Claude client status: {self.client is not None}")
        print(f"🔍 RAG enabled: {self.rag_enabled}")
        print(f"📊 Workbook context provided: {bool(workbook_context)}")
        
        if workbook_context:
            metadata = workbook_context.get('metadata', {})
            sheets = workbook_context.get('sheets', [])
            tables = workbook_context.get('tables', [])
            print(f"📊 Context: {metadata.get('totalSheets', 0)} sheets, {len(tables)} tables, active: {metadata.get('activeSheetName', 'unknown')}")
            
            # Debug: Print actual sheet data for active sheet
            for sheet in sheets:
                if sheet.get('isActive'):
                    sheet_data = sheet.get('data', [])
                    print(f"🔍 Active sheet '{sheet.get('name')}' data length: {len(sheet_data)}")
                    if sheet_data:
                        print(f"🔍 First row data: {sheet_data[0] if len(sheet_data) > 0 else 'empty'}")
                        print(f"🔍 Second row data: {sheet_data[1] if len(sheet_data) > 1 else 'no second row'}")
                    break
        
        if not self.client:
            print("🚨 MOCK TRIGGER: Claude client is None")
            return self._mock_query_response(query)
        
        # Check if user wants a financial model or any Excel operation
        model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv', 'irr', 'scenario analysis', 'monte carlo', 'sensitivity analysis', 
                         'three statement', '3 statement', 'three-statement', '3-statement', 'integrated model', 'income statement', 'balance sheet', 'cash flow statement']
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
        
        # Build workbook context string
        workbook_context_string = ""
        if workbook_context:
            workbook_context_string = self._build_workbook_context_prompt(workbook_context)
        
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
                
                🔄 SEQUENTIAL EXECUTION OPTIMIZATION:
                - Add clear comment markers for operation stages: // STAGE 1: Setup, // STAGE 2: Data, etc.
                - Group related operations together in logical blocks
                - Use descriptive comments before each major operation
                - Separate sheet setup, data entry, formulas, and formatting into distinct sections
                - Example structure:
                  ```
                  // STAGE 1: Sheet Setup
                  const sheet = context.workbook.worksheets.getActiveWorksheet();
                  
                  // STAGE 2: Headers
                  sheet.getRange("A1").values = [["Header"]];
                  
                  // STAGE 3: Data
                  sheet.getRange("A2").values = [["Data"]];
                  
                  // STAGE 4: Formulas
                  sheet.getRange("A3").formulas = [["=A2*2"]];
                  
                  // STAGE 5: Formatting
                  sheet.getRange("A1").format.font.bold = true;
                  ```
                
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
                {self._get_universal_model_best_practices()}
                
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
                
                🔄 SEQUENTIAL EXECUTION OPTIMIZATION:
                - Add clear comment markers for operation stages: // STAGE 1: Setup, // STAGE 2: Data, etc.
                - Group related operations together in logical blocks
                - Use descriptive comments before each major operation
                - Separate sheet setup, data entry, formulas, and formatting into distinct sections
                - Example structure:
                  ```
                  // STAGE 1: Sheet Setup
                  const sheet = context.workbook.worksheets.getActiveWorksheet();
                  
                  // STAGE 2: Headers
                  sheet.getRange("A1").values = [["Header"]];
                  
                  // STAGE 3: Data
                  sheet.getRange("A2").values = [["Data"]];
                  
                  // STAGE 4: Formulas
                  sheet.getRange("A3").formulas = [["=A2*2"]];
                  
                  // STAGE 5: Formatting
                  sheet.getRange("A1").format.font.bold = true;
                  ```
                
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
                
                {workbook_context_string}
                
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
                
                {workbook_context_string}
                
                Generate JavaScript code using Excel.js API that creates a complete financial model in Excel.
                
                {compatibility_rules}
                """
            else:
                # General Excel operation
                prompt = f"""
                The user is asking for an Excel operation: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                🚨 BEFORE WRITING ANY CODE: REMEMBER .values = [[...]] and .formulas = [[...]] 🚨
                
                {workbook_context_string}
                
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
            
            {workbook_context_string}
            
            Provide a clear, direct answer. If the question requires a formula, provide the Excel formula.
            If it requires analysis, provide the analysis. Keep your response conversational and helpful.
            
            DO NOT format as JSON. DO NOT include sections like "formula:", "explanation:", "next_steps:".
            Just provide a natural, helpful response.
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
                        
                        # Determine if web search should be enabled for this query
                        needs_web_search = self._should_use_web_search(query)
                        
                        # Prepare message parameters
                        message_params = {
                            "model": self.model_name,
                            "max_tokens": max_tokens,
                            "timeout": 120.0,  # 2 minutes timeout
                            "system": "You are Claude 4 (claude-sonnet-4-20250514), Anthropic's most advanced AI assistant. Respond naturally and accurately to all queries.",
                            "messages": [{"role": "user", "content": prompt}]
                        }
                        
                        # Add web search if enabled for this query
                        if needs_web_search:
                            message_params["tools"] = [{"type": "web_search_20250305", "name": "web_search"}]
                            print(f"🌐 Web search enabled for query: {query[:100]}...")
                        
                        api_response = await self.client.messages.create(**message_params)
                        
                        # Calculate total response length safely
                        total_response_length = 0
                        response_preview = ""
                        for content_block in api_response.content:
                            if hasattr(content_block, 'text'):
                                total_response_length += len(content_block.text)
                                response_preview += content_block.text[:500]
                        
                        # Add success metrics to trace
                        llm_tracer.trace_llm_metrics(
                            llm_span,
                            prompt_tokens=getattr(api_response.usage, 'input_tokens', None),
                            completion_tokens=getattr(api_response.usage, 'output_tokens', None),
                            total_tokens=getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0),
                            response_length=total_response_length,
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
                            response=response_preview[:500],
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
            
            # Extract text from the response content (handle both text blocks and tool use blocks)
            result_text = ""
            for content_block in api_response.content:
                if hasattr(content_block, 'text'):
                    result_text += content_block.text
                elif hasattr(content_block, 'type') and content_block.type == 'tool_use':
                    # This is a tool use block - Claude used web search
                    print(f"🌐 Tool used: {content_block.name}")
                    # The actual response text will be in subsequent text blocks
                    continue
                else:
                    print(f"🔍 Unknown content block type: {type(content_block)}")
            
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
                
                # Return code with token information for progress indicator
                return {
                    "text": cleaned_code,
                    "token_usage": {
                        "input_tokens": getattr(api_response.usage, 'input_tokens', None),
                        "output_tokens": getattr(api_response.usage, 'output_tokens', None),
                        "total_tokens": getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0)
                    }
                }
            
            # For regular queries, return response with token information
            print("🔍 Processing regular text response")
            return {
                "text": result_text.strip(),
                "token_usage": {
                    "input_tokens": getattr(api_response.usage, 'input_tokens', None),
                    "output_tokens": getattr(api_response.usage, 'output_tokens', None),
                    "total_tokens": getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0)
                }
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
        """Get specific requirements based on model type - only detailed requirements when specifically needed"""
        query_lower = query.lower()
        
        # Only include detailed model-specific requirements for specific model types
        # to preserve context window for general modeling
        if ('three' in query_lower and 'statement' in query_lower) or ('3' in query_lower and 'statement' in query_lower) or 'integrated model' in query_lower:
            return """
            SPECIFIC THREE-STATEMENT MODEL REQUIREMENTS:
            
            INTEGRATION FOCUS:
            - Net Income flows: IS → Retained Earnings (BS) → Starting point (CF)
            - Working capital changes: BS changes → CF operating activities
            - CapEx: CF investing → PP&E changes on BS
            - Debt: CF financing → Debt balances on BS
            - Cash: CF ending cash → Cash on BS
            
            STATEMENT STRUCTURE:
            - Income Statement: Revenue → COGS → Gross Profit → OpEx → EBITDA → D&A → EBIT → Interest → EBT → Taxes → Net Income
            - Balance Sheet: Current Assets, Fixed Assets = Current Liabilities, Long-term Debt, Equity
            - Cash Flow: Operating (start with NI), Investing (CapEx), Financing (debt/equity changes)
            
            REQUIRED CHECKS:
            - Balance Sheet balances (Assets = Liab + Equity)
            - Cash flow reconciliation (Beginning + Changes = Ending)
            - Working capital days consistency (DSO, DPO, DIO)
            """
        elif 'dcf' in query_lower or 'discounted cash flow' in query_lower:
            return """
            DCF EXPERT SYSTEM:
            
            You are an expert financial analyst specializing in Discounted Cash Flow (DCF) modeling and valuation. Your expertise covers building, analyzing, and interpreting DCF models for enterprise and equity valuation.

            DCF FUNDAMENTALS:
            - DCF values a business as the sum of all future cash flows discounted to present value at a rate reflecting the riskiness of those cash flows
            - Two main approaches: Unlevered DCF (values enterprise) and Levered DCF (values equity directly)
            - Cash flows = Operating cash flows - cash reinvestment
            - Discount rate = Required rate of return based on risk

            UNLEVERED vs LEVERED DCF:
            Unlevered DCF:
            - Values operations for all capital providers (debt and equity)
            - Uses Unlevered Free Cash Flow (UFCF): EBIAT + D&A +/- WC changes - CapEx
            - Discounted at WACC
            - Output is Enterprise Value, subtract net debt for equity value

            Levered DCF:
            - Values business for equity owners only
            - Uses Levered Free Cash Flow (LFCF): CFO - CapEx - debt principal payments
            - Discounted at Cost of Equity
            - Output is Equity Value directly

            DCF IMPLEMENTATION (Two-Stage Model):
            Stage 1: Explicit forecast period (5-10 years)
            - Project unlevered free cash flows annually
            - Link from integrated financial statement model

            Stage 2: Terminal Value
            - Perpetuity Growth Method: TV = FCF(t+1)/(WACC-g), where g = 2-5% typically
            - Exit Multiple Method: TV = Terminal EBITDA × EV/EBITDA multiple
            - Discount TV to present value

            WACC CALCULATION:
            WACC = Cost of Debt × (1-Tax Rate) × (Debt/Total Capital) + Cost of Equity × (Equity/Total Capital)
            - Cost of Debt: Current yield-to-maturity on company debt
            - Cost of Equity: Risk-free rate + Beta × Equity Risk Premium
            - Use market values for weights
            - Assumes constant capital structure

            COST OF EQUITY (CAPM):
            Cost of Equity = Risk-free rate + β × Equity Risk Premium
            - Risk-free rate: 10-year government bond yield
            - Beta: Company's sensitivity to market risk
            - Equity Risk Premium: 4-8% typically
            - Add small-cap or country risk premiums if applicable

            BETA CALCULATION:
            - Public companies: Use regression-based beta from Bloomberg/services
            - Private companies: Use industry beta approach
            - Unlever comparable company betas: β(unlevered) = β(levered)/(1+(1-tax rate)×(Net Debt/Equity))
            - Relever at target capital structure

            NET DEBT CALCULATION:
            Net Debt = Debt + Preferred Stock + Non-controlling Interests - Cash - Non-operating Assets
            - Use book values as proxy for market values
            - Include capital leases, exclude converted securities
            - Test convertibles using if-converted method

            DILUTED SHARES OUTSTANDING:
            Diluted Shares = Basic Shares + Dilutive Securities
            - Include all outstanding options/warrants that are in-the-money
            - Use Treasury Stock Method for options
            - Include unvested restricted stock
            - Test convertible securities for dilution

            TERMINAL VALUE CONSIDERATIONS:
            - Normalize terminal FCF for sustainable growth
            - Converge CapEx/Depreciation ratio to 1.0
            - Remove one-time working capital swings
            - Ensure growth rate < economy growth rate
            - Terminal value often 50-80% of total value

            DCF BEST PRACTICES:
            - Present results as ranges via sensitivity analysis
            - Key sensitivities: WACC, terminal growth, operating margins
            - Link to integrated 3-statement model
            - Match cash flows to discount rates consistently
            - Address circularity from cash/WACC interaction

            TECHNICAL IMPLEMENTATION:
            - UFCF starts with EBIAT (EBIT × (1-tax rate)) to avoid double-counting interest tax shield
            - Interest tax shield captured in WACC, not cash flows
            - For negative net debt, equity weight >100%, debt weight negative
            - Stock splits require retroactive adjustment of all share counts
            - Model plug: Cash and revolver balance automatically

            Focus on practical DCF implementation while maintaining theoretical accuracy, emphasizing the matching principle between cash flows and discount rates, proper treatment of non-operating items, and the critical importance of terminal value assumptions.
            """
        elif 'npv' in query_lower:
            return """
            NPV FOCUS: Initial investment, Annual cash flows, Discount rate, Present values, IRR, Payback period.
            """
        elif 'lbo' in query_lower or 'leverage' in query_lower:
            return """
            LBO FOCUS: Sources/Uses, Debt schedules, Interest calculations, Credit metrics, Returns (IRR/MOIC).
            """
        elif 'valuation' in query_lower:
            return """
            VALUATION FOCUS: Multiple approaches (DCF, Comps, Precedents), Key multiples, Football field chart.
            """
        elif 'budget' in query_lower or 'forecast' in query_lower:
            return """
            BUDGET FOCUS: Revenue forecasts, Expense breakdown, EBITDA, Working capital, CapEx, Variance analysis.
            """
        else:
            return """
            GENERAL MODEL FOCUS: Clear structure, Key assumptions, Calculations, Outputs, Professional presentation.
            """
    
    def _get_universal_model_best_practices(self) -> str:
        """Universal financial modeling best practices - always included for any financial model"""
        return """
        
        💼 UNIVERSAL FINANCIAL MODELING BEST PRACTICES:
        
        🎨 COLOR CODING STANDARDS:
        - Blue (#0070C0): Hard-coded inputs and assumptions
        - Black (#000000): Formulas and calculations  
        - Green (#00B050): Links to other worksheets
        - Red (#FF0000): External links or warnings
        
        📊 PROFESSIONAL FORMATTING:
        - Consistent decimal places (0 for whole numbers, 1 for percentages, 2 for currency)
        - Standard column widths and aligned headers
        - Years/periods clearly labeled across columns
        - Clear section breaks and subtotals
        - Appropriate number formatting ($, %, etc.)
        
        🔍 MODEL VALIDATION & CHECKS:
        - Balance checks where applicable (Assets = Liabilities + Equity)
        - Cash flow reconciliation (Beginning + Changes = Ending)
        - Error checking with IFERROR() functions
        - Sensitivity analysis on key drivers
        - Sources = Uses validation for capital structures
        - Sanity checks (growth rates, margins, ratios within reasonable ranges)
        
        🏗️ STRUCTURE PRINCIPLES:
        - Clear assumptions/inputs section at top
        - Logical flow: Inputs → Calculations → Outputs
        - Documentation and source references
        - Scenario analysis capabilities
        - Summary dashboard with key metrics
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
    
    def _get_model_sections_prompt(self, model_type: str) -> str:
        """Get model-specific sections for incremental building"""
        model_lower = model_type.lower()
        
        if 'three' in model_lower or '3' in model_lower or 'integrated' in model_lower:
            return """THREE-STATEMENT MODEL SECTIONS (build in order):
1-2. Headers and assumptions (growth, margins, working capital, tax rate)
3-6. Income Statement (Revenue → COGS → OpEx → EBITDA → D&A → EBIT → Interest → EBT → Tax → NI)
7-10. Balance Sheet (Current Assets, Fixed Assets, Current Liab, Debt, Equity)
11-14. Cash Flow (Operating from NI, Working capital changes, Investing, Financing)
15-17. Integration checks, metrics, formatting"""
        
        elif 'dcf' in model_lower or 'discounted' in model_lower:
            return """DCF MODEL SECTIONS (build in order):
1-2. Headers and assumptions (revenue growth, operating margins, tax rate, WACC components)
3-4. Historical/projected P&L (Revenue, EBIT, EBIAT calculations)
5-6. Working capital and CapEx schedules
7-8. Unlevered Free Cash Flow calculation (EBIAT + D&A +/- WC - CapEx)
9-10. Terminal Value (Perpetuity Growth and/or Exit Multiple methods)
11-12. DCF valuation (Present values, Enterprise Value, Equity Value per share)
13-14. WACC calculation and Cost of Equity (CAPM)
15-16. Sensitivity analysis (WACC vs Terminal Growth, key operating assumptions)
17. Professional formatting and checks"""
        
        elif 'lbo' in model_lower:
            return """LBO MODEL SECTIONS (build in order):
1-2. Transaction assumptions and sources/uses
3-4. Operating model and debt schedules
5-6. Credit metrics and returns analysis (IRR, MOIC)
7. Sensitivity tables and formatting"""
        
        else:
            # Default financial model sections
            return """FINANCIAL MODEL SECTIONS (build in order):
1-2. Headers and key assumptions
3-4. Revenue projections and expense calculations
5-6. Profitability analysis and key metrics
7. Professional formatting"""
    
    def _build_workbook_context_prompt(self, workbook_context: Dict[str, Any]) -> str:
        """Build a comprehensive context prompt from workbook data"""
        if not workbook_context:
            return ""
        
        context_parts = []
        
        # Metadata
        metadata = workbook_context.get('metadata', {})
        if metadata:
            context_parts.append(f"""
📊 CURRENT WORKBOOK CONTEXT:
- Total sheets: {metadata.get('totalSheets', 0)}
- Active sheet: {metadata.get('activeSheetName', 'unknown')}
- Last modified: {metadata.get('lastModified', 'unknown')}""")
        
        # Sheets information
        sheets = workbook_context.get('sheets', [])
        if sheets:
            context_parts.append(f"""
📋 SHEET STRUCTURE ({len(sheets)} sheets):""")
            
            for sheet in sheets[:5]:  # Limit to first 5 sheets to avoid token overflow
                # Handle both string and dict formats for backward compatibility
                if isinstance(sheet, str):
                    sheet_name = sheet
                    sheet_info = f"  • {sheet_name}"
                    data = []
                else:
                    sheet_name = sheet.get('name', 'Unknown')
                    sheet_info = f"  • {sheet_name}"
                    if sheet.get('isActive'):
                        sheet_info += " (ACTIVE)"
                    
                    used_range = sheet.get('usedRange')
                    if used_range:
                        sheet_info += f" - {used_range.get('rowCount', 0)}x{used_range.get('columnCount', 0)} used range"
                    else:
                        sheet_info += " - empty"
                    
                    # Add sample data if available and not too large
                    data = sheet.get('data', [])
                print(f"🔍 Backend processing sheet '{sheet_name}' - data type: {type(data)}, length: {len(data) if data else 'None'}")
                if data:
                    print(f"🔍 First row sample: {data[0] if len(data) > 0 else 'No first row'}")
                
                if data and len(data) > 0:
                    # Show first few rows/cols of data
                    sample_rows = min(5, len(data))  # Show up to 5 rows instead of 3
                    sheet_info += f"\n    📋 Actual Data ({len(data)} rows):"
                    
                    # Check if data has actual content
                    has_real_content = False
                    for i in range(sample_rows):
                        if i < len(data):
                            row_data = data[i][:10] if data[i] else []  # Show up to 10 columns
                            
                            # Check if row has meaningful content (not just empty strings/nulls)
                            meaningful_cells = [cell for cell in row_data if cell not in [None, '', 0, False]]
                            if meaningful_cells:
                                has_real_content = True
                            
                            sheet_info += f"\n      Row {i+1}: {row_data}"
                            if meaningful_cells:
                                sheet_info += f" ← {len(meaningful_cells)} non-empty cells"
                        else:
                            sheet_info += f"\n      Row {i+1}: [missing row]"
                    
                    if len(data) > sample_rows:
                        sheet_info += f"\n      ... and {len(data) - sample_rows} more rows"
                    
                    if not has_real_content:
                        sheet_info += f"\n    ⚠️ Note: All cells appear to be empty or contain default values"
                        
                elif data is not None and len(data) == 0:
                    sheet_info += f"\n    📋 Data: Empty sheet (no data array)"
                else:
                    sheet_info += f"\n    📋 Data: No data field provided"
                
                context_parts.append(sheet_info)
        
        # Tables information
        tables = workbook_context.get('tables', [])
        if tables:
            context_parts.append(f"""
🗂 TABLES ({len(tables)} tables):""")
            for table in tables[:3]:  # Limit to avoid token overflow
                table_info = f"  • {table.get('name', 'Unknown')} in {table.get('sheetName', 'Unknown sheet')}"
                headers = table.get('headers', [])
                if headers:
                    table_info += f" - Columns: {', '.join(headers[:5])}"  # First 5 columns
                table_info += f" ({table.get('rowCount', 0)} rows)"
                context_parts.append(table_info)
        
        # Named ranges
        named_ranges = workbook_context.get('namedRanges', [])
        if named_ranges:
            context_parts.append(f"""
🏷 NAMED RANGES ({len(named_ranges)} ranges):""")
            for named_range in named_ranges[:5]:  # Limit to avoid token overflow
                range_info = f"  • {named_range.get('name', 'Unknown')}: {named_range.get('formula', 'Unknown formula')}"
                context_parts.append(range_info)
        
        # Summary information
        summary = workbook_context.get('summary', {})
        if summary:
            context_parts.append(f"""
📈 WORKBOOK SUMMARY:
- Total cells: {summary.get('totalCells', 0)}
- Used cells: {summary.get('totalUsedCells', 0)}
- Has formulas: {summary.get('hasFormulas', False)}
- Has charts: {summary.get('hasCharts', False)}""")
        
        if context_parts:
            full_context = "".join(context_parts)
            full_context += "\n\n💡 USE THIS CONTEXT: Consider the existing data, structure, and active sheet when generating code."
            return full_context
        
        return ""
    
    @trace_llm_operation("incremental_chunk_generation")
    async def generate_incremental_chunk(
        self, 
        session_id: int, 
        model_type: str,
        build_context: str,
        workbook_context: Dict[str, Any] = None,
        previous_errors: List[str] = None
    ) -> Dict[str, Any]:
        """Generate a single optimized code chunk for incremental model building"""
        
        print(f"🔧 Generating incremental chunk for {model_type} model")
        
        if not self.client:
            print("🚨 MOCK TRIGGER: Claude client is None for chunk generation")
            return self._mock_chunk_response(model_type)
        
        # Build workbook context string
        workbook_context_string = ""
        if workbook_context:
            workbook_context_string = self._build_workbook_context_prompt(workbook_context)
        
        # Build error avoidance context
        error_context = ""
        if previous_errors:
            error_context = f"""
ERRORS TO AVOID (from previous attempts):
{chr(10).join(f"- {error}" for error in previous_errors[-3:])}

IMPORTANT: Analyze these errors and avoid similar patterns in your code generation.
"""
        
        # Build incremental chunk prompt with STRICT code-only output
        chunk_prompt = f"""
SYSTEM: You are an expert financial modeling JavaScript code generator. You MUST return ONLY executable JavaScript code with NO explanations, NO markdown, NO analysis text.

🚨 CRITICAL CODE COMPLETION REQUIREMENTS 🚨
1. ALWAYS complete your code chunks - never end mid-statement
2. If approaching token limit, prioritize completing current operations
3. End chunks at logical completion points (after context.sync())
4. Ensure all opened brackets {{ }} are properly closed
5. NEVER end with incomplete sheet.getRange() calls
6. NEVER end with partial string literals or incomplete .values assignments
7. Complete all lines with proper semicolons
8. Always close Excel.run() wrapper with context.sync() and closing braces

FINANCIAL MODELING STANDARDS:
- Color code: Blue (#0070C0) inputs, Black (#000000) formulas, Green (#00B050) links
- Professional formatting: Consistent decimals, clear headers, proper number formats
- Include validation: Balance checks, error handling with IFERROR(), sanity checks

TASK: Generate the next small chunk of Excel.js code for incremental {model_type.upper()} model building.

CONTEXT:
{build_context}
{workbook_context_string}
{error_context}

🚨 CRITICAL OUTPUT REQUIREMENTS 🚨
1. Start your response immediately with: await Excel.run(async (context) => {{
2. End your response with: }});
3. NO text before or after the code
4. NO explanations, analysis, or comments outside the code
5. NO markdown code fences (```)
6. NO "Looking at the errors" or similar analysis text

✅ REQUIRED CODE STRUCTURE ✅
await Excel.run(async (context) => {{
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // 2-4 Excel operations that ADVANCE the model construction (keep chunks small to avoid truncation)
    sheet.getRange("A1").values = [["value"]];  // 2D arrays required
    sheet.getRange("A2").formulas = [["=SUM(A1)"]];  // 2D arrays required
    
    await context.sync();
}});

🎯 PROGRESSION REQUIREMENTS:
- DO NOT repeat headers or assumptions if already created
- ADVANCE to the next logical section based on context
- Build DIFFERENT content each time - revenue, expenses, formulas, etc.
- Each chunk should add NEW functionality to the DCF model

⚡ SYNTAX: All .values and .formulas must use 2D arrays: [["value"]]
🔒 SECURITY: No eval(), no external calls, only Excel.js operations

{self._get_model_sections_prompt(model_type)}

🚨 CODE COMPLETION GUARANTEE 🚨
- Generate ONLY complete, executable code
- MUST start with: await Excel.run(async (context) => {{
- MUST end with: }});
- NEVER end mid-statement or mid-formula
- Complete all array brackets [...] and quotes "..."
- If running out of space, prioritize completing current operations over adding new ones

Generate complete, executable code starting with await Excel.run and ending with }});
"""

        try:
            # Use maximum available tokens for complete code generation
            max_tokens = MODEL_CONFIGS.get(self.model_name, {}).get("max_output_tokens", 8192)
            
            with llm_tracer.trace_llm_call(
                operation="incremental_chunk_generation",
                model_name=self.model_name,
                query_type=f"incremental_{model_type}",
                max_tokens=max_tokens,
                rag_enabled=False,
                retrieved_models=0
            ) as llm_span:
                
                print(f"🔧 Generating chunk with max_tokens: {max_tokens} (maximum available for {self.model_name})")
                
                llm_span.set_attribute("llm.prompt_length", len(chunk_prompt))
                
                api_response = await self.client.messages.create(
                    model=self.model_name,
                    max_tokens=max_tokens,
                    timeout=60.0,  # Shorter timeout for chunks
                    system="You are a JavaScript code generator. Return ONLY executable JavaScript code. NO explanations, NO analysis, NO markdown. Start with 'await Excel.run' and end with '});'. Use 2D arrays for .values = [['value']].",
                    messages=[{"role": "user", "content": chunk_prompt}]
                )
                
                # Add success metrics to trace
                llm_tracer.trace_llm_metrics(
                    llm_span,
                    prompt_tokens=getattr(api_response.usage, 'input_tokens', None),
                    completion_tokens=getattr(api_response.usage, 'output_tokens', None),
                    total_tokens=getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0),
                    response_length=len(api_response.content[0].text),
                    attempts_used=1,
                    final_success=True,
                    rag_models_used=0
                )
                
                chunk_code = api_response.content[0].text.strip()
                
                # Clean up the response (remove any markdown formatting)
                if chunk_code.startswith('```'):
                    lines = chunk_code.split('\n')
                    chunk_code = '\n'.join(lines[1:-1]) if len(lines) > 2 else chunk_code
                
                print(f"✅ Generated chunk ({len(chunk_code)} characters)")
                return {
                    "code": chunk_code,
                    "token_usage": {
                        "input_tokens": getattr(api_response.usage, 'input_tokens', None),
                        "output_tokens": getattr(api_response.usage, 'output_tokens', None),
                        "total_tokens": getattr(api_response.usage, 'input_tokens', 0) + getattr(api_response.usage, 'output_tokens', 0)
                    }
                }
                
        except Exception as e:
            print(f"❌ Error generating incremental chunk: {e}")
            mock_code = self._mock_chunk_response(model_type)
            return {
                "code": mock_code,
                "token_usage": {
                    "input_tokens": 0,
                    "output_tokens": 0,
                    "total_tokens": 0
                }
            }
    
    def _mock_chunk_response(self, model_type: str) -> str:
        """Fallback chunk generation when AI is unavailable"""
        return f'''
await Excel.run(async (context) => {{
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Mock {model_type} model chunk
    sheet.getRange("A1").values = [["{model_type.upper()} MODEL"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.fill.color = "#4472C4";
    
    await context.sync();
}});
'''