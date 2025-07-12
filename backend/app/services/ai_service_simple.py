import anthropic
from anthropic import APIError
from app.core.config import settings
from app.models.spreadsheet import Spreadsheet
import json
import asyncio
import random
import os
import re
from typing import Dict, Any, List

MODEL_CONFIGS = {
    "claude-3-5-sonnet-20241022": {
        "max_output_tokens": 8192,
    }
}

class AIService:
    def __init__(self):
        print("🔧 Initializing AIService...")
        self.model_name = "claude-3-5-sonnet-20241022"
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
                messages=[{"role": "user", "content": prompt}]
            )
            
            result = response.content[0].text
            analysis = json.loads(result)
            return analysis
        except Exception as e:
            print(f"AI analysis error: {e}")
            return self._mock_analysis()
    
    async def process_natural_language_query(self, session_id: int, query: str) -> Dict[str, Any]:
        """Process natural language query about spreadsheet data"""
        print(f"🔍 Processing query: '{query[:50]}...'")
        print(f"🔍 Claude client status: {self.client is not None}")
        
        if not self.client:
            print("🚨 MOCK TRIGGER: Claude client is None")
            return self._mock_query_response(query)
        
        # Check if user wants a financial model or complex calculation
        model_keywords = ['model', 'dcf', 'financial model', 'valuation', 'cash flow', 'npv', 'irr', 'scenario analysis', 'monte carlo', 'sensitivity analysis']
        wants_model = any(keyword in query.lower() for keyword in model_keywords)
        
        if wants_model:
            # Check if we should use a template
            query_lower = query.lower()
            use_template = any(keyword in query_lower for keyword in ['dcf', 'npv', 'discounted cash flow'])
            
            # Common compatibility rules for all financial models
            compatibility_rules = f"""
            EXCEL.JS API COMPATIBILITY RULES (CRITICAL FOR EXECUTION):
            
            ✅ ALWAYS USE (100% Compatible):
            - sheet.getRange("A1").values = [["value"]] (single cell)
            - sheet.getRange("A1:B2").values = [["a","b"],["c","d"]] (exact dimensions)
            - range.format.fill.color = "#4472C4"
            - range.format.font.bold = true
            - range.format.numberFormat = "$#,##0.00"
            
            ❌ NEVER USE (Causes failures):
            - sheet.getCell() - not available in web Excel  
            - borders.setItem() - not supported
            - Mismatched array dimensions
            
            FINANCIAL MODELING BEST PRACTICES:
            📊 Professional structure with assumptions, calculations, results
            🧮 Use Excel functions: NPV(), IRR(), PMT(), FV(), PV()
            💼 PROFESSIONAL MODELS include: {self._get_model_requirements(query)}
            🎨 Headers: Bold, colored (#4472C4), Assumptions: Light blue (#E7F3FF)
            
            Create a complete, professional-grade {query} model.
            """
            
            if use_template:
                prompt = f"""
                The user is asking for a financial model: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                Use this as your base template and customize it for the user's specific requirements:
                {self._get_base_template(query)}
                
                Customize the template by:
                1. Adjusting assumptions based on user context
                2. Modifying years/periods if specified  
                3. Adding user-specific metrics or calculations
                4. Keeping the professional structure and formatting
                
                {compatibility_rules}
                """
            else:
                prompt = f"""
                The user is asking for a financial model: "{query}"
                
                RETURN ONLY EXECUTABLE JAVASCRIPT CODE - NO JSON, NO EXPLANATIONS, NO MARKDOWN.
                
                Generate JavaScript code using Excel.js API that creates a complete financial model in Excel.
                
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

            # Use higher token limit for financial models
            max_tokens = model_max_tokens if wants_model else 2000

            api_response = None
            for attempt in range(max_retries):
                try:
                    print(f"🔍 Attempting API call {attempt + 1}/{max_retries} with max_tokens: {max_tokens} (financial model: {wants_model})")
                    api_response = await self.client.messages.create(
                        model=self.model_name,
                        max_tokens=max_tokens,
                        timeout=120.0,  # 2 minutes timeout
                        messages=[{"role": "user", "content": prompt}]
                    )
                    break  # Success, exit retry loop
                except APIError as e:
                    # Check for status_code attribute for retriable errors
                    status_code = getattr(e, 'status_code', None)
                    if status_code not in [429, 529]: # 429: RateLimit, 529: Overloaded
                        # Not a retriable error we know about, re-raise to be caught by the outer block
                        raise e

                    if attempt < max_retries - 1:
                        # Longer delays for overload situations
                        if status_code == 529:  # Overloaded
                            delay = base_delay_seconds * (3 ** attempt) + random.uniform(1, 3)
                        else:  # Rate limited
                            delay = base_delay_seconds * (2 ** attempt) + random.uniform(0, 1)
                        
                        error_type = "API Overloaded" if status_code == 529 else "Rate Limited"
                        print(f"🚨 {error_type}. Attempt {attempt + 1}/{max_retries}. Retrying in {delay:.2f} seconds...")
                        await asyncio.sleep(delay)
                    else:
                        print(f"🚨 Max retries reached. Failing after {max_retries} attempts.")
                        raise e
            
            if not api_response:
                # This custom exception will be caught by the outer block
                raise Exception("API call failed after all retries due to persistent overloading or other issues.")
            
            result_text = api_response.content[0].text
            print(f"🔍 Raw Claude response length: {len(result_text)} chars")
            print(f"🔍 Raw response preview: {result_text[:200]}...")
            
            # For financial models, Claude returns raw JavaScript code
            if wants_model:
                print("🔍 Processing financial model response as raw JavaScript")
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