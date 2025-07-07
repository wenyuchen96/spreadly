import anthropic
from typing import Dict, Any, List
from app.core.config import settings
from app.models.spreadsheet import Spreadsheet
import json
import os

class AIService:
    def __init__(self):
        api_key = settings.ANTHROPIC_API_KEY if hasattr(settings, 'ANTHROPIC_API_KEY') else os.getenv('ANTHROPIC_API_KEY')
        if not api_key:
            print("Warning: ANTHROPIC_API_KEY not found. AI features will use mock responses.")
            self.client = None
        else:
            self.client = anthropic.Anthropic(api_key=api_key)
    
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
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
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
        if not self.client:
            return self._mock_query_response(query)
        
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
        """
        
        try:
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}]
            )
            
            result = response.content[0].text
            parsed_response = json.loads(result)
            return parsed_response
        except Exception as e:
            print(f"AI query error: {e}")
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
            response = self.client.messages.create(
                model="claude-3-sonnet-20240229",
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
        # Mock implementation for now
        mock_patterns = [
            {
                "id": 1,
                "pattern_type": "formula",
                "description": "Calculate percentage growth",
                "formula": "=(B2-A2)/A2*100",
                "confidence": 0.85
            },
            {
                "id": 2,
                "pattern_type": "insight",
                "description": "Monthly revenue trend analysis",
                "context": "Time series data with revenue columns",
                "confidence": 0.78
            }
        ]
        
        if pattern_type != "all":
            mock_patterns = [p for p in mock_patterns if p["pattern_type"] == pattern_type]
        
        return mock_patterns
    
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
        return {
            "answer": f"I understand you're asking about: '{query}'. This would normally be processed by Claude AI to provide detailed insights about your spreadsheet data.",
            "formula": "=SUM(A1:A10)",
            "explanation": "This is a mock response. In production, Claude AI would analyze your specific data and provide tailored insights.",
            "next_steps": ["Connect to real AI service", "Upload your spreadsheet data", "Ask specific questions about your data"]
        }
    
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