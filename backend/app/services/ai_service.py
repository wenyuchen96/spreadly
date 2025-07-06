from langchain_anthropic import ChatAnthropic
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from typing import Dict, Any, List
from app.core.config import settings
from app.models.spreadsheet import Spreadsheet
from app.models.pattern import Pattern
import json

class AIService:
    def __init__(self):
        self.llm = ChatAnthropic(
            anthropic_api_key=settings.ANTHROPIC_API_KEY,
            model="claude-3-sonnet-20240229"
        )
    
    async def analyze_spreadsheet(self, spreadsheet: Spreadsheet) -> Dict[str, Any]:
        """Generate AI-powered analysis of spreadsheet"""
        prompt = PromptTemplate(
            input_variables=["data_summary", "sheet_names", "data_types"],
            template="""
            Analyze the following Excel spreadsheet data and provide insights:
            
            Summary: {data_summary}
            Sheet Names: {sheet_names}
            Data Types: {data_types}
            
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
        )
        
        chain = LLMChain(llm=self.llm, prompt=prompt)
        
        result = await chain.arun(
            data_summary=spreadsheet.summary_stats,
            sheet_names=spreadsheet.sheet_names,
            data_types=spreadsheet.data_types
        )
        
        try:
            analysis = json.loads(result)
        except json.JSONDecodeError:
            analysis = {"raw_analysis": result}
        
        return analysis
    
    async def process_natural_language_query(self, session_id: int, query: str) -> Dict[str, Any]:
        """Process natural language query about spreadsheet data"""
        prompt = PromptTemplate(
            input_variables=["query", "context"],
            template="""
            Answer the following question about the Excel spreadsheet:
            
            Question: {query}
            Context: {context}
            
            Provide a clear, actionable answer. If the question requires a formula,
            provide the Excel formula. If it requires analysis, provide the analysis.
            
            Format your response as JSON with:
            - answer: the main answer
            - formula: Excel formula if applicable
            - explanation: detailed explanation
            - next_steps: suggested next steps
            """
        )
        
        chain = LLMChain(llm=self.llm, prompt=prompt)
        
        result = await chain.arun(
            query=query,
            context=f"Session ID: {session_id}"
        )
        
        try:
            response = json.loads(result)
        except json.JSONDecodeError:
            response = {"answer": result}
        
        return response
    
    async def generate_formulas(self, description: str, context: str = None) -> List[Dict[str, Any]]:
        """Generate Excel formulas from natural language description"""
        prompt = PromptTemplate(
            input_variables=["description", "context"],
            template="""
            Generate Excel formulas based on this description:
            
            Description: {description}
            Context: {context}
            
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
        )
        
        chain = LLMChain(llm=self.llm, prompt=prompt)
        
        result = await chain.arun(
            description=description,
            context=context or "General Excel usage"
        )
        
        try:
            formulas = json.loads(result)
        except json.JSONDecodeError:
            formulas = [{"formula": "Error parsing response", "description": result}]
        
        return formulas
    
    async def search_similar_patterns(self, query: str, pattern_type: str = "all") -> List[Dict[str, Any]]:
        """Search for similar patterns using vector similarity"""
        # This would integrate with Pinecone for actual vector search
        # For now, returning mock data
        
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
        
        # Filter by pattern type if specified
        if pattern_type != "all":
            mock_patterns = [p for p in mock_patterns if p["pattern_type"] == pattern_type]
        
        return mock_patterns
    
    async def extract_embeddings(self, text: str) -> List[float]:
        """Extract embeddings for text (for Pinecone storage)"""
        # This would use a proper embedding model
        # For now, returning mock embeddings
        return [0.1] * 1536  # Mock 1536-dimensional vector