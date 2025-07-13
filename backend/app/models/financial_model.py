"""
Financial Model Data Structures for RAG Implementation
"""

from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional
from datetime import datetime
from enum import Enum


class ModelType(str, Enum):
    """Types of financial models"""
    DCF = "dcf"
    NPV = "npv"
    VALUATION = "valuation"
    LBO = "lbo"
    BUDGET = "budget"
    SENSITIVITY = "sensitivity"
    SCENARIO = "scenario"


class Industry(str, Enum):
    """Industry classifications"""
    TECHNOLOGY = "technology"
    HEALTHCARE = "healthcare"
    FINANCE = "finance"
    ENERGY = "energy"
    RETAIL = "retail"
    MANUFACTURING = "manufacturing"
    REAL_ESTATE = "real_estate"
    SAAS = "saas"
    GENERAL = "general"


class ComplexityLevel(str, Enum):
    """Model complexity levels"""
    BASIC = "basic"
    INTERMEDIATE = "intermediate"
    ADVANCED = "advanced"
    EXPERT = "expert"


class ModelMetadata(BaseModel):
    """Metadata for financial model templates"""
    components: List[str] = Field(description="Key components like 'free_cash_flow', 'terminal_value'")
    excel_functions: List[str] = Field(description="Excel functions used like 'NPV', 'IRR', 'XNPV'")
    formatting_features: List[str] = Field(description="Formatting elements like 'conditional_formatting', 'charts'")
    business_assumptions: List[str] = Field(description="Business logic assumptions")
    time_horizon_years: Optional[int] = Field(description="Projection period in years")
    currencies: List[str] = Field(default=["USD"], description="Supported currencies")
    regions: List[str] = Field(default=["global"], description="Geographic applicability")


class PerformanceMetrics(BaseModel):
    """Model performance tracking"""
    execution_success_rate: float = Field(ge=0.0, le=1.0, description="Rate of successful executions")
    user_rating: float = Field(ge=0.0, le=5.0, description="Average user rating")
    usage_count: int = Field(ge=0, description="Number of times model was retrieved")
    last_used: Optional[datetime] = Field(description="Last time model was used")
    error_count: int = Field(ge=0, description="Number of execution errors")
    modification_frequency: float = Field(ge=0.0, description="How often users modify the generated model")


class FinancialModel(BaseModel):
    """Complete financial model document for vector storage"""
    
    # Identification
    id: str = Field(description="Unique identifier for the model")
    name: str = Field(description="Human-readable name")
    description: str = Field(description="Detailed description of the model")
    
    # Classification
    model_type: ModelType = Field(description="Type of financial model")
    industry: Industry = Field(description="Target industry")
    complexity: ComplexityLevel = Field(description="Complexity level")
    
    # Content
    excel_code: str = Field(description="JavaScript code for Excel.js execution")
    business_description: str = Field(description="Business context and use case")
    sample_inputs: Dict[str, Any] = Field(description="Example input values")
    expected_outputs: Dict[str, Any] = Field(description="Expected calculation results")
    
    # Metadata
    metadata: ModelMetadata = Field(description="Rich metadata for retrieval")
    performance: PerformanceMetrics = Field(description="Performance and usage metrics")
    
    # Versioning and provenance
    version: str = Field(default="1.0.0", description="Model version")
    created_by: str = Field(description="Source or creator")
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    
    # Tags for enhanced retrieval
    tags: List[str] = Field(default=[], description="Additional tags for search")
    keywords: List[str] = Field(description="Keywords for semantic search")
    
    class Config:
        use_enum_values = True


class ModelSearchQuery(BaseModel):
    """Query structure for model retrieval"""
    query_text: str = Field(description="Natural language query")
    model_type: Optional[ModelType] = Field(description="Filter by model type")
    industry: Optional[Industry] = Field(description="Filter by industry")
    complexity: Optional[ComplexityLevel] = Field(description="Filter by complexity")
    min_rating: Optional[float] = Field(ge=0.0, le=5.0, description="Minimum user rating")
    limit: int = Field(default=5, ge=1, le=20, description="Number of results to return")
    include_metadata: bool = Field(default=True, description="Include metadata in results")


class ModelSearchResult(BaseModel):
    """Result from model search"""
    model: FinancialModel = Field(description="The retrieved model")
    similarity_score: float = Field(ge=0.0, le=1.0, description="Semantic similarity score")
    relevance_explanation: str = Field(description="Why this model was selected")


class ModelSearchResponse(BaseModel):
    """Response from model search operation"""
    query: ModelSearchQuery = Field(description="Original search query")
    results: List[ModelSearchResult] = Field(description="Retrieved models")
    total_found: int = Field(description="Total number of matching models")
    search_time_ms: float = Field(description="Search execution time")
    retrieval_strategy: str = Field(description="Strategy used for retrieval")