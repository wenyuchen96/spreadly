from sqlalchemy import Column, Integer, String, DateTime, Text, JSON, Float
from sqlalchemy.sql import func
from app.core.database import Base

class Pattern(Base):
    __tablename__ = "patterns"
    
    id = Column(Integer, primary_key=True, index=True)
    pattern_type = Column(String, nullable=False)  # formula, insight, anomaly
    description = Column(Text)
    formula = Column(Text)
    context = Column(JSON)  # Additional context data
    embedding_vector = Column(JSON)  # Vector representation for similarity search
    confidence_score = Column(Float)
    usage_count = Column(Integer, default=0)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())