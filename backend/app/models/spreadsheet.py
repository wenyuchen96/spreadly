from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, JSON
from sqlalchemy.sql import func
from sqlalchemy.orm import relationship
from app.core.database import Base

class Spreadsheet(Base):
    __tablename__ = "spreadsheets"
    
    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("sessions.id"))
    name = Column(String, nullable=False)
    file_hash = Column(String, unique=True, index=True)
    sheet_names = Column(JSON)  # List of sheet names
    row_count = Column(Integer)
    column_count = Column(Integer)
    data_types = Column(JSON)  # Column data types analysis
    summary_stats = Column(JSON)  # Statistical summary
    ai_insights = Column(Text)  # AI-generated insights
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())
    
    session = relationship("Session", back_populates="spreadsheets")