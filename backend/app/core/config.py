from pydantic_settings import BaseSettings
from typing import List

class Settings(BaseSettings):
    # Database
    DATABASE_URL: str = "sqlite:///./spreadly.db"
    
    # Redis
    REDIS_URL: str = "redis://localhost:6379"
    
    # API Keys
    ANTHROPIC_API_KEY: str = ""
    PINECONE_API_KEY: str = ""
    PINECONE_ENVIRONMENT: str = "us-west1-gcp"
    PINECONE_INDEX_NAME: str = "spreadly-patterns"
    
    # Security
    SECRET_KEY: str = "your-secret-key-change-this"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    
    # CORS
    ALLOWED_HOSTS: List[str] = [
        "http://localhost:3000", 
        "https://localhost:3000", 
        "https://localhost:3001",
        "https://excel.officeapps.live.com",
        "https://excel.office.com", 
        "https://office.live.com",
        "*"  # Allow all origins for development
    ]
    
    # File Upload
    MAX_FILE_SIZE: int = 10 * 1024 * 1024  # 10MB
    UPLOAD_DIR: str = "uploads"
    
    # Environment
    ENVIRONMENT: str = "development"
    
    class Config:
        env_file = ".env"
        extra = "allow"

settings = Settings()