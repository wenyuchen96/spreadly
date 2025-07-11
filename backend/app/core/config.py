from pydantic_settings import BaseSettings
from typing import List
import json

class Settings(BaseSettings):
    # Database
    DATABASE_URL: str = "sqlite:///./spreadly.db"
    
    # Redis
    REDIS_URL: str = "redis://localhost:6379"
    
    # API Keys
    ANTHROPIC_API_KEY: str
    PINECONE_API_KEY: str
    PINECONE_ENVIRONMENT: str
    PINECONE_INDEX_NAME: str
    
    # Security
    SECRET_KEY: str
    ALGORITHM: str
    ACCESS_TOKEN_EXPIRE_MINUTES: int
    
    # CORS
    ALLOWED_HOSTS: str  # Will be parsed from .env as JSON string
    
    # File Upload
    MAX_FILE_SIZE: int
    UPLOAD_DIR: str
    
    # Environment
    ENVIRONMENT: str
    
    class Config:
        env_file = ".env"
        extra = "allow"
    
    @property
    def allowed_hosts_list(self) -> List[str]:
        """Parse ALLOWED_HOSTS from JSON string in .env"""
        try:
            return json.loads(self.ALLOWED_HOSTS)
        except (json.JSONDecodeError, TypeError):
            # Fallback to default values
            return [
                "http://localhost:3000", 
                "https://localhost:3000", 
                "https://localhost:3001",
                "https://excel.officeapps.live.com",
                "https://excel.office.com", 
                "https://office.live.com",
                "*"
            ]

settings = Settings()