#!/usr/bin/env python3
"""
Tool to bulk load financial models into the RAG vector store
"""

import asyncio
import json
from pathlib import Path
from typing import List, Dict, Any
import sys
import os

# Add parent directory to path to import app modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.models.financial_model import FinancialModel, ModelType, Industry, ComplexityLevel
from app.services.model_vector_store import get_vector_store
from xlsx_to_model_converter import XLSXToModelConverter

class BulkModelLoader:
    """Load financial models into the RAG vector store from various sources"""
    
    def __init__(self):
        self.vector_store = get_vector_store()
        self.xlsx_converter = XLSXToModelConverter()
    
    async def load_from_xlsx_directory(self, directory_path: str, auto_detect: bool = True) -> Dict[str, Any]:
        """
        Load all XLSX files from a directory into the vector store
        
        Args:
            directory_path: Path to directory containing XLSX files
            auto_detect: Whether to auto-detect model types from filenames/content
        """
        
        results = {
            "total_processed": 0,
            "successful": 0,
            "failed": 0,
            "errors": [],
            "loaded_models": []
        }
        
        directory = Path(directory_path)
        if not directory.exists():
            raise ValueError(f"Directory does not exist: {directory_path}")
        
        xlsx_files = list(directory.glob("*.xlsx")) + list(directory.glob("*.xls"))
        results["total_processed"] = len(xlsx_files)
        
        print(f"📁 Found {len(xlsx_files)} Excel files in {directory_path}")
        
        for xlsx_file in xlsx_files:
            try:
                print(f"🔄 Processing: {xlsx_file.name}")
                
                # Auto-detect model characteristics from filename
                model_type, industry, complexity = self._detect_from_filename(xlsx_file.name) if auto_detect else (ModelType.DCF, Industry.GENERAL, ComplexityLevel.INTERMEDIATE)
                
                # Convert XLSX to FinancialModel
                model = self.xlsx_converter.convert_xlsx_to_model(
                    str(xlsx_file),
                    f"uploaded_{xlsx_file.stem}_{model_type}",
                    model_type,
                    industry,
                    complexity
                )
                
                # Add to vector store
                success = await self.vector_store.add_model(model)
                
                if success:
                    results["successful"] += 1
                    results["loaded_models"].append({
                        "file": xlsx_file.name,
                        "model_id": model.id,
                        "type": model_type,
                        "industry": industry
                    })
                    print(f"✅ Loaded: {xlsx_file.name} as {model_type} model")
                else:
                    results["failed"] += 1
                    results["errors"].append(f"Vector store failed for {xlsx_file.name}")
                    print(f"❌ Vector store failed: {xlsx_file.name}")
                
            except Exception as e:
                results["failed"] += 1
                error_msg = f"Failed to process {xlsx_file.name}: {str(e)}"
                results["errors"].append(error_msg)
                print(f"❌ {error_msg}")
        
        return results
    
    async def load_from_json_models(self, json_file_path: str) -> Dict[str, Any]:
        """
        Load pre-defined models from a JSON file
        
        JSON format:
        [
            {
                "id": "custom_dcf_001",
                "name": "Custom DCF Model",
                "description": "...",
                "model_type": "dcf",
                "industry": "technology",
                "complexity": "advanced",
                "excel_code": "await Excel.run(async (context) => { ... });",
                "keywords": ["dcf", "valuation"],
                ...
            }
        ]
        """
        
        results = {
            "total_processed": 0,
            "successful": 0,
            "failed": 0,
            "errors": [],
            "loaded_models": []
        }
        
        with open(json_file_path, 'r') as f:
            models_data = json.load(f)
        
        results["total_processed"] = len(models_data)
        
        for model_data in models_data:
            try:
                # Create FinancialModel from JSON data
                model = FinancialModel(**model_data)
                
                # Add to vector store
                success = await self.vector_store.add_model(model)
                
                if success:
                    results["successful"] += 1
                    results["loaded_models"].append({
                        "model_id": model.id,
                        "name": model.name,
                        "type": model.model_type
                    })
                    print(f"✅ Loaded: {model.name}")
                else:
                    results["failed"] += 1
                    results["errors"].append(f"Vector store failed for {model.id}")
                
            except Exception as e:
                results["failed"] += 1
                error_msg = f"Failed to load model {model_data.get('id', 'unknown')}: {str(e)}"
                results["errors"].append(error_msg)
                print(f"❌ {error_msg}")
        
        return results
    
    def _detect_from_filename(self, filename: str) -> tuple[ModelType, Industry, ComplexityLevel]:
        """Auto-detect model characteristics from filename"""
        filename_lower = filename.lower()
        
        # Detect model type
        if any(keyword in filename_lower for keyword in ['dcf', 'discounted', 'enterprise']):
            model_type = ModelType.DCF
        elif any(keyword in filename_lower for keyword in ['npv', 'project', 'investment']):
            model_type = ModelType.NPV
        elif any(keyword in filename_lower for keyword in ['lbo', 'leveraged', 'buyout']):
            model_type = ModelType.LBO
        elif any(keyword in filename_lower for keyword in ['budget', 'forecast', 'planning']):
            model_type = ModelType.BUDGET
        elif any(keyword in filename_lower for keyword in ['valuation', 'comps', 'comparable']):
            model_type = ModelType.VALUATION
        else:
            model_type = ModelType.DCF  # Default
        
        # Detect industry
        if any(keyword in filename_lower for keyword in ['tech', 'software', 'saas']):
            industry = Industry.TECHNOLOGY
        elif any(keyword in filename_lower for keyword in ['healthcare', 'pharma', 'medical']):
            industry = Industry.HEALTHCARE
        elif any(keyword in filename_lower for keyword in ['energy', 'oil', 'renewable']):
            industry = Industry.ENERGY
        elif any(keyword in filename_lower for keyword in ['retail', 'consumer']):
            industry = Industry.RETAIL
        else:
            industry = Industry.GENERAL
        
        # Detect complexity
        if any(keyword in filename_lower for keyword in ['basic', 'simple', 'beginner']):
            complexity = ComplexityLevel.BASIC
        elif any(keyword in filename_lower for keyword in ['advanced', 'complex', 'professional']):
            complexity = ComplexityLevel.ADVANCED
        elif any(keyword in filename_lower for keyword in ['expert', 'investment_grade']):
            complexity = ComplexityLevel.EXPERT
        else:
            complexity = ComplexityLevel.INTERMEDIATE
        
        return model_type, industry, complexity
    
    async def get_vector_store_stats(self) -> Dict[str, Any]:
        """Get current vector store statistics"""
        if not self.vector_store.is_available():
            return {"error": "Vector store not available"}
        
        return self.vector_store.get_stats()

# CLI interface
async def main():
    loader = BulkModelLoader()
    
    print("🚀 Financial Model Bulk Loader")
    print("=" * 50)
    
    # Show current stats
    print("📊 Current Vector Store Stats:")
    stats = await loader.get_vector_store_stats()
    print(json.dumps(stats, indent=2))
    print()
    
    # Example usage - you would modify these paths
    
    # Option 1: Load from XLSX directory
    # results = await loader.load_from_xlsx_directory("./sample_models/", auto_detect=True)
    
    # Option 2: Load from JSON file
    # results = await loader.load_from_json_models("./model_definitions.json")
    
    print("💡 Usage Examples:")
    print("1. Load XLSX files: await loader.load_from_xlsx_directory('./excel_models/')")
    print("2. Load JSON models: await loader.load_from_json_models('./models.json')")
    print("3. Check stats: await loader.get_vector_store_stats()")

if __name__ == "__main__":
    asyncio.run(main())