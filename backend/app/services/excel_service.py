import pandas as pd
import openpyxl
from typing import Dict, Any, List
from fastapi import UploadFile
from sqlalchemy.orm import Session
from app.models.spreadsheet import Spreadsheet
from app.core.config import settings
import hashlib
import json
import os

class ExcelService:
    def __init__(self):
        self.upload_dir = settings.UPLOAD_DIR
        os.makedirs(self.upload_dir, exist_ok=True)
    
    async def process_file(self, file: UploadFile, session_id: int) -> Dict[str, Any]:
        """Process uploaded Excel file and extract data"""
        file_path = os.path.join(self.upload_dir, file.filename)
        
        # Save file temporarily
        with open(file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        # Generate file hash
        file_hash = hashlib.md5(content).hexdigest()
        
        try:
            # Load workbook
            workbook = openpyxl.load_workbook(file_path)
            sheet_names = workbook.sheetnames
            
            # Process each sheet
            sheets_data = {}
            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                sheets_data[sheet_name] = self._analyze_sheet(df)
            
            # Create spreadsheet record
            spreadsheet_data = {
                "name": file.filename,
                "file_hash": file_hash,
                "sheet_names": sheet_names,
                "sheets_analysis": sheets_data,
                "summary": self._generate_summary(sheets_data)
            }
            
            return spreadsheet_data
            
        except Exception as e:
            raise Exception(f"Error processing Excel file: {str(e)}")
        finally:
            # Clean up temporary file
            if os.path.exists(file_path):
                os.remove(file_path)
    
    def _analyze_sheet(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Analyze individual sheet data"""
        analysis = {
            "shape": df.shape,
            "columns": df.columns.tolist(),
            "data_types": df.dtypes.to_dict(),
            "missing_values": df.isnull().sum().to_dict(),
            "summary_stats": {}
        }
        
        # Generate summary statistics for numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            analysis["summary_stats"] = df[numeric_cols].describe().to_dict()
        
        # Sample data (first 5 rows)
        analysis["sample_data"] = df.head().to_dict('records')
        
        return analysis
    
    def _generate_summary(self, sheets_data: Dict[str, Any]) -> str:
        """Generate a summary of the Excel file"""
        total_sheets = len(sheets_data)
        total_rows = sum(data["shape"][0] for data in sheets_data.values())
        total_cols = sum(data["shape"][1] for data in sheets_data.values())
        
        summary = f"Excel file contains {total_sheets} sheets with {total_rows} total rows and {total_cols} total columns. "
        
        for sheet_name, data in sheets_data.items():
            summary += f"Sheet '{sheet_name}' has {data['shape'][0]} rows and {data['shape'][1]} columns. "
        
        return summary
    
    async def process_data(self, data: List[List], session_id: int, file_name: str) -> Dict[str, Any]:
        """Process Excel data from frontend (array of arrays)"""
        try:
            # Convert data to pandas DataFrame
            df = pd.DataFrame(data)
            
            # Generate file hash from data
            data_string = str(data)
            file_hash = hashlib.md5(data_string.encode()).hexdigest()
            
            # Analyze the data
            analysis = self._analyze_sheet(df)
            
            # Create spreadsheet data
            spreadsheet_data = {
                "name": file_name,
                "file_hash": file_hash,
                "sheet_names": ["Sheet1"],
                "sheets_analysis": {"Sheet1": analysis},
                "summary": self._generate_summary_from_analysis({"Sheet1": analysis})
            }
            
            return spreadsheet_data
            
        except Exception as e:
            raise Exception(f"Error processing Excel data: {str(e)}")
    
    def _generate_summary_from_analysis(self, sheets_data: Dict[str, Any]) -> str:
        """Generate a summary from analysis data"""
        total_sheets = len(sheets_data)
        total_rows = sum(data["shape"][0] for data in sheets_data.values())
        total_cols = sum(data["shape"][1] for data in sheets_data.values())
        
        summary = f"Excel data contains {total_sheets} sheets with {total_rows} total rows and {total_cols} total columns. "
        
        for sheet_name, data in sheets_data.items():
            summary += f"Sheet '{sheet_name}' has {data['shape'][0]} rows and {data['shape'][1]} columns. "
        
        return summary