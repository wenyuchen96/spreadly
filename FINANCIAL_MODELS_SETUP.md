# Adding Your Financial Model Collection to RAG

## Overview

Your financial models will be stored as **vector embeddings** in ChromaDB for semantic search. The system converts models to Excel.js JavaScript code rather than storing raw XLSX/XML files.

## Three Ways to Add Your Models

### 🎯 Option 1: API Upload (Recommended for End Users)

**Single File Upload:**
```bash
curl -X POST "http://localhost:8000/api/models/upload-xlsx" \
  -F "file=@my_dcf_model.xlsx" \
  -F "model_type=dcf" \
  -F "industry=technology" \
  -F "complexity=advanced" \
  -F "name=Tech Company DCF Model" \
  -F "description=Professional DCF model for SaaS companies"
```

**Bulk Upload:**
```bash
curl -X POST "http://localhost:8000/api/models/bulk-upload" \
  -F "files=@model1.xlsx" \
  -F "files=@model2.xlsx" \
  -F "files=@model3.xlsx" \
  -F "auto_detect=true"
```

### 🛠️ Option 2: Command Line Tools

**Convert XLSX to Models:**
```python
from tools.xlsx_to_model_converter import convert_excel_collection

# Convert all XLSX files in a directory
models = convert_excel_collection("./your_excel_models/")
print(f"Converted {len(models)} models")
```

**Bulk Load into Vector Store:**
```python
from tools.bulk_model_loader import BulkModelLoader
import asyncio

async def load_models():
    loader = BulkModelLoader()
    
    # Load from XLSX directory
    results = await loader.load_from_xlsx_directory(
        "./your_excel_models/", 
        auto_detect=True
    )
    
    print(f"Loaded {results['successful']} models successfully")
    print(f"Failed: {results['failed']}")

asyncio.run(load_models())
```

### 📄 Option 3: Pre-defined JSON Models

Create a JSON file with your models:

```json
[
  {
    "id": "my_custom_dcf_001",
    "name": "Custom Technology DCF",
    "description": "DCF model for high-growth technology companies",
    "model_type": "dcf",
    "industry": "technology", 
    "complexity": "advanced",
    "excel_code": "await Excel.run(async (context) => { ... });",
    "business_description": "Professional DCF for tech companies",
    "sample_inputs": {
      "revenue": 100000000,
      "growth_rate": 0.25
    },
    "expected_outputs": {
      "enterprise_value": "calculated_ev"
    },
    "metadata": {
      "components": ["revenue_projections", "dcf_calculation"],
      "excel_functions": ["NPV", "IRR"],
      "formatting_features": ["professional_layout"],
      "business_assumptions": ["high_growth", "scalable_business"],
      "time_horizon_years": 10,
      "currencies": ["USD"],
      "regions": ["north_america"]
    },
    "performance": {
      "execution_success_rate": 0.95,
      "user_rating": 4.8,
      "usage_count": 0,
      "error_count": 0,
      "modification_frequency": 0.1
    },
    "created_by": "your_organization",
    "keywords": ["dcf", "technology", "high_growth", "saas"],
    "tags": ["custom", "professional", "tech_focused"]
  }
]
```

Then load:
```python
from tools.bulk_model_loader import BulkModelLoader
results = await loader.load_from_json_models("./my_models.json")
```

## File Naming Conventions (for Auto-Detection)

The system can auto-detect model types from filenames:

**Model Types:**
- `DCF_Tech_Company.xlsx` → DCF model
- `NPV_Project_Analysis.xlsx` → NPV model  
- `LBO_Healthcare_Deal.xlsx` → LBO model
- `Budget_2024_Forecast.xlsx` → Budget model
- `Comps_Valuation.xlsx` → Valuation model

**Industries:**
- `DCF_Tech_SaaS.xlsx` → Technology industry
- `NPV_Healthcare_Drug.xlsx` → Healthcare industry
- `Budget_Energy_Renewable.xlsx` → Energy industry

**Complexity:**
- `DCF_Basic_Template.xlsx` → Basic complexity
- `DCF_Advanced_Model.xlsx` → Advanced complexity
- `DCF_Expert_Investment_Grade.xlsx` → Expert complexity

## Management API Endpoints

### Check Current Models
```bash
# Get statistics
curl "http://localhost:8000/api/models/stats"

# List all models  
curl "http://localhost:8000/api/models/list"

# Search models
curl "http://localhost:8000/api/models/search?query=DCF%20technology&limit=5"
```

### Delete Models
```bash
curl -X DELETE "http://localhost:8000/api/models/my_model_id"
```

## Vector Store Technical Details

**Storage**: ChromaDB (local, no external dependencies)
**Embeddings**: SentenceTransformers (all-MiniLM-L6-v2)
**Search**: Semantic similarity + metadata filtering
**Location**: `./chroma_db/` directory

## Model Conversion Process

1. **XLSX Analysis**: Extracts structure, formulas, formatting
2. **Code Generation**: Converts to Excel.js JavaScript
3. **Metadata Extraction**: Identifies components, functions, patterns  
4. **Vector Embedding**: Creates searchable embeddings
5. **Storage**: Saves to ChromaDB with metadata

## Example: Adding Your DCF Collection

```python
# 1. Organize your files
mkdir /path/to/your/models/dcf_models
# Put your XLSX files there with descriptive names

# 2. Bulk load
from tools.bulk_model_loader import BulkModelLoader
import asyncio

async def load_dcf_collection():
    loader = BulkModelLoader()
    results = await loader.load_from_xlsx_directory(
        "/path/to/your/models/dcf_models/",
        auto_detect=True
    )
    
    print("📊 Results:")
    print(f"  ✅ Successful: {results['successful']}")
    print(f"  ❌ Failed: {results['failed']}")
    
    for model in results['loaded_models']:
        print(f"  📝 {model['file']} → {model['type']} ({model['industry']})")

asyncio.run(load_dcf_collection())
```

After loading, your models will be available for RAG-enhanced queries like:
- "Create a DCF model for a SaaS company" 
- "Generate a technology valuation model"
- "Build a healthcare NPV analysis"

The system will automatically retrieve and use your professional templates to enhance the generated models! 🚀