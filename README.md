# Spreadly 🚀

**AI-Powered Financial Modeling for Excel** - Generate professional DCF, NPV, and valuation models with context-aware code generation that never overwrites your existing data.

[![Context-Aware](https://img.shields.io/badge/Context-Aware-brightgreen)](#context-awareness)
[![Financial Models](https://img.shields.io/badge/Models-20%2B%20Templates-blue)](#financial-modeling)
[![RAG Enhanced](https://img.shields.io/badge/RAG-Enhanced%20AI-purple)](#rag-integration)

Transform Excel into a powerful financial modeling platform with AI that understands your spreadsheet context and generates professional-grade models incrementally.

---

## ✨ Key Features

### 🎯 **Incremental Financial Model Building**
- **Context-Aware Generation**: AI sees your existing data and builds around it
- **Smart Collision Avoidance**: Never overwrites existing formulas or data
- **Incremental Construction**: Models built chunk-by-chunk with real-time progress
- **Professional Templates**: Based on 20+ curated industry models

### 🧠 **RAG-Enhanced AI**
- **Retrieval-Augmented Generation**: Uses professional model templates
- **Industry Specialization**: Technology, Healthcare, Energy, Finance, and more
- **Complexity Levels**: From basic templates to expert investment-grade models
- **Semantic Search**: Finds the most relevant templates for your use case

### 🛡️ **Smart Placement System**
- **Forbidden Cell Detection**: Creates explicit lists of occupied cells
- **Safe Zone Identification**: Suggests optimal placement areas
- **Range Collision Prevention**: "NEVER use: A1, A3, A4..." guidance
- **Dynamic Layout Adaptation**: Adjusts to your existing spreadsheet structure

### 📊 **Professional Financial Models**
- **DCF Models**: Discounted Cash Flow with terminal value calculations
- **NPV Analysis**: Net Present Value with sensitivity analysis
- **Three-Statement Models**: Integrated P&L, Balance Sheet, and Cash Flow
- **Valuation Models**: Comparable company and precedent transaction analysis
- **LBO Models**: Leveraged buyout with returns analysis

---

## 🔥 Recent Improvements

### ✅ **Context Awareness Fixed** (Latest)
- **Issue Resolved**: AI now properly detects and respects existing data
- **Placement Guidance**: Explicit forbidden cell lists and safe placement zones
- **Data Extraction**: Enhanced reliability with multiple fallback methods
- **Collision Prevention**: "STRICTLY AVOID overwriting existing data" prompts

### ✅ **Token Limit Removed**
- **Issue Resolved**: Removed arbitrary 1500 token limit causing code truncation
- **Full Capacity**: Now uses complete 8192 token capacity for complex models
- **Complete Generation**: No more "Unexpected end of script" errors
- **Quality Improvement**: More sophisticated and complete model generation

### ✅ **Enhanced Data Extraction**
- **Robust Fallbacks**: Multiple methods ensure reliable Excel data access
- **Better Debugging**: Comprehensive logging throughout the pipeline
- **Structured Context**: Proper data formatting from frontend to AI
- **Reliability**: Prioritized simple methods over complex comprehensive ones

---

## 🚀 Quick Start

### Prerequisites
- Microsoft Excel (Office 365 or 2019+)
- Node.js 16+ and Python 3.8+
- Anthropic API key for Claude AI

### 1. Clone and Setup Frontend
```bash
git clone https://github.com/your-org/spreadly.git
cd spreadly/frontend
npm install
npm run dev-server
```

### 2. Setup Backend
```bash
cd ../backend
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure Environment
Create `backend/.env`:
```env
ANTHROPIC_API_KEY=your_claude_api_key
RAG_ENABLED=true
CHROMA_DB_PATH=./chroma_db
MAX_RETRIEVED_MODELS=3
```

### 4. Initialize Model Library
```bash
python -c "
from app.services.model_curator import get_model_curator
import asyncio
curator = get_model_curator()
asyncio.run(curator.initialize_model_library())
"
```

### 5. Start Services
```bash
# Terminal 1: Backend
uvicorn app.main:app --reload

# Terminal 2: Frontend  
npm start
```

---

## 💡 Usage Examples

### Basic Financial Model Generation
```
"Create a DCF model for a SaaS technology company"
```
**Result**: Professional DCF with revenue projections, EBIT calculations, free cash flow, terminal value, and WACC-based valuation.

### Context-Aware Building
```
"Add NPV sensitivity analysis below my existing assumptions"
```
**Result**: AI detects existing data in rows 1-10, places new sensitivity table starting at row 12+.

### Industry-Specific Models
```
"Generate a healthcare drug development NPV model"
```
**Result**: Specialized model with clinical trial phases, regulatory milestones, and risk-adjusted returns.

---

## 🏗️ Architecture

### Frontend (Excel Add-in)
- **TypeScript/React**: Modern Office Add-in development
- **Excel.js API**: Native Excel integration
- **Context Extraction**: Multi-method data reading with fallbacks
- **Incremental Executor**: Chunk-by-chunk model building

### Backend (FastAPI)
- **FastAPI**: High-performance Python web framework
- **Claude 4**: Latest Anthropic AI model (8192 token capacity)
- **ChromaDB**: Local vector storage for model templates
- **RAG Pipeline**: Retrieval-augmented generation system

### Financial AI Components
- **Model Curator**: 20+ professional templates
- **Vector Store**: Semantic search for model retrieval  
- **Incremental Builder**: Context-aware chunk generation
- **Placement Engine**: Smart collision avoidance

```
User Query → Context Detection → RAG Retrieval → AI Generation → Safe Placement → Excel Output
     ↓              ↓               ↓               ↓               ↓              ↓
"DCF model"    Existing data    Templates    Code chunks    Safe zones    Professional model
```

---

## 🎯 Financial Modeling Capabilities

### Supported Model Types
| Model Type | Description | Complexity Levels |
|------------|-------------|-------------------|
| **DCF** | Discounted Cash Flow with terminal value | Basic → Expert |
| **NPV** | Net Present Value with sensitivity | Basic → Advanced |
| **Three-Statement** | Integrated financial statements | Intermediate → Expert |
| **Valuation** | Comparable company analysis | Advanced → Expert |
| **LBO** | Leveraged buyout modeling | Expert |

### Industry Specializations
- **Technology**: SaaS, hardware, semiconductors
- **Healthcare**: Biotech, pharma, medical devices  
- **Energy**: Oil & gas, renewables, utilities
- **Finance**: Banking, insurance, asset management
- **Real Estate**: Development, REITs, commercial

### Context Awareness Features
- **Data Detection**: Automatically identifies existing formulas and data
- **Placement Optimization**: Finds optimal areas for new content
- **Structure Preservation**: Maintains your existing model organization
- **Incremental Building**: Adds components without disrupting current work

---

## 🛠️ Development Setup

### Backend Dependencies
```bash
pip install fastapi uvicorn anthropic
pip install chromadb sentence-transformers
pip install pandas openpyxl xlsxwriter
```

### Environment Variables
```env
# Core API
ANTHROPIC_API_KEY=your_key_here
API_BASE_URL=http://localhost:8000

# RAG Configuration  
RAG_ENABLED=true
CHROMA_DB_PATH=./chroma_db
MAX_RETRIEVED_MODELS=3
VECTOR_STORE_COLLECTION=financial_models

# Model Generation
MODEL_COMPLEXITY_DEFAULT=intermediate
AUTO_DETECT_INDUSTRY=true
INCREMENTAL_CHUNK_SIZE=medium
```

### Model Library Management
```bash
# Check model library status
curl http://localhost:8000/api/models/stats

# Add custom models
curl -X POST "http://localhost:8000/api/models/upload-xlsx" \
  -F "file=@my_dcf.xlsx" \
  -F "model_type=dcf" \
  -F "industry=technology"

# Search available models
curl "http://localhost:8000/api/models/search?query=healthcare%20NPV"
```

---

## 📋 API Endpoints

### Financial Model Generation
- `POST /api/incremental/start` - Initialize model building session
- `POST /api/incremental/next-chunk` - Generate next model component
- `GET /api/incremental/status/{session}` - Check build progress
- `POST /api/incremental/handle-error` - Auto-fix generation errors

### Model Library Management
- `GET /api/models/stats` - Library statistics
- `POST /api/models/upload-xlsx` - Add custom Excel models
- `GET /api/models/search` - Semantic model search
- `GET /api/models/list` - Browse available templates

### Excel Integration  
- `POST /api/excel/analyze` - Analyze existing spreadsheet data
- `POST /api/excel/context` - Extract workbook context
- `POST /api/excel/validate` - Validate generated formulas

---

## 🤝 Contributing

### Adding Financial Models
1. **Create Excel Template**: Professional formatting and formulas
2. **Use Naming Convention**: `ModelType_Industry_Complexity.xlsx`
3. **Upload via API**: Use bulk upload endpoints
4. **Test Integration**: Verify RAG retrieval works correctly

### Development Guidelines
- **Context First**: Always consider existing data impact
- **Incremental Design**: Build models in logical chunks
- **Professional Quality**: Match investment banking standards
- **Error Recovery**: Handle Excel API limitations gracefully

### Model Quality Standards
- ✅ Professional formatting and layout
- ✅ Robust formula construction  
- ✅ Industry-appropriate assumptions
- ✅ Clear documentation and structure
- ✅ Error handling and validation

---

## 📈 Project Structure

```
spreadly/
├── README.md                          # This file
├── backend/                           # FastAPI backend
│   ├── app/
│   │   ├── api/endpoints/            # REST API endpoints
│   │   │   ├── incremental_model.py  # Model building API
│   │   │   └── model_management.py   # Model library API
│   │   ├── services/                 # Business logic
│   │   │   ├── ai_service_simple.py  # Claude AI integration
│   │   │   ├── incremental_model_builder.py # Context-aware building
│   │   │   ├── model_vector_store.py # RAG vector storage
│   │   │   └── model_curator.py      # Template library
│   │   ├── models/                   # Data models
│   │   │   └── financial_model.py    # Model definitions
│   │   └── core/                     # Configuration
│   ├── chroma_db/                    # Vector database storage
│   └── requirements.txt              # Python dependencies
├── frontend/                         # Excel Add-in
│   ├── src/
│   │   ├── taskpane/                # Task pane UI
│   │   ├── utils/                   # Utilities
│   │   │   └── IncrementalExecutor.ts # Model building orchestration
│   │   └── services/                # Excel integration
│   │       └── excel-data.ts        # Context extraction
│   ├── manifest.xml                 # Office Add-in manifest
│   └── package.json                 # Frontend dependencies
├── FINANCIAL_MODELS_SETUP.md        # Model library guide
├── RAG_IMPLEMENTATION_GUIDE.md      # RAG architecture
└── FINANCIAL_MODEL_STRATEGY.md      # Model building strategy
```

---

## 🏆 Why Spreadly?

### 🎯 **Context Intelligence**
Unlike basic code generators, Spreadly **sees** your existing work and builds around it intelligently.

### 🧠 **Professional Quality**  
Models based on real investment banking and corporate finance templates, not generic examples.

### 🛡️ **Safe & Reliable**
Never overwrites your data. Explicit collision avoidance with forbidden cell lists and safe zones.

### 🚀 **Incremental Building**
Watch your model build piece by piece with real-time progress tracking and error recovery.

### 📚 **Continuously Learning**
RAG system means every model generation gets better by learning from a curated library of professional templates.

---

**Transform your Excel workflow from manual formula writing to AI-powered financial modeling. Get started today!** 🚀

*Built with ❤️ for finance professionals who demand precision and efficiency.*