# Spreadly

A Microsoft Office Add-in for Excel that provides enhanced spreadsheet functionality and task pane integration.

## Project Structure

```
spreadly/
├── README.md
├── backend/          # FastAPI backend services
│   ├── app/          # FastAPI application
│   │   ├── api/      # API endpoints
│   │   ├── core/     # Core configuration
│   │   ├── models/   # Database models
│   │   └── services/ # Business logic
│   ├── requirements.txt
│   └── .env          # Environment variables
├── frontend/         # Office Add-in frontend
    ├── assets/       # Icons and static assets
    ├── src/          # TypeScript source code
    │   ├── commands/ # Office commands
    │   └── taskpane/ # Task pane implementation
    ├── manifest.xml  # Office Add-in manifest
    └── package.json  # Frontend dependencies
```

## Getting Started

### Prerequisites
- Node.js (v14 or higher)
- npm
- Microsoft Office (Excel)

### Installation
1. Clone the repository
2. Install frontend dependencies:
   ```bash
   cd frontend
   npm install
   ```

### Development
- Start development server: `npm run dev-server`
- Build for production: `npm run build`
- Start debugging: `npm start`

### Scripts
- `npm run build` - Build for production
- `npm run dev-server` - Start development server
- `npm run lint` - Run ESLint
- `npm run validate` - Validate manifest
- `npm start` - Start Office debugging
- `npm stop` - Stop Office debugging

## Features
- Excel task pane integration
- Custom commands and functions
- Modern TypeScript development with Webpack
- AI-powered spreadsheet analysis and insights
- Natural language queries on Excel data
- Automated data processing and formula generation

## Backend Integration

### Technology Stack
- **FastAPI**: High-performance Python web framework
- **LangChain**: AI agent framework for spreadsheet operations
- **Anthropic Claude**: Primary AI model for natural language processing
- **PostgreSQL**: Primary database for user data and sessions
- **Redis**: Caching and real-time data storage
- **Pinecone**: Vector database for spreadsheet pattern search
- **Celery**: Background task processing

### Backend Setup
1. Navigate to backend directory:
   ```bash
   cd backend
   ```

2. Create virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Set environment variables in `.env`:
   ```
   ANTHROPIC_API_KEY=your_anthropic_key
   DATABASE_URL=postgresql://user:password@localhost/spreadly
   REDIS_URL=redis://localhost:6379
   PINECONE_API_KEY=your_pinecone_key
   ```

5. Run database migrations:
   ```bash
   alembic upgrade head
   ```

6. Start the FastAPI server:
   ```bash
   uvicorn app.main:app --reload
   ```

### API Endpoints
- `POST /api/excel/upload` - Upload Excel files for processing
- `GET /api/excel/analyze` - Get AI-powered data insights
- `POST /api/excel/query` - Natural language queries on spreadsheet data
- `GET /api/excel/formulas` - Generate Excel formulas from descriptions
- `POST /api/excel/search` - Vector search for similar spreadsheet patterns

### Database Schema
- **Users**: User profiles and authentication
- **Sessions**: Excel processing sessions
- **Spreadsheets**: Metadata and analysis results
- **Patterns**: Cached AI-generated insights and formulas
