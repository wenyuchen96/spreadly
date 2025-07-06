# Spreadly

A Microsoft Office Add-in for Excel that provides enhanced spreadsheet functionality and task pane integration.

## Project Structure

```
spreadly/
├── README.md
├── backend/          # Backend services and API
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
