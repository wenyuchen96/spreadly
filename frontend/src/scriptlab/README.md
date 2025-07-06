# Script Lab Engine Integration

This module integrates Script Lab's execution engine for runtime TypeScript/JavaScript code generation and execution within Excel.

## Overview

The Script Lab engine allows AI agents to generate and execute Excel automation code dynamically through a chat interface.

## Components

### Core Classes

- **ScriptLabEngine**: Main execution engine for running TypeScript/JavaScript code
- **TypeScriptCompiler**: Handles TypeScript compilation to JavaScript
- **OfficeJsHost**: Office.js integration and host detection utilities

### Interfaces

- **ISnippet**: Defines code snippet structure
- **IExecutionResult**: Execution result format
- **IScriptLabEngineOptions**: Engine configuration options

## Features

- ✅ TypeScript compilation and execution
- ✅ Office.js API integration
- ✅ Sandboxed code execution in iframe
- ✅ Error handling and validation
- ✅ Built-in Excel operation helpers
- ✅ Chat interface integration

## Usage

### Basic Usage

```typescript
import { ScriptLabEngine } from './scriptlab';

const engine = new ScriptLabEngine();

// Create a snippet
const snippet = engine.createSnippet(
  'Highlight Cells',
  `
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    await context.sync();
});
  `
);

// Execute the snippet
const result = await engine.executeSnippet(snippet);
console.log(result);
```

### Chat Interface Integration

The engine is integrated with the chat interface and responds to natural language commands:

- **"highlight cells"** - Highlights selected cells
- **"insert data [[1,2],[3,4]]"** - Inserts data into spreadsheet
- **"create chart A1:C5"** - Creates chart from range
- **"format cells A1:B2"** - Formats cell range
- **"test"** - Runs demonstration code

### Excel Operations Helpers

```typescript
import { ExcelOperations } from './scriptlab';

// Highlight cells
const highlightSnippet = ExcelOperations.createHighlightCellsSnippet('yellow');

// Insert data
const dataSnippet = ExcelOperations.createInsertDataSnippet([
  ['Name', 'Age'],
  ['Alice', 25],
  ['Bob', 30]
]);

// Create chart
const chartSnippet = ExcelOperations.createChartSnippet('A1:B3', 'ColumnClustered');
```

## Testing

### Manual Testing

1. Open Excel with the add-in loaded
2. Use the chat interface with test commands:
   ```
   highlight cells yellow
   insert data [[1,2,3],[4,5,6]]
   create chart A1:C5
   test
   ```

### Programmatic Testing

```typescript
import { testScriptLabEngine } from './scriptlab/test';

// Run comprehensive tests
await testScriptLabEngine();
```

### Console Testing

In the browser console (when add-in is loaded):

```javascript
// Run full test suite
await testScriptLab();

// Get test commands
console.log(testCommands);

// Run specific test command
getTestCommand('highlight'); // Returns: "highlight cells yellow"
```

## Architecture

```
Chat Interface → Message Processor → Script Lab Engine → TypeScript Compiler → Office.js API
                                                     ↓
                                                   Iframe Execution Environment
```

### Security

- Code execution happens in sandboxed iframe
- TypeScript compilation and validation before execution
- Timeout protection (30 seconds default)
- Error boundary and exception handling

## Configuration

```typescript
const engine = new ScriptLabEngine({
  timeout: 30000,           // Execution timeout in ms
  sandboxMode: true,        // Enable iframe sandbox
  compilerOptions: {        // TypeScript compiler options
    target: 'ES2017',
    strict: false
  }
});
```

## Error Handling

The engine provides comprehensive error handling:

- **Syntax Validation**: Pre-execution TypeScript syntax checking
- **Compilation Errors**: TypeScript compilation error reporting
- **Runtime Errors**: Exception catching during execution
- **Timeout Protection**: Automatic timeout for long-running code

## Future Enhancements

- [ ] AI/LLM integration for intelligent code generation
- [ ] Code optimization and caching
- [ ] Advanced Excel API pattern recognition
- [ ] Multi-file snippet support
- [ ] Debugging and step-through capabilities
- [ ] Code completion and IntelliSense integration

## Dependencies

- `typescript`: TypeScript compiler
- `@types/office-js`: Office.js type definitions
- `monaco-editor`: Code editor (optional)

## License

This integration maintains compatibility with the original Script Lab license while being part of the Spreadly project.