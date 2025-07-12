/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// import { ExcelOperations } from '../scriptlab'; // Unused
import { SimpleScriptLabEngine } from '../scriptlab/SimpleEngine';
import { spreadlyAPI } from '../services/api';
// Removed unused imports
import { getSelectedRangeData } from '../services/excel-data';
import { BASE_URL, getApiUrl } from '../config/api-config';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    initializeChat();
  }
});

function initializeChat() {
  const chatInput = document.getElementById("chat-input") as HTMLTextAreaElement;
  const sendButton = document.getElementById("send-button") as HTMLButtonElement;
  const chatMessages = document.getElementById("chat-messages") as HTMLDivElement;
  // Use SimpleEngine for testing (fallback without TypeScript compilation issues)
  const scriptLabEngine = new SimpleScriptLabEngine();

  // Auto-resize textarea
  chatInput.addEventListener("input", () => {
    chatInput.style.height = "auto";
    chatInput.style.height = Math.min(chatInput.scrollHeight, 100) + "px";
    
    // Enable/disable send button based on input
    sendButton.disabled = chatInput.value.trim() === "";
  });

  // Send message on Enter (but not Shift+Enter)
  chatInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (!sendButton.disabled) {
        sendMessage();
      }
    }
  });

  // Send button click
  sendButton.addEventListener("click", sendMessage);

  async function sendMessage() {
    const message = chatInput.value.trim();
    if (!message) return;

    // Add user message
    addMessage(message, "user");
    
    // Clear input
    chatInput.value = "";
    chatInput.style.height = "auto";
    sendButton.disabled = true;

    // Show typing indicator
    addMessage("Processing your request...", "assistant", true);

    try {
      // Process the user message and generate appropriate response
      const response = await processUserMessage(message, scriptLabEngine);
      
      // Remove typing indicator
      removeLastMessage();
      
      // Add actual response
      addMessage(response.message, "assistant");
      
      // Execute code if generated
      console.log('🔍 Frontend: Checking for code execution...');
      console.log('🔍 Frontend: response.code exists:', !!response.code, 'response.execute:', response.execute);
      
      if (response.code && response.execute) {
        console.log('🔍 Frontend: Code found, executing...');
        console.log('🔍 Frontend: Code preview:', response.code.substring(0, 100) + '...');
        addMessage("🔄 Executing financial model in Excel...", "assistant", true);
        const result = await executeGeneratedCode(response.code, scriptLabEngine);
        removeLastMessage();
        addMessage(result, "assistant");
      } else {
        console.log('🔍 Frontend: No code to execute');
      }
    } catch (error) {
      removeLastMessage();
      addMessage(`❌ Error: ${error instanceof Error ? error.message : 'Unknown error'}`, "assistant");
    }
  }

  function addMessage(text: string, sender: "user" | "assistant", isTemporary: boolean = false) {
    const messageDiv = document.createElement("div");
    messageDiv.className = `chat-message ${sender}-message`;
    if (isTemporary) {
      messageDiv.setAttribute("data-temporary", "true");
    }

    const avatarDiv = document.createElement("div");
    avatarDiv.className = "message-avatar";
    
    const avatarIcon = document.createElement("i");
    avatarIcon.className = sender === "user" 
      ? "ms-Icon ms-Icon--Contact ms-font-m"
      : "ms-Icon ms-Icon--Robot ms-font-m";
    avatarDiv.appendChild(avatarIcon);

    const contentDiv = document.createElement("div");
    contentDiv.className = "message-content";
    
    const textDiv = document.createElement("div");
    textDiv.className = "message-text";
    textDiv.textContent = text;
    contentDiv.appendChild(textDiv);

    messageDiv.appendChild(avatarDiv);
    messageDiv.appendChild(contentDiv);

    chatMessages.appendChild(messageDiv);
    
    // Scroll to bottom
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  function removeLastMessage() {
    const lastMessage = chatMessages.querySelector('[data-temporary="true"]');
    if (lastMessage) {
      lastMessage.remove();
    }
  }
}

async function processUserMessage(message: string, engine: SimpleScriptLabEngine): Promise<{ message: string; code?: string; execute?: boolean }> {
  // Direct AI chat via backend
  return await chatWithAI(message, engine);
}


// Test backend connection
async function testBackendConnection(): Promise<{ message: string }> {
  try {
    const response = await fetch(getApiUrl('health'), {
      method: 'GET',
      mode: 'cors',
      headers: {
        'ngrok-skip-browser-warning': 'true'
      }
    });
    
    if (response.ok) {
      const data = await response.json();
      return { message: `✅ **Connected to AI Backend!**\n\n${data.message || 'Backend is healthy'}` };
    } else {
      return { message: `❌ Backend responded with HTTP ${response.status}` };
    }
  } catch (error) {
    return { message: `❌ Connection failed: ${error instanceof Error ? error.message : 'Unknown error'}\n\nMake sure backend is running at ${BASE_URL}` };
  }
}


// Direct AI chat with backend
async function chatWithAI(message: string, _engine: SimpleScriptLabEngine): Promise<{ message: string; code?: string; execute?: boolean }> {
  try {
    // First, get or create a session token
    let sessionToken = spreadlyAPI.getSessionToken();
    
    if (!sessionToken) {
      // Create a session by uploading some sample data
      try {
        const excelData = await getSelectedRangeData();
        const uploadResponse = await spreadlyAPI.uploadExcelData(excelData.data, 'ChatSession');
        sessionToken = uploadResponse.session_token;
      } catch {
        // No data selected, create session with minimal data
        const uploadResponse = await spreadlyAPI.uploadExcelData([['Chat', 'Session']], 'DirectChat');
        sessionToken = uploadResponse.session_token;
      }
    }
    
    // Use the query endpoint for direct AI conversation
    const response = await fetch(`${BASE_URL}/api/excel/query`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'ngrok-skip-browser-warning': 'true'
      },
      body: JSON.stringify({
        session_token: sessionToken,
        query: message
      }),
      mode: 'cors'
    });
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    
    const data = await response.json();
    
    // Extract AI response
    let aiMessage = '';
    let generatedCode = null;
    let executeCode = false;
    
    // Check if response is raw JavaScript code (for financial models)
    if (typeof data.result === 'string' && data.result.includes('Excel.run')) {
      console.log('🔍 Frontend: Detected raw JavaScript financial model code');
      generatedCode = data.result;
      executeCode = true;
      aiMessage = '✅ **Financial model generated successfully!**\n\nThe model code has been created and will be executed in your Excel spreadsheet.';
    } else if (data.result) {
      if (data.result.answer) {
        aiMessage = data.result.answer;
      }
      
      if (data.result.formula) {
        aiMessage += `\n\n**Formula:** \`${data.result.formula}\`\n`;
      }
      
      if (data.result.explanation) {
        aiMessage += `\n**Explanation:** ${data.result.explanation}\n`;
      }
      
      if (data.result.code) {
        generatedCode = data.result.code;
      }
      
      // Rely on the backend to tell us when to execute code.
      executeCode = !!data.result.execute_code;
      
      if (data.result.next_steps && data.result.next_steps.length > 0) {
        aiMessage += `\n**Next Steps:**\n`;
        data.result.next_steps.forEach((step: string) => {
          aiMessage += `• ${step}\n`;
        });
      }
    } else {
      aiMessage = data.answer || JSON.stringify(data, null, 2);
    }
    
    return { 
      message: aiMessage || '🤖 AI response received but was empty',
      code: generatedCode,
      execute: executeCode
    };
    
  } catch (error) {
    console.error('chatWithAI error details:', error);
    const errorMsg = error instanceof Error ? error.message : 'Unknown error';
    console.error('Error message:', errorMsg);
    console.error('Error name:', error instanceof Error ? error.name : 'Unknown');
    return { message: `❌ AI chat failed: ${errorMsg}\n\nBackend URL: ${BASE_URL}\nCheck browser console (F12) for details.` };
  }
}

// Removed processWithMockBackend - no longer needed with direct AI integration

// Core code execution function
async function executeGeneratedCode(code: string, engine: SimpleScriptLabEngine): Promise<string> {
  try {
    // Handle all direct Excel operations
    if (code.startsWith("DIRECT_")) {
      return await executeDirectOperation(code);
    }
    
    // Handle Excel formulas
    if (code.includes("=") && !code.includes("function") && !code.includes("await")) {
      return await insertFormulaToSelectedCell(code);
    }
    
    // Handle JavaScript/TypeScript code for financial models
    if (code.includes("Excel.run") || code.includes("context.workbook") || code.includes("worksheet")) {
      return await executeExcelScriptCode(code, engine);
    }
    
    // Handle general JavaScript code
    return await executeGeneralCode(code, engine);
    
  } catch (error) {
    return `❌ Error executing code: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

// Execute Excel script code (for financial models, data manipulation, etc.)
async function executeExcelScriptCode(code: string, engine: SimpleScriptLabEngine): Promise<string> {
  try {
    console.log('🔍 Attempting Script Lab execution...');
    
    // Pre-execution validation
    const validationResult = validateGeneratedCode(code);
    if (!validationResult.isValid) {
      console.warn('🚨 Code validation failed:', validationResult.errors);
      return `❌ **Code validation failed**\n\n${validationResult.errors.join('\n')}\n\nTry asking for the model again to get different code.`;
    }
    
    // Try Script Lab first
    let wrappedCode = code;
    if (!code.includes("Excel.run")) {
      wrappedCode = `
await Excel.run(async (context) => {
  ${code}
  await context.sync();
});
`;
    }
    
    const snippet = engine.createSnippet(
      'AI Generated Financial Model',
      wrappedCode,
      '<div id="output">Creating financial model in your spreadsheet...</div>',
      'body { padding: 10px; font-family: monospace; color: #2d3748; }'
    );
    
    const result = await engine.executeSnippet(snippet);
    
    if (result.success) {
      return `✅ **Financial model executed successfully!**\n\n${result.result || 'Excel operations completed.'}`;
    } else {
      console.log('🔍 Script Lab failed, trying direct execution...');
      return await executeDirectExcel(code);
    }
  } catch (error) {
    console.log('🔍 Script Lab error, trying direct execution...', error);
    return await executeDirectExcel(code);
  }
}

// Validate generated code before execution
function validateGeneratedCode(code: string): { isValid: boolean; errors: string[] } {
  const errors: string[] = [];
  
  // Check for unsupported APIs
  const unsupportedAPIs = [
    'getCell(',
    'borders.setItem',
    'setItem(',
    'border.style',
    'outline.',
    'Table.'  // Tables can be problematic in web Excel
  ];
  
  unsupportedAPIs.forEach(api => {
    if (code.includes(api)) {
      errors.push(`❌ Unsupported API detected: ${api} (not available in web Excel)`);
    }
  });
  
  // Check for proper Excel.run wrapper
  if (!code.includes('Excel.run') && !code.includes('context.workbook')) {
    errors.push(`⚠️ Code should use Excel.run() wrapper for proper execution`);
  }
  
  // Check for potential dimension mismatches in array assignments
  const rangeValuePattern = /getRange\("([^"]+)"\)\.values\s*=\s*(\[.*?\])/g;
  let match: RegExpExecArray | null;
  while ((match = rangeValuePattern.exec(code)) !== null) {
    const range = match[1];
    const arrayStr = match[2];
    
    // Simple dimension check for common patterns
    if (range.includes(':')) {
      const [start, end] = range.split(':');
      if (start.length === 2 && end.length === 2) { // Like A1:C1
        const colDiff = end.charCodeAt(0) - start.charCodeAt(0) + 1;
        const rowDiff = parseInt(end.slice(1)) - parseInt(start.slice(1)) + 1;
        
        // Count array dimensions (simple check)
        const outerArrays = (arrayStr.match(/\[/g) || []).length - 1;
        if (outerArrays > 0 && (colDiff > 1 || rowDiff > 1)) {
          // This is a multi-cell range, warn about potential dimension issues
          console.warn(`🔍 Dimension check: Range ${range} expects ${rowDiff}x${colDiff}, found ${outerArrays} array levels`);
        }
      }
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors
  };
}

// Fallback: Direct Excel execution for web Excel compatibility
async function executeDirectExcel(code: string): Promise<string> {
  try {
    console.log('🔍 Executing code directly in Excel...');
    
    // Execute the code directly without iframe
    const wrappedCode = code.includes("Excel.run") ? code : `
await Excel.run(async (context) => {
  ${code}
  await context.sync();
});
`;
    
    // Use eval to execute the code directly (risky but necessary for web Excel)
    await eval(`(async () => { ${wrappedCode} })()`);
    
    return `✅ **Financial model created successfully!**\n\nThe NPV model has been added to your spreadsheet. Check your current worksheet for the new financial model with formulas and formatting.`;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    console.error('🚨 Direct execution error:', errorMessage);
    
    // Check for common dimension errors and provide helpful guidance
    if (errorMessage.includes('number of rows or columns') || errorMessage.includes('dimensions of the range')) {
      return `❌ **Array dimension mismatch error**\n\nThe generated code tried to set data arrays that don't match Excel range sizes. This is a common issue with dynamically generated code.\n\n**Solutions:**\n• Try asking for the model again (AI will generate different code)\n• Use desktop Excel for better API support\n• Ask for a simpler model structure\n\n**Technical details:** ${errorMessage}`;
    }
    
    if (errorMessage.includes('getCell is not a function')) {
      return `❌ **API compatibility issue**\n\nThe generated code uses Excel APIs that may not be available in web Excel.\n\n**Solutions:**\n• Try using desktop Excel for full API support\n• Ask for the model again (AI may generate different code)\n\n**Technical details:** ${errorMessage}`;
    }
    
    return `❌ **Direct execution failed:**\n\n${errorMessage}\n\nThis might be due to web Excel security restrictions or API limitations. Try using desktop Excel for full Script Lab support.`;
  }
}

// Execute general JavaScript code
async function executeGeneralCode(code: string, engine: SimpleScriptLabEngine): Promise<string> {
  try {
    const snippet = engine.createSnippet(
      'AI Generated Code',
      code,
      '<div id="output">Executing...</div>',
      'body { padding: 10px; }'
    );
    
    const result = await engine.executeSnippet(snippet);
    
    if (result.success) {
      return `✅ Code executed successfully! ${result.result || 'Operation completed.'}`;
    } else {
      return `❌ Execution failed: ${result.error}`;
    }
  } catch (error) {
    return `❌ Error executing code: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

// Insert formula directly into selected cell
async function insertFormulaToSelectedCell(formula: string): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      // Clean the formula (remove surrounding quotes if present)
      const cleanFormula = formula.replace(/^["'`]|["'`]$/g, '');
      
      range.formulas = [[cleanFormula]];
      await context.sync();
      
      return `✅ **Formula inserted into ${range.address}:**\n\n\`${cleanFormula}\``;
    });
    
    return `✅ Formula inserted successfully!`;
  } catch (error) {
    return `❌ Failed to insert formula: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeDirectOperation(code: string): Promise<string> {
  try {
    if (code === "QUICK_CONNECTION_TEST") {
      return await executeQuickConnectionTest();
    }
    
    if (code === "TEST_CONNECTION") {
      return await executeTestConnection();
    }
    
    if (code === "DIRECT_EXCEL_TEST") {
      return await executeDirectExcelTest();
    }
    
    if (code.startsWith("DIRECT_HIGHLIGHT:")) {
      const color = code.split(":")[1];
      return await executeHighlightCells(color);
    }
    
    if (code.startsWith("DIRECT_INSERT_DATA:")) {
      const dataString = code.substring("DIRECT_INSERT_DATA:".length);
      const data = JSON.parse(dataString);
      return await executeInsertData(data);
    }
    
    if (code.startsWith("DIRECT_CREATE_CHART:")) {
      const parts = code.split(":");
      const range = parts[1];
      const chartType = parts[2];
      return await executeCreateChart(range, chartType);
    }
    
    if (code.startsWith("DIRECT_FORMAT:")) {
      const range = code.split(":")[1];
      return await executeFormatCells(range);
    }
    
    if (code.startsWith("DIRECT_INSERT_FORMULA:")) {
      const formula = code.substring("DIRECT_INSERT_FORMULA:".length);
      return await executeInsertFormula(formula);
    }
    
    return "❌ Unknown direct operation";
  } catch (error) {
    return `❌ Direct operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeHighlightCells(color: string): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = color;
      await context.sync();
    });
    return `✅ Selected cells highlighted with ${color} color!`;
  } catch (error) {
    return `❌ Failed to highlight cells: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeInsertData(data: any[][]): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange("A1");
      const resizedRange = range.getResizedRange(data.length - 1, data[0]?.length - 1 || 0);
      resizedRange.values = data;
      await context.sync();
    });
    return `✅ Data inserted successfully starting at A1! (${data.length} rows, ${data[0]?.length || 0} columns)`;
  } catch (error) {
    return `❌ Failed to insert data: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeCreateChart(range: string, chartType: string): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const dataRange = worksheet.getRange(range);
      const chart = worksheet.charts.add(Excel.ChartType[chartType as keyof typeof Excel.ChartType] || Excel.ChartType.columnClustered, dataRange);
      chart.title.text = "Generated Chart";
      await context.sync();
    });
    return `✅ ${chartType} chart created from range ${range}!`;
  } catch (error) {
    return `❌ Failed to create chart: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeFormatCells(range: string): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const cellRange = worksheet.getRange(range);
      cellRange.format.fill.color = "lightblue";
      cellRange.format.font.bold = true;
      cellRange.format.font.size = 12;
      await context.sync();
    });
    return `✅ Range ${range} formatted with blue background and bold text!`;
  } catch (error) {
    return `❌ Failed to format cells: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeQuickConnectionTest(): Promise<string> {
  try {
    console.log('=== Quick Connection Test ===');
    
    // Simple fetch without timeout
    const response = await spreadlyAPI.healthCheck();
    
    return response 
      ? `✅ Quick test: Backend connection successful!` 
      : `❌ Quick test: Backend connection failed. Check Console for details.`;
  } catch (error) {
    return `❌ Quick test error: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeTestConnection(): Promise<string> {
  try {
    console.log('=== Starting Connection Test ===');
    
    // Simple connection test using our testBackendConnection function
    const result = await testBackendConnection();
    return result.message;
  } catch (error) {
    return `❌ Connection test error: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeInsertFormula(formula: string): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      
      // Ensure formula starts with =
      const cleanFormula = formula.startsWith('=') ? formula : `=${formula}`;
      
      // Insert formula into the selected cell(s)
      range.formulas = [[cleanFormula]];
      
      await context.sync();
      return range.address;
    });
    
    return `✅ Formula inserted successfully: ${formula}`;
  } catch (error) {
    return `❌ Failed to insert formula: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

async function executeDirectExcelTest(): Promise<string> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount"]);
      await context.sync();
      
      // Simple test: write data that fits the selection
      if (range.rowCount === 1 && range.columnCount === 1) {
        // Single cell selected
        range.values = [["✅ Spreadly Test!"]];
      } else {
        // Multiple cells selected - expand to fit our data
        const testRange = range.getResizedRange(1, 1); // 2x2 range
        testRange.values = [["Hello", "from"], ["Spreadly", "Direct!"]];
        range.format.fill.color = "lightblue";
        range.format.font.bold = true;
        range.format.font.size = 14;
        return;
      }
      
      range.format.fill.color = "lightblue";
      range.format.font.bold = true;
      range.format.font.size = 14;
      
      await context.sync();
      return range.address;
    });
    
    return "✅ Direct Excel test completed! Check your selected cells - they should now show 'Hello from Spreadly Direct!' with blue background and bold formatting.";
  } catch (error) {
    return `❌ Direct Excel test failed: ${error instanceof Error ? error.message : 'Unknown error'}`;
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}