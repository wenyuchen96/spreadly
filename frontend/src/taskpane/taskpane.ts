/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { ExcelOperations } from '../scriptlab';
import { SimpleScriptLabEngine } from '../scriptlab/SimpleEngine';
import { spreadlyAPI } from '../services/api';
import { getSelectedRangeData, getWorksheetData, getWorksheetInfo } from '../services/excel-data';

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
      if (response.code) {
        addMessage("🔄 Executing code...", "assistant", true);
        const result = await executeGeneratedCode(response.code, scriptLabEngine);
        removeLastMessage();
        addMessage(result, "assistant");
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

async function processUserMessage(message: string, _engine: SimpleScriptLabEngine): Promise<{ message: string; code?: string }> {
  const lowerMessage = message.toLowerCase();
  
  // Check if backend is available
  const backendAvailable = await spreadlyAPI.healthCheck();
  
  if (!backendAvailable) {
    return await processOfflineMessage(message);
  }
  
  // AI-powered processing with backend
  try {
    // Handle special commands first
    if (lowerMessage.includes('analyze') || lowerMessage.includes('analysis') || lowerMessage.includes('insights')) {
      return await handleAnalysisRequest();
    }
    
    if (lowerMessage.includes('upload') || lowerMessage.includes('process') || lowerMessage.includes('send data')) {
      return await handleDataUpload();
    }
    
    if (lowerMessage.includes('formula') || lowerMessage.includes('generate formula')) {
      return await handleFormulaGeneration(message);
    }
    
    // For general queries, try to get context from current spreadsheet
    const hasSessionToken = spreadlyAPI.getSessionToken();
    
    if (hasSessionToken) {
      // Send query to AI with existing session context
      const response = await spreadlyAPI.sendQuery(message);
      
      let responseMessage = response.result.answer || "I processed your request.";
      let code: string | undefined;
      
      // If AI suggests a formula, extract it
      if (response.result.formula) {
        responseMessage += `\n\nSuggested formula: ${response.result.formula}`;
        code = `DIRECT_INSERT_FORMULA:${response.result.formula}`;
      }
      
      return { message: responseMessage, code };
    } else {
      // No session yet, suggest uploading data first
      return {
        message: `I'd love to help with: "${message}"\n\nTo provide the best assistance, I need to analyze your spreadsheet data first. Would you like me to:\n\n• "upload data" - to analyze your current spreadsheet\n• "analyze" - to get AI insights\n• "generate formula [description]" - to create Excel formulas\n\nOr I can help with basic operations without AI analysis.`
      };
    }
  } catch (error) {
    console.error('AI processing error:', error);
    return await processOfflineMessage(message);
  }
}

async function processOfflineMessage(message: string): Promise<{ message: string; code?: string }> {
  const lowerMessage = message.toLowerCase();
  
  // Fallback to simple pattern matching when backend is unavailable
  if (lowerMessage.includes('highlight') || lowerMessage.includes('color')) {
    const color = extractColor(message) || 'yellow';
    return {
      message: `I'll highlight the selected cells with ${color} color.`,
      code: `DIRECT_HIGHLIGHT:${color}`
    };
  }
  
  if (lowerMessage.includes('insert') && lowerMessage.includes('data')) {
    const dataMatch = message.match(/\[\[.*?\]\]/);
    if (dataMatch) {
      try {
        const data = JSON.parse(dataMatch[0]);
        return {
          message: `I'll insert the data into your spreadsheet starting at A1.`,
          code: `DIRECT_INSERT_DATA:${JSON.stringify(data)}`
        };
      } catch (e) {
        return {
          message: "I couldn't parse the data format. Please use format like: [[1,2,3],[4,5,6]]"
        };
      }
    }
  }
  
  if (lowerMessage.includes('chart') || lowerMessage.includes('graph')) {
    const range = extractRange(message) || 'A1:C5';
    const chartType = extractChartType(message) || 'ColumnClustered';
    return {
      message: `I'll create a ${chartType} chart from range ${range}.`,
      code: `DIRECT_CREATE_CHART:${range}:${chartType}`
    };
  }
  
  if (lowerMessage.includes('format') || lowerMessage.includes('style')) {
    const range = extractRange(message) || 'A1:A1';
    return {
      message: `I'll format the range ${range} with blue background and bold text.`,
      code: `DIRECT_FORMAT:${range}`
    };
  }
  
  if (lowerMessage.includes('test') || lowerMessage.includes('demo')) {
    return {
      message: "I'll run a simple test directly with Excel API (no iframe).",
      code: "DIRECT_EXCEL_TEST"
    };
  }
  
  return {
    message: `Backend AI is currently unavailable. I can still help with basic operations:
    
• "highlight cells [color]" - to highlight selected cells
• "insert data [[1,2],[3,4]]" - to insert data  
• "create chart A1:C5" - to create charts
• "format cells A1:B2" - to format cells
• "test" - to run a demo

For AI-powered analysis, please ensure the backend is running at http://localhost:8000`
  };
}

async function handleAnalysisRequest(): Promise<{ message: string; code?: string }> {
  try {
    // First upload current data if no session exists
    if (!spreadlyAPI.getSessionToken()) {
      const uploadResult = await handleDataUpload();
      if (uploadResult.message.includes('Error')) {
        return uploadResult;
      }
    }
    
    // Get AI analysis
    const analysis = await spreadlyAPI.getAnalysis();
    
    let message = "🤖 **AI Analysis Results:**\n\n";
    
    if (analysis.analysis.insights) {
      message += "**Key Insights:**\n";
      analysis.analysis.insights.forEach((insight: string, index: number) => {
        message += `${index + 1}. ${insight}\n`;
      });
      message += "\n";
    }
    
    if (analysis.analysis.suggestions) {
      message += "**Suggestions:**\n";
      analysis.analysis.suggestions.forEach((suggestion: string, index: number) => {
        message += `${index + 1}. ${suggestion}\n`;
      });
      message += "\n";
    }
    
    if (analysis.analysis.formulas) {
      message += "**Recommended Formulas:**\n";
      analysis.analysis.formulas.forEach((formula: any, index: number) => {
        message += `${index + 1}. ${formula.formula} - ${formula.description}\n`;
      });
    }
    
    return { message };
  } catch (error) {
    return { message: `❌ Error getting analysis: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function handleDataUpload(): Promise<{ message: string; code?: string }> {
  try {
    // Get current worksheet data
    const worksheetData = await getWorksheetData();
    const worksheetInfo = await getWorksheetInfo();
    
    if (worksheetData.data.length === 0) {
      return { message: "❌ No data found in the current worksheet. Please add some data first." };
    }
    
    // Upload to backend
    const result = await spreadlyAPI.uploadExcelData(
      worksheetData.data, 
      `${worksheetInfo.activeSheet.name}.xlsx`
    );
    
    return { 
      message: `✅ Data uploaded successfully!\n\n📊 **Data Summary:**\n• Range: ${worksheetData.range}\n• Rows: ${worksheetData.rowCount}\n• Columns: ${worksheetData.columnCount}\n• Session ID: ${result.session_token.substring(0, 8)}...\n\nYou can now ask me to "analyze" the data or ask questions about your spreadsheet!` 
    };
  } catch (error) {
    return { message: `❌ Error uploading data: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function handleFormulaGeneration(message: string): Promise<{ message: string; code?: string }> {
  try {
    // Extract the formula description from the message
    const description = message.replace(/generate formula|formula|create formula/gi, '').trim();
    
    if (!description) {
      return { message: "Please specify what kind of formula you need. For example: 'generate formula to calculate percentage growth'" };
    }
    
    // Get current worksheet context for better formula generation
    let context = "";
    try {
      const worksheetData = await getWorksheetData();
      context = `Worksheet has ${worksheetData.rowCount} rows and ${worksheetData.columnCount} columns. Data types: ${worksheetData.dataTypes.join(', ')}.`;
    } catch (e) {
      // Continue without context if worksheet data can't be read
    }
    
    const response = await spreadlyAPI.generateFormulas(description, context);
    
    let message = `🧮 **Generated Formulas for:** "${description}"\n\n`;
    
    response.formulas.forEach((formula, index) => {
      message += `**${index + 1}. ${formula.difficulty.toUpperCase()}**\n`;
      message += `Formula: \`${formula.formula}\`\n`;
      message += `Description: ${formula.description}\n`;
      if (formula.example) {
        message += `Example: ${formula.example}\n`;
      }
      message += "\n";
    });
    
    // Return the first formula as executable code
    const firstFormula = response.formulas[0];
    return { 
      message, 
      code: firstFormula ? `DIRECT_INSERT_FORMULA:${firstFormula.formula}` : undefined 
    };
  } catch (error) {
    return { message: `❌ Error generating formulas: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function executeGeneratedCode(code: string, engine: SimpleScriptLabEngine): Promise<string> {
  try {
    // Handle all direct Excel operations
    if (code.startsWith("DIRECT_")) {
      return await executeDirectOperation(code);
    }
    
    // Original iframe-based execution (fallback)
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

async function executeDirectOperation(code: string): Promise<string> {
  try {
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
      range.load("address");
      await context.sync();
      
      // Simple test: write data and format cells
      range.values = [["Hello", "from"], ["Spreadly", "Direct!"]];
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

// Helper functions for parsing user input
function extractColor(message: string): string | null {
  const colorMatch = message.match(/(red|blue|green|yellow|orange|purple|pink|cyan|magenta|lime|brown|gray|grey|black|white)/i);
  return colorMatch ? colorMatch[1] : null;
}

function extractRange(message: string): string | null {
  const rangeMatch = message.match(/[A-Z]+\d+:[A-Z]+\d+/i);
  return rangeMatch ? rangeMatch[0].toUpperCase() : null;
}

function extractChartType(message: string): string | null {
  const chartTypes: { [key: string]: string } = {
    'column': 'ColumnClustered',
    'bar': 'BarClustered',
    'line': 'Line',
    'pie': 'Pie',
    'scatter': 'XYScatter',
    'area': 'Area'
  };
  
  for (const [keyword, chartType] of Object.entries(chartTypes)) {
    if (message.toLowerCase().includes(keyword)) {
      return chartType;
    }
  }
  return null;
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
