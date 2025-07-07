/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { ExcelOperations } from '../scriptlab';
import { SimpleScriptLabEngine } from '../scriptlab/SimpleEngine';
import { spreadlyAPI } from '../services/api';
import { dialogAPI } from '../services/dialog-api';
import { mockBackend } from '../services/mock-backend';
import { getSelectedRangeData, getWorksheetData, getWorksheetInfo } from '../services/excel-data';
import { testBackendConnection, testFetchMethods } from '../services/test-connection';

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
  
  // Try Dialog API FIRST for better compatibility with Excel Add-ins
  if (lowerMessage.includes('dialog') || lowerMessage.includes('real') || lowerMessage.includes('backend')) {
    return await tryDialogApiCall(message);
  }
  
  // Try to connect to backend FIRST (enable real AI)
  let backendAvailable = false;
  
  // Special handling for backend test commands
  if (lowerMessage.includes('test backend') || lowerMessage.includes('real test')) {
    try {
      console.log('🔍 Testing backend connection...');
      const controller = new AbortController();
      setTimeout(() => controller.abort(), 5000); // 5 second timeout for manual test
      
      const response = await fetch('http://127.0.0.1:8000/api/excel/test', {
        signal: controller.signal,
        mode: 'cors'
      });
      
      if (response.ok) {
        const data = await response.json();
        console.log('✅ Backend test successful:', data);
        return { message: `🎉 **Backend Connection Successful!**\n\n${data.message}\n\nFull response: ${JSON.stringify(data, null, 2)}` };
      } else {
        console.log('❌ Backend responded with error:', response.status);
        return { message: `❌ Backend test failed with HTTP ${response.status}` };
      }
    } catch (error) {
      console.log('❌ Backend connection error:', error);
      return { message: `❌ Backend connection failed: ${error instanceof Error ? error.message : 'Unknown error'}\n\nMake sure backend is running at http://127.0.0.1:8000` };
    }
  }
  
  try {
    // Quick connection test with shorter timeout for regular commands
    const controller = new AbortController();
    setTimeout(() => controller.abort(), 2000); // 2 second timeout
    
    const response = await fetch('http://127.0.0.1:8000/health', {
      signal: controller.signal,
      mode: 'cors'
    });
    backendAvailable = response.ok;
    console.log('Backend connection:', backendAvailable ? 'SUCCESS' : 'FAILED');
  } catch (error) {
    console.log('Backend connection failed:', error instanceof Error ? error.message : error);
    backendAvailable = false;
  }
  
  if (!backendAvailable) {
    // Try mock backend as intelligent fallback
    if (lowerMessage.includes('mock') || lowerMessage.includes('smart') || lowerMessage.includes('ai demo')) {
      return await processWithMockBackend(message);
    }
    return await processOfflineMessage(message);
  }
  
  // 🎉 REAL AI-powered processing with backend!
  console.log('Using REAL backend AI!');
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

async function processWithMockBackend(message: string): Promise<{ message: string; code?: string }> {
  const lowerMessage = message.toLowerCase();
  
  try {
    if (lowerMessage.includes('formula') || lowerMessage.includes('generate formula')) {
      const description = message.replace(/mock|smart|ai demo|generate formula|formula/gi, '').trim() || 'calculate percentage';
      
      const response = await mockBackend.generateFormulas(description);
      
      let responseMessage = `🤖 **Smart AI Demo - Formula Generated:**\n\n`;
      responseMessage += `**Request:** "${description}"\n\n`;
      
      response.formulas.forEach((formula, index) => {
        responseMessage += `**${index + 1}. ${formula.difficulty.toUpperCase()}**\n`;
        responseMessage += `Formula: \`${formula.formula}\`\n`;
        responseMessage += `Description: ${formula.description}\n`;
        if (formula.example) {
          responseMessage += `Example: ${formula.example}\n`;
        }
        responseMessage += `\n`;
      });
      
      responseMessage += `*This is powered by our smart mock AI - realistic responses without network connectivity!*`;
      
      const firstFormula = response.formulas[0];
      return { 
        message: responseMessage, 
        code: firstFormula ? `DIRECT_INSERT_FORMULA:${firstFormula.formula}` : undefined 
      };
    }
    
    if (lowerMessage.includes('analyze') || lowerMessage.includes('analysis')) {
      const analysis = await mockBackend.getAnalysis();
      
      let message = `🤖 **Smart AI Analysis:**\n\n`;
      
      message += `**Key Insights:**\n`;
      analysis.analysis.insights.forEach((insight, index) => {
        message += `${index + 1}. ${insight}\n`;
      });
      message += `\n`;
      
      message += `**Recommendations:**\n`;
      analysis.analysis.suggestions.forEach((suggestion, index) => {
        message += `${index + 1}. ${suggestion}\n`;
      });
      
      message += `\n*Smart AI Demo - providing realistic analysis without backend connectivity!*`;
      
      return { message };
    }
    
    if (lowerMessage.includes('upload') || lowerMessage.includes('process data')) {
      const worksheetData = await getWorksheetData();
      const worksheetInfo = await getWorksheetInfo();
      
      const uploadResponse = await mockBackend.uploadData(worksheetData.data, worksheetInfo.activeSheet.name);
      
      const message = `🤖 **Smart AI Upload Complete:**\n\n` +
        `📊 **Data Summary:**\n` +
        `• File: ${worksheetInfo.activeSheet.name}\n` +
        `• Rows: ${worksheetData.rowCount}\n` +
        `• Columns: ${worksheetData.columnCount}\n` +
        `• Session: ${uploadResponse.session_token.substring(0, 12)}...\n\n` +
        `**Next Steps:**\n` +
        `• Try "smart analyze" for AI insights\n` +
        `• Ask "smart formula percentage" for formulas\n` +
        `• Query your data with "smart [your question]"\n\n` +
        `*Smart AI Demo - full functionality without network requirements!*`;
      
      return { message };
    }
    
    // General query processing
    const queryResponse = await mockBackend.processQuery(message);
    
    let responseMessage = `🤖 **Smart AI Response:**\n\n${queryResponse.result.answer}`;
    
    let code: string | undefined;
    if (queryResponse.result.formula) {
      responseMessage += `\n\n**Suggested Formula:** \`${queryResponse.result.formula}\``;
      code = `DIRECT_INSERT_FORMULA:${queryResponse.result.formula}`;
    }
    
    responseMessage += `\n\n*Smart AI Demo - providing intelligent responses without backend connectivity!*`;
    
    return { message: responseMessage, code };
    
  } catch (error) {
    return { message: `❌ Smart AI error: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function processOfflineMessage(message: string): Promise<{ message: string; code?: string }> {
  const lowerMessage = message.toLowerCase();
  
  // Try dialog API for specific commands
  if (lowerMessage.includes('dialog api') || lowerMessage.includes('use dialog')) {
    return await tryDialogApiCall(message);
  }
  
  // Try direct API call for specific commands (bypass health check)
  if (lowerMessage.includes('force api') || lowerMessage.includes('try backend')) {
    return await tryDirectApiCall(message);
  }
  
  // Use mock AI responses as fallback when backend unavailable
  if (lowerMessage.includes('generate formula') || lowerMessage.includes('formula')) {
    return await generateMockFormula(message);
  }
  
  if (lowerMessage.includes('analyze') || lowerMessage.includes('analysis')) {
    return await generateMockAnalysis();
  }
  
  if (lowerMessage.includes('upload') || lowerMessage.includes('process data')) {
    return await generateMockUpload();
  }
  
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
  
  if (lowerMessage.includes('debug simple') || lowerMessage.includes('quick test')) {
    return {
      message: "I'll do a quick connection test without iframe.",
      code: "QUICK_CONNECTION_TEST"
    };
  }
  
  if (lowerMessage.includes('test connection') || lowerMessage.includes('debug connection')) {
    return {
      message: "I'll test the backend connection and show debug info in the console.",
      code: "TEST_CONNECTION"
    };
  }
  
  if (lowerMessage.includes('test backend') || lowerMessage.includes('real test')) {
    return {
      message: "❌ Backend connection failed in offline mode. Make sure the FastAPI server is running at http://127.0.0.1:8000",
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
    
🔧 **Basic Operations:**
• "highlight cells [color]" - to highlight selected cells
• "insert data [[1,2],[3,4]]" - to insert data  
• "create chart A1:C5" - to create charts
• "format cells A1:B2" - to format cells
• "test" - to run a demo

🚀 **Try Backend Connection:**
• "dialog api" - use dialog window method (recommended)
• "dialog api formula percentage" - test formula generation via dialog
• "force api" - attempt direct API call (likely blocked)

Note: Excel Add-ins block direct network requests. Dialog method bypasses this.`
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

async function tryDialogApiCall(message: string): Promise<{ message: string; code?: string }> {
  try {
    console.log('Attempting dialog API call...');
    
    // Try formula generation through dialog
    if (message.includes('formula')) {
      const description = message.replace(/dialog api|use dialog|formula/gi, '').trim() || 'calculate percentage';
      
      const response = await dialogAPI.generateFormulas(description);
      
      let responseMessage = `🎉 **Dialog API Success!** Generated formulas for: "${description}"\n\n`;
      
      response.formulas.forEach((formula: any, index: number) => {
        responseMessage += `**${index + 1}. ${formula.difficulty.toUpperCase()}**\n`;
        responseMessage += `Formula: \`${formula.formula}\`\n`;
        responseMessage += `Description: ${formula.description}\n\n`;
      });
      
      const firstFormula = response.formulas[0];
      return { 
        message: responseMessage, 
        code: firstFormula ? `DIRECT_INSERT_FORMULA:${firstFormula.formula}` : undefined 
      };
    }
    
    // Try health check through dialog
    const healthCheck = await dialogAPI.healthCheck();
    return { 
      message: healthCheck 
        ? `🎉 **Dialog API Success!** Backend is accessible through dialog method!`
        : `❌ Dialog API health check failed. Backend may not be running.`
    };
  } catch (error) {
    return { 
      message: `❌ Dialog API error: ${error instanceof Error ? error.message : 'Unknown error'}\n\nThis method opens a dialog window to bypass Excel's network restrictions.` 
    };
  }
}

async function tryDirectApiCall(message: string): Promise<{ message: string; code?: string }> {
  try {
    console.log('Attempting direct API call...');
    
    // Try formula generation as a test
    if (message.includes('formula')) {
      const description = message.replace(/force api|try backend|formula/gi, '').trim() || 'calculate percentage';
      
      const response = await spreadlyAPI.generateFormulas(description);
      
      let responseMessage = `🎉 **Direct API Success!** Generated formulas for: "${description}"\n\n`;
      
      response.formulas.forEach((formula, index) => {
        responseMessage += `**${index + 1}. ${formula.difficulty.toUpperCase()}**\n`;
        responseMessage += `Formula: \`${formula.formula}\`\n`;
        responseMessage += `Description: ${formula.description}\n\n`;
      });
      
      const firstFormula = response.formulas[0];
      return { 
        message: responseMessage, 
        code: firstFormula ? `DIRECT_INSERT_FORMULA:${firstFormula.formula}` : undefined 
      };
    }
    
    // Try a simple health check
    const response = await fetch('http://127.0.0.1:8000/health');
    if (response.ok) {
      const data = await response.json();
      return { 
        message: `🎉 **Direct API Success!** Backend is responding: ${JSON.stringify(data)}` 
      };
    } else {
      return { 
        message: `❌ Direct API call failed with status: ${response.status}` 
      };
    }
  } catch (error) {
    return { 
      message: `❌ Direct API error: ${error instanceof Error ? error.message : 'Unknown error'}\n\nThe Excel Add-in environment likely blocks external network requests for security.` 
    };
  }
}

// Mock AI functions that simulate backend responses
async function generateMockFormula(message: string): Promise<{ message: string; code?: string }> {
  try {
    // Extract what kind of formula they want
    const description = message.toLowerCase();
    
    let formula = "SUM(A1:A10)";
    let explanation = "Calculates the sum of values in range A1:A10";
    
    if (description.includes('percentage') || description.includes('percent')) {
      formula = "=(B2-A2)/A2*100";
      explanation = "Calculates percentage change between two values";
    } else if (description.includes('average') || description.includes('mean')) {
      formula = "=AVERAGE(A1:A10)";
      explanation = "Calculates the average of values in range A1:A10";
    } else if (description.includes('count')) {
      formula = "=COUNTA(A1:A10)";
      explanation = "Counts non-empty cells in range A1:A10";
    } else if (description.includes('max') || description.includes('maximum')) {
      formula = "=MAX(A1:A10)";
      explanation = "Finds the maximum value in range A1:A10";
    } else if (description.includes('min') || description.includes('minimum')) {
      formula = "=MIN(A1:A10)";
      explanation = "Finds the minimum value in range A1:A10";
    } else if (description.includes('growth') || description.includes('change')) {
      formula = "=((B2-A2)/A2)*100";
      explanation = "Calculates growth rate as a percentage";
    }
    
    const responseMessage = `🤖 **AI Formula Generated:**

**Formula:** \`${formula}\`
**Description:** ${explanation}

**Usage Tips:**
• Click to insert this formula into the selected cell
• Adjust cell references (A1, B2, etc.) as needed
• This formula will calculate automatically when cell values change

*Note: This is a demo AI response. In production, this would be powered by Claude AI through the backend.*`;

    return {
      message: responseMessage,
      code: `DIRECT_INSERT_FORMULA:${formula}`
    };
  } catch (error) {
    return { message: `❌ Error generating formula: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function generateMockAnalysis(): Promise<{ message: string; code?: string }> {
  try {
    // Get current worksheet data for analysis
    const worksheetData = await getWorksheetData();
    
    let analysisMessage = `🤖 **AI Analysis Results:**

📊 **Data Summary:**
• Spreadsheet Range: ${worksheetData.range}
• Total Rows: ${worksheetData.rowCount}
• Total Columns: ${worksheetData.columnCount}
• Data Types: ${worksheetData.dataTypes.join(', ')}

🔍 **Key Insights:**
• Your data appears to be ${worksheetData.hasHeaders ? 'well-structured with headers' : 'numeric data without headers'}
• Consider adding charts to visualize trends
• Look for outliers in numeric columns
• Validate data consistency across rows

💡 **Recommendations:**
• Use conditional formatting to highlight important values
• Create pivot tables for data summarization
• Apply data validation to ensure accuracy
• Consider using formulas for calculated fields

*Note: This is a demo AI analysis. In production, Claude AI would provide much more detailed insights based on actual data content.*`;

    return { message: analysisMessage };
  } catch (error) {
    return { message: `❌ Error analyzing data: ${error instanceof Error ? error.message : 'Unknown error'}` };
  }
}

async function generateMockUpload(): Promise<{ message: string; code?: string }> {
  try {
    const worksheetData = await getWorksheetData();
    const worksheetInfo = await getWorksheetInfo();
    
    const responseMessage = `✅ **Mock Data Upload Successful!**

📄 **File Info:**
• Sheet Name: ${worksheetInfo.activeSheet.name}
• Data Range: ${worksheetData.range}
• Rows: ${worksheetData.rowCount}
• Columns: ${worksheetData.columnCount}

🤖 **AI Processing Complete:**
• Data structure analyzed
• Statistical summary generated
• Ready for AI queries and insights

**Try these commands:**
• "analyze" - Get detailed AI insights
• "generate formula percentage" - Create custom formulas
• Ask questions about your data

*Note: This is a demo upload. In production, your data would be securely processed by Claude AI through the backend API.*`;

    return { message: responseMessage };
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
    
    // Test basic connection
    const workingUrl = await testBackendConnection();
    
    if (workingUrl) {
      console.log(`Found working URL: ${workingUrl}`);
      
      // Test different fetch methods with the working URL
      const workingConfig = await testFetchMethods(workingUrl);
      
      return `✅ Connection test completed! Check the Console (F12) for detailed results.\n\nWorking URL: ${workingUrl}\nConfiguration: ${workingConfig ? JSON.stringify(workingConfig) : 'Default'}`;
    } else {
      return `❌ Connection test failed! Check the Console (F12) for detailed error messages.\n\nPossible issues:\n• Backend not running on http://127.0.0.1:8000\n• CORS restrictions in Excel Add-in environment\n• Network connectivity issues\n\nTry opening http://127.0.0.1:8000/health in your browser to verify the backend is working.`;
    }
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
