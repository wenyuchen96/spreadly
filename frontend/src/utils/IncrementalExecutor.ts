/**
 * Incremental Executor for Excel Financial Models
 * 
 * Orchestrates chunk-by-chunk model building with real-time error recovery,
 * progress tracking, and adaptive optimization.
 */

import { BASE_URL } from '../config/api-config';
import { getComprehensiveWorkbookData } from '../services/excel-data';

export interface ChunkInfo {
  id: string;
  type: string;
  complexity: string;
  code: string;
  description: string;
  estimated_operations: number;
  stage: number;
}

export interface ExecutionProgress {
  session_id: string;
  model_type: string;
  progress_percentage: number;
  success_rate: number;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  current_chunk_id?: string;
  execution_history: string[];
  error_patterns: string[];
  elapsed_time: number;
}

export interface ExecutionResult {
  success: boolean;
  error_message?: string;
  execution_time: number;
  chunk_id: string;
}

export type ProgressCallback = (progress: ExecutionProgress) => void;
export type ChunkCallback = (chunk: ChunkInfo) => void;
export type ErrorCallback = (error: string, chunk: ChunkInfo) => void;

export class IncrementalExecutor {
  private sessionToken: string;
  private isExecuting: boolean = false;
  private currentChunk: ChunkInfo | null = null;
  private progressCallback?: ProgressCallback;
  private chunkCallback?: ChunkCallback;
  private errorCallback?: ErrorCallback;
  
  constructor(sessionToken: string) {
    this.sessionToken = sessionToken;
  }
  
  /**
   * Start incremental model building
   */
  async startIncrementalBuild(
    modelType: string,
    query: string,
    progressCallback?: ProgressCallback,
    chunkCallback?: ChunkCallback,
    errorCallback?: ErrorCallback
  ): Promise<boolean> {
    
    this.progressCallback = progressCallback;
    this.chunkCallback = chunkCallback;
    this.errorCallback = errorCallback;
    
    try {
      console.log(`🔧 Starting incremental ${modelType} model building...`);
      
      // Get current workbook context
      const workbookContext = await getComprehensiveWorkbookData();
      
      // Initialize incremental build session
      const initResponse = await fetch(`${BASE_URL}/api/incremental/start`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_token: this.sessionToken,
          model_type: modelType,
          query: query,
          workbook_context: workbookContext
        })
      });
      
      if (!initResponse.ok) {
        throw new Error(`Failed to start incremental build: ${initResponse.statusText}`);
      }
      
      const initResult = await initResponse.json();
      console.log('✅ Incremental build initialized:', initResult);
      
      // Start the execution loop
      this.isExecuting = true;
      return await this.executeIncrementalLoop();
      
    } catch (error) {
      console.error('❌ Failed to start incremental build:', error);
      this.errorCallback?.(String(error), this.currentChunk!);
      return false;
    }
  }
  
  /**
   * Main execution loop - continues until model is complete
   */
  private async executeIncrementalLoop(): Promise<boolean> {
    let consecutiveErrors = 0;
    let totalRetries = 0;
    const maxConsecutiveErrors = 3;
    const maxTotalRetries = 10; // Circuit breaker for infinite loops
    const retryBackoffMs = 1000;
    
    while (this.isExecuting) {
      try {
        // Circuit breaker: Stop if too many total retries
        if (totalRetries >= maxTotalRetries) {
          console.error('🛑 Circuit breaker: Too many total retries, stopping to prevent infinite loop');
          this.errorCallback?.('Circuit breaker activated: Too many retries', this.currentChunk!);
          return false;
        }
        
        // Get current Excel context
        const currentContext = await getComprehensiveWorkbookData();
        
        // Report previous chunk result if we have one
        let lastResult = null;
        if (this.currentChunk) {
          // Assume previous chunk succeeded if we got here without errors
          lastResult = {
            chunk_id: this.currentChunk.id,
            success: true,
            execution_time: 0.5
          };
        }
        
        // Request next chunk
        const nextChunk = await this.getNextChunk(currentContext, lastResult);
        
        if (nextChunk.completed) {
          console.log('🎉 Model building completed!');
          this.progressCallback?.(nextChunk.progress);
          return true;
        }
        
        if (!nextChunk.chunk) {
          console.log('⚠️ No more chunks available');
          return false;
        }
        
        this.currentChunk = nextChunk.chunk;
        this.chunkCallback?.(this.currentChunk);
        this.progressCallback?.(nextChunk.progress);
        
        // Execute the chunk with timeout
        const execResult = await this.executeChunkWithTimeout(this.currentChunk, 30000); // 30s timeout
        
        if (execResult.success) {
          console.log(`✅ Chunk ${this.currentChunk.id} executed successfully`);
          consecutiveErrors = 0;
          totalRetries = 0; // Reset retry counter on success
        } else {
          console.warn(`❌ Chunk ${this.currentChunk.id} failed:`, execResult.error_message);
          consecutiveErrors++;
          totalRetries++;
          
          this.errorCallback?.(execResult.error_message || 'Unknown error', this.currentChunk);
          
          // Check if this is a persistent error pattern
          if (this.isPersistentError(execResult.error_message)) {
            console.error('🛑 Detected persistent error pattern, moving to next chunk');
            consecutiveErrors = 0; // Reset to continue with next chunk
            await this.delay(retryBackoffMs);
            continue;
          }
          
          // Handle the error with rate limiting
          await this.delay(retryBackoffMs * consecutiveErrors); // Exponential backoff
          
          const errorHandlingResult = await this.handleChunkError(this.currentChunk, execResult);
          
          // If chunk was auto-fixed, retry it immediately (but with limits)
          if (errorHandlingResult?.auto_fixed && totalRetries < maxTotalRetries) {
            console.log(`🔄 Chunk ${this.currentChunk.id} was auto-fixed, retrying...`);
            // Get the updated chunk and retry execution
            const updatedChunk = await this.getUpdatedChunk(this.currentChunk.id);
            if (updatedChunk) {
              this.currentChunk = updatedChunk;
              // Don't increment consecutiveErrors since we're retrying with fixed code
              consecutiveErrors = Math.max(0, consecutiveErrors - 1);
              continue; // Retry the fixed chunk
            }
          }
          
          // Stop if too many consecutive errors
          if (consecutiveErrors >= maxConsecutiveErrors) {
            console.error('🛑 Too many consecutive errors, stopping execution');
            return false;
          }
        }
        
        // Small delay between chunks for stability
        await this.delay(500);
        
      } catch (error) {
        console.error('❌ Error in execution loop:', error);
        consecutiveErrors++;
        totalRetries++;
        
        if (consecutiveErrors >= maxConsecutiveErrors || totalRetries >= maxTotalRetries) {
          console.error('🛑 Too many consecutive errors in execution loop');
          return false;
        }
        
        await this.delay(retryBackoffMs * consecutiveErrors); // Exponential backoff
      }
    }
    
    return false;
  }
  
  /**
   * Get the next chunk to execute
   */
  private async getNextChunk(currentContext: any, lastResult: any = null): Promise<any> {
    const response = await fetch(`${BASE_URL}/api/incremental/next-chunk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        session_token: this.sessionToken,
        current_context: currentContext,
        last_execution_result: lastResult
      })
    });
    
    if (!response.ok) {
      throw new Error(`Failed to get next chunk: ${response.statusText}`);
    }
    
    return await response.json();
  }
  
  /**
   * Execute a single code chunk in Excel with timeout
   */
  private async executeChunkWithTimeout(chunk: ChunkInfo, timeoutMs: number = 30000): Promise<ExecutionResult> {
    const startTime = performance.now();
    
    try {
      console.log(`🔧 Executing chunk ${chunk.id} (${chunk.type}, ${chunk.complexity})`);
      console.log('🔧 Chunk code to execute:', chunk.code.substring(0, 200) + '...');
      
      // Validate and clean the code before execution
      const cleanedCode = this.validateAndCleanCode(chunk.code);
      
      if (!cleanedCode || cleanedCode.trim().length === 0) {
        throw new Error('Generated code is empty or invalid');
      }
      
      // Execute with timeout to prevent hanging
      const result = await Promise.race([
        this.executeChunk(cleanedCode),
        new Promise<never>((_, reject) => 
          setTimeout(() => reject(new Error('Execution timeout')), timeoutMs)
        )
      ]);
      
      const executionTime = performance.now() - startTime;
      
      return {
        success: true,
        execution_time: executionTime,
        chunk_id: chunk.id
      };
      
    } catch (error) {
      const executionTime = performance.now() - startTime;
      console.error(`❌ Chunk execution failed:`, error);
      
      // Enhanced error details for debugging
      const errorDetails = {
        message: String(error),
        name: error instanceof Error ? error.name : 'Unknown',
        stack: error instanceof Error ? error.stack : undefined,
        code_preview: chunk.code.substring(0, 500),
        syntax_error: error instanceof SyntaxError,
        timeout_error: error.message === 'Execution timeout',
        chunk_info: {
          id: chunk.id,
          type: chunk.type,
          complexity: chunk.complexity,
          code_length: chunk.code.length
        }
      };
      
      return {
        success: false,
        error_message: JSON.stringify(errorDetails),
        execution_time: executionTime,
        chunk_id: chunk.id
      };
    }
  }
  
  /**
   * Execute a single code chunk in Excel (internal method)
   */
  private async executeChunk(cleanedCode: string): Promise<void> {
    // Check if Excel API is available
    if (typeof Excel === 'undefined') {
      throw new Error('Excel API not available. Make sure this code runs in Excel context.');
    }
    
    // Execute the actual chunk code in proper Excel context
    if (cleanedCode.includes("Excel.run")) {
      // Code already has Excel.run wrapper
      await eval(`(async () => { ${cleanedCode} })()`);
    } else {
      // Wrap in Excel.run
      await Excel.run(async (context) => {
        try {
          // Execute code in Excel context
          await eval(`(async () => { ${cleanedCode} })()`);
          await context.sync();
        } catch (error) {
          console.error('Error in Excel context:', error);
          throw error;
        }
      });
    }
  }
  
  /**
   * Validate and clean JavaScript code before execution
   */
  private validateAndCleanCode(code: string): string {
    let cleaned = code.trim();
    
    // Remove any markdown code fences
    cleaned = cleaned.replace(/^```(javascript|js)?\n?/gm, '');
    cleaned = cleaned.replace(/\n?```$/gm, '');
    
    // Fix common syntax issues
    cleaned = cleaned.replace(/;{2,}/g, ';'); // Remove multiple semicolons
    cleaned = cleaned.replace(/\s+/g, ' '); // Normalize whitespace
    cleaned = cleaned.trim();
    
    // Fix truncated code issues
    cleaned = this.fixTruncatedCode(cleaned);
    
    return cleaned;
  }
  
  /**
   * Fix truncated/incomplete JavaScript code
   */
  private fixTruncatedCode(code: string): string {
    let fixed = code;
    
    // Check if code ends mid-statement
    const lines = fixed.split('\n');
    const lastLine = lines[lines.length - 1].trim();
    
    // Fix incomplete statements
    if (lastLine.endsWith('sheet.getRange("A3') || 
        lastLine.endsWith('sheet.getRange(') ||
        lastLine.includes('getRange(') && !lastLine.includes(')')) {
      // Remove incomplete line
      lines.pop();
      fixed = lines.join('\n');
    }
    
    // Ensure proper Excel.run closure
    if (fixed.includes('Excel.run(async (context) => {')) {
      // Count braces to ensure proper closure
      const openBraces = (fixed.match(/\{/g) || []).length;
      const closeBraces = (fixed.match(/\}/g) || []).length;
      
      if (openBraces > closeBraces) {
        // Add missing context.sync() and closing braces
        if (!fixed.includes('await context.sync()')) {
          fixed += '\n    await context.sync();';
        }
        
        // Add missing closing braces
        const missingBraces = openBraces - closeBraces;
        for (let i = 0; i < missingBraces; i++) {
          fixed += '\n}';
        }
      }
    }
    
    return fixed.trim();
  }
  
  /**
   * Handle chunk execution error
   */
  private async handleChunkError(chunk: ChunkInfo, result: ExecutionResult): Promise<any> {
    try {
      const response = await fetch(`${BASE_URL}/api/incremental/handle-error`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_token: this.sessionToken,
          chunk_id: chunk.id,
          error_message: result.error_message,
          execution_time: result.execution_time,
          current_context: await getComprehensiveWorkbookData()
        })
      });
      
      if (response.ok) {
        const errorResult = await response.json();
        console.log('🔧 Error handling result:', errorResult);
        return errorResult;
      }
      
      return null;
      
    } catch (error) {
      console.warn('⚠️ Failed to handle chunk error:', error);
      return null;
    }
  }
  
  /**
   * Get updated chunk after auto-fixing
   */
  private async getUpdatedChunk(chunkId: string): Promise<ChunkInfo | null> {
    try {
      const response = await fetch(`${BASE_URL}/api/incremental/next-chunk`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_token: this.sessionToken,
          current_context: await getComprehensiveWorkbookData(),
          retry_chunk_id: chunkId  // Request the specific fixed chunk
        })
      });
      
      if (response.ok) {
        const result = await response.json();
        // If we get the same chunk ID back, it was updated
        if (result.chunk && result.chunk.id === chunkId) {
          return result.chunk;
        }
      }
      
      return null;
      
    } catch (error) {
      console.warn('⚠️ Failed to get updated chunk:', error);
      return null;
    }
  }
  
  /**
   * Get current build progress
   */
  async getProgress(): Promise<ExecutionProgress | null> {
    try {
      const response = await fetch(`${BASE_URL}/api/incremental/status/${this.sessionToken}`);
      
      if (!response.ok) {
        return null;
      }
      
      const result = await response.json();
      return result.progress;
      
    } catch (error) {
      console.warn('⚠️ Failed to get progress:', error);
      return null;
    }
  }
  
  /**
   * Stop the execution
   */
  async stop(): Promise<void> {
    this.isExecuting = false;
    
    try {
      await fetch(`${BASE_URL}/api/incremental/cancel/${this.sessionToken}`, {
        method: 'POST'
      });
      
      console.log('🛑 Incremental execution stopped');
      
    } catch (error) {
      console.warn('⚠️ Failed to cancel build session:', error);
    }
  }
  
  /**
   * Check if error indicates a persistent issue that should skip retries
   */
  private isPersistentError(errorMessage?: string): boolean {
    if (!errorMessage) return false;
    
    const persistentPatterns = [
      'Excel API not available',
      'context is undefined',
      'Excel is not defined',
      'Office.js not loaded',
      'Permission denied',
      'Network error',
      'Rate limit exceeded',
      'Service unavailable'
    ];
    
    const errorLower = errorMessage.toLowerCase();
    return persistentPatterns.some(pattern => errorLower.includes(pattern.toLowerCase()));
  }
  
  /**
   * Utility function for delays
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * Detect if a query should use incremental building
 */
export function shouldUseIncrementalBuild(query: string): boolean {
  const incrementalKeywords = [
    'financial model', 'dcf', 'npv', 'lbo', 'leverage buyout',
    'valuation model', 'cash flow', 'model', 'build'
  ];
  
  const queryLower = query.toLowerCase();
  return incrementalKeywords.some(keyword => queryLower.includes(keyword));
}

/**
 * Extract model type from query
 */
export function extractModelType(query: string): string {
  const queryLower = query.toLowerCase();
  
  if (queryLower.includes('three statement') || queryLower.includes('3 statement') || 
      queryLower.includes('three-statement') || queryLower.includes('3-statement') ||
      queryLower.includes('integrated model') || queryLower.includes('income statement') ||
      queryLower.includes('balance sheet') || queryLower.includes('cash flow statement')) {
    return 'three-statement';
  } else if (queryLower.includes('dcf') || queryLower.includes('discounted cash flow')) {
    return 'dcf';
  } else if (queryLower.includes('npv') || queryLower.includes('net present value')) {
    return 'npv';
  } else if (queryLower.includes('lbo') || queryLower.includes('leverage buyout')) {
    return 'lbo';
  } else if (queryLower.includes('valuation')) {
    return 'valuation';
  } else {
    return 'financial';
  }
}