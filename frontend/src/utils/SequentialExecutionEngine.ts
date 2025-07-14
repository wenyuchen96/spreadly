/**
 * Sequential Execution Engine for Excel Financial Models
 * 
 * Replaces monolithic code execution with resilient, step-by-step operation execution.
 * Provides fault isolation, error recovery, and detailed progress tracking.
 */

export interface ExecutionOperation {
  id: string;
  type: 'sheet_setup' | 'header' | 'data' | 'formula' | 'formatting' | 'validation';
  code: string;
  dependencies: string[];
  optional: boolean;
  description: string;
  stage: number;
}

export interface ExecutionResult {
  operationId: string;
  success: boolean;
  error?: string;
  executionTime: number;
  retryCount: number;
}

export interface ExecutionProgress {
  currentStage: number;
  totalStages: number;
  currentOperation: string;
  operationsCompleted: number;
  totalOperations: number;
  errors: ExecutionResult[];
  timeElapsed: number;
}

export type ExecutionStrategy = 'default' | 'fast' | 'safe' | 'adaptive';

export interface ExecutionConfig {
  strategy: ExecutionStrategy;
  maxRetries: number;
  continueOnError: boolean;
  validateEachStage: boolean;
  progressCallback?: (progress: ExecutionProgress) => void;
  operationTimeout: number;
}

export class SequentialExecutionEngine {
  private config: ExecutionConfig;
  private operations: ExecutionOperation[] = [];
  private results: ExecutionResult[] = [];
  private startTime: number = 0;
  private currentStage: number = 0;

  constructor(config: Partial<ExecutionConfig> = {}) {
    this.config = {
      strategy: 'default',
      maxRetries: 2,
      continueOnError: true,
      validateEachStage: false, // Disable for faster execution
      operationTimeout: 5000, // 5 seconds per operation
      ...config
    };
    
    console.log('🔄 SequentialExecutionEngine initialized with strategy:', this.config.strategy);
    console.log('🔄 Config:', this.config);
  }

  /**
   * Parse monolithic Excel.js code into sequential operations
   */
  parseCodeIntoOperations(code: string, description: string = 'Financial Model'): ExecutionOperation[] {
    console.log('🔍 Parsing code into sequential operations...');
    console.log('📝 Code length:', code.length, 'characters');
    console.log('📝 Code preview:', code.substring(0, 200) + '...');
    
    const operations: ExecutionOperation[] = [];
    let operationCounter = 0;

    try {
      // Split by Excel.run() blocks first
      console.log('🔍 Extracting Excel.run() blocks...');
      const excelRunBlocks = this.extractExcelRunBlocks(code);
      console.log(`🔍 Found ${excelRunBlocks.length} Excel.run() blocks`);
      
      if (excelRunBlocks.length === 0) {
        console.log('⚠️ No Excel.run blocks found, treating as single operation');
        // No Excel.run blocks found, treat as single operation
        operations.push({
          id: `op_${++operationCounter}`,
          type: 'data',
          code: code,
          dependencies: [],
          optional: false,
          description: description,
          stage: 1
        });
        return operations;
      }

      console.log('🔍 Parsing individual Excel.run() blocks...');
      excelRunBlocks.forEach((block, blockIndex) => {
        console.log(`🔍 Parsing block ${blockIndex + 1}/${excelRunBlocks.length}`);
        try {
          const blockOperations = this.parseExcelRunBlock(block, blockIndex, operationCounter);
          operations.push(...blockOperations);
          operationCounter += blockOperations.length;
          console.log(`✅ Block ${blockIndex + 1} parsed: ${blockOperations.length} operations`);
        } catch (error) {
          console.error(`❌ Error parsing block ${blockIndex + 1}:`, error);
          // Add fallback operation for failed block
          operations.push({
            id: `op_${++operationCounter}_fallback`,
            type: 'data',
            code: block,
            dependencies: [],
            optional: true,
            description: `Fallback for block ${blockIndex + 1}`,
            stage: blockIndex + 1
          });
        }
      });

      console.log(`✅ Parsed ${operations.length} operations across ${this.getMaxStage(operations)} stages`);
      return operations;
    } catch (error) {
      console.error('❌ Critical error in parseCodeIntoOperations:', error);
      // Return single fallback operation
      return [{
        id: 'op_1_critical_fallback',
        type: 'data',
        code: code,
        dependencies: [],
        optional: false,
        description: 'Critical fallback - single operation',
        stage: 1
      }];
    }
  }

  /**
   * Extract Excel.run() blocks from code
   */
  private extractExcelRunBlocks(code: string): string[] {
    const blocks: string[] = [];
    const excelRunPattern = /Excel\.run\s*\(\s*async\s*\(\s*context\s*\)\s*=>\s*\{/g;
    
    let match;
    const matches: { index: number, fullMatch: string }[] = [];
    
    while ((match = excelRunPattern.exec(code)) !== null) {
      matches.push({ index: match.index, fullMatch: match[0] });
    }

    if (matches.length === 0) {
      return [code]; // No Excel.run blocks, return entire code
    }

    matches.forEach((match, index) => {
      const startIndex = match.index;
      const endIndex = this.findMatchingBrace(code, startIndex + match.fullMatch.length - 1);
      
      if (endIndex !== -1) {
        const blockCode = code.substring(startIndex, endIndex + 2); // Include closing }});
        blocks.push(blockCode);
      }
    });

    return blocks;
  }

  /**
   * Find matching closing brace for Excel.run block
   */
  private findMatchingBrace(code: string, startIndex: number): number {
    let braceCount = 1;
    let inString = false;
    let inComment = false;
    let i = startIndex + 1;

    while (i < code.length && braceCount > 0) {
      const char = code[i];
      const prevChar = i > 0 ? code[i - 1] : '';
      const nextChar = i < code.length - 1 ? code[i + 1] : '';

      // Handle string literals
      if (char === '"' && prevChar !== '\\') {
        inString = !inString;
      }
      
      // Handle comments
      if (!inString && char === '/' && nextChar === '/') {
        inComment = true;
      }
      if (inComment && char === '\n') {
        inComment = false;
      }

      // Count braces only outside strings and comments
      if (!inString && !inComment) {
        if (char === '{') {
          braceCount++;
        } else if (char === '}') {
          braceCount--;
        }
      }

      i++;
    }

    return braceCount === 0 ? i - 1 : -1;
  }

  /**
   * Parse single Excel.run block into operations
   */
  private parseExcelRunBlock(block: string, blockIndex: number, startCounter: number): ExecutionOperation[] {
    console.log(`🔍 Parsing Excel.run block ${blockIndex + 1}...`);
    const operations: ExecutionOperation[] = [];
    let operationCounter = startCounter;

    try {
      // Extract the content inside Excel.run()
      const contentMatch = block.match(/Excel\.run\s*\(\s*async\s*\(\s*context\s*\)\s*=>\s*\{(.*)\}\s*\)\s*;?/s);
      if (!contentMatch) {
        console.warn('⚠️ Could not parse Excel.run block, treating as single operation');
        // Fallback: treat entire block as single operation
        operations.push({
          id: `op_${++operationCounter}_block_${blockIndex}`,
          type: 'data',
          code: block,
          dependencies: [],
          optional: false,
          description: `Excel.run block ${blockIndex + 1}`,
          stage: blockIndex + 1
        });
        return operations;
      }

      const content = contentMatch[1];
      const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
      
      console.log(`🔍 Found ${lines.length} lines to parse`);

      // Simple approach: if the block has stage comments, split by them
      if (content.includes('// STAGE')) {
        console.log('🔍 Using stage-based parsing');
        return this.parseByStageComments(block, blockIndex, operationCounter);
      }

      // Otherwise, use line-by-line approach but with simpler logic
      console.log('🔍 Using line-by-line parsing');
      let currentOperation: Partial<ExecutionOperation> = {
        dependencies: [],
        optional: false,
        stage: 1
      };
      let operationLines: string[] = [];

      for (const line of lines) {
        // Detect operation boundaries and types
        if (this.isOperationBoundary(line)) {
          // Finish current operation if it has content
          if (operationLines.length > 0) {
            operations.push(this.finalizeOperation(currentOperation, operationLines, ++operationCounter, blockIndex));
            operationLines = [];
          }
          
          // Start new operation
          currentOperation = {
            dependencies: [],
            optional: false,
            ...this.analyzeOperationType(line)
          };
        }
        
        operationLines.push(line);
      }

      // Finalize last operation
      if (operationLines.length > 0) {
        operations.push(this.finalizeOperation(currentOperation, operationLines, ++operationCounter, blockIndex));
      }

      console.log(`✅ Parsed block into ${operations.length} operations`);
      return operations;
      
    } catch (error) {
      console.error('❌ Error in parseExcelRunBlock:', error);
      // Fallback: single operation for entire block
      operations.push({
        id: `op_${++operationCounter}_fallback_${blockIndex}`,
        type: 'data',
        code: block,
        dependencies: [],
        optional: false,
        description: `Fallback operation for block ${blockIndex + 1}`,
        stage: blockIndex + 1
      });
      return operations;
    }
  }

  /**
   * Parse by stage comments (// STAGE 1:, // STAGE 2:, etc.)
   */
  private parseByStageComments(block: string, blockIndex: number, startCounter: number): ExecutionOperation[] {
    console.log('🔍 Parsing by stage comments...');
    const operations: ExecutionOperation[] = [];
    let operationCounter = startCounter;
    
    // Split by stage comments
    const stagePattern = /\/\/\s*STAGE\s*(\d+):\s*([^\n]*)/gi;
    const parts = block.split(stagePattern);
    
    for (let i = 1; i < parts.length; i += 3) {
      const stageNumber = parseInt(parts[i]);
      const stageDescription = parts[i + 1] || 'Stage operation';
      const stageCode = parts[i + 2] || '';
      
      if (stageCode.trim()) {
        operations.push({
          id: `op_${++operationCounter}_stage_${stageNumber}`,
          type: this.getStageType(stageNumber),
          code: this.wrapInExcelRun([stageCode.trim()]),
          dependencies: [],
          optional: false,
          description: stageDescription.trim(),
          stage: stageNumber
        });
      }
    }
    
    // If no stages found, fallback to single operation
    if (operations.length === 0) {
      operations.push({
        id: `op_${++operationCounter}_single`,
        type: 'data',
        code: block,
        dependencies: [],
        optional: false,
        description: 'Single operation fallback',
        stage: 1
      });
    }
    
    console.log(`✅ Stage-based parsing created ${operations.length} operations`);
    return operations;
  }

  /**
   * Get operation type based on stage number
   */
  private getStageType(stageNumber: number): ExecutionOperation['type'] {
    if (stageNumber === 1) return 'sheet_setup';
    if (stageNumber <= 3) return 'header';
    if (stageNumber <= 7) return 'data';
    if (stageNumber <= 9) return 'formula';
    return 'formatting';
  }

  /**
   * Check if line represents an operation boundary
   */
  private isOperationBoundary(line: string): boolean {
    const boundaryPatterns = [
      /const\s+sheet\s*=/, // Sheet creation/selection
      /sheet\.getRange\(.*\)\.values\s*=/, // Data assignment
      /sheet\.getRange\(.*\)\.formulas\s*=/, // Formula assignment  
      /sheet\.getRange\(.*\)\.format/, // Formatting
      /await\s+context\.sync\(\)/, // Sync operations
      /\/\/\s*(HEADER|SECTION|STEP)/ // Comment markers
    ];

    return boundaryPatterns.some(pattern => pattern.test(line));
  }

  /**
   * Analyze operation type from line content
   */
  private analyzeOperationType(line: string): Partial<ExecutionOperation> {
    if (line.includes('worksheets.add') || line.includes('worksheets.getItem')) {
      return { type: 'sheet_setup', stage: 1, description: 'Sheet setup' };
    }
    
    if (line.includes('.format.') || line.includes('format.fill') || line.includes('format.font')) {
      return { type: 'formatting', stage: 4, description: 'Apply formatting' };
    }
    
    if (line.includes('.formulas =')) {
      return { type: 'formula', stage: 3, description: 'Add formulas' };
    }
    
    if (line.includes('.values =')) {
      // Determine if it's header or data based on content
      if (line.toUpperCase().includes('HEADER') || line.includes('format.font.bold')) {
        return { type: 'header', stage: 2, description: 'Create headers' };
      }
      return { type: 'data', stage: 2, description: 'Add data' };
    }

    return { type: 'data', stage: 2, description: 'General operation' };
  }

  /**
   * Finalize operation object
   */
  private finalizeOperation(
    operation: Partial<ExecutionOperation>, 
    lines: string[], 
    counter: number, 
    blockIndex: number
  ): ExecutionOperation {
    const codeBlock = this.wrapInExcelRun(lines);
    
    return {
      id: operation.id || `op_${counter}`,
      type: operation.type || 'data',
      code: codeBlock,
      dependencies: operation.dependencies || [],
      optional: operation.optional || false,
      description: operation.description || `Operation ${counter}`,
      stage: operation.stage || 1
    };
  }

  /**
   * Wrap operation lines in Excel.run() block
   */
  private wrapInExcelRun(lines: string[]): string {
    return `await Excel.run(async (context) => {
    ${lines.join('\n    ')}
    await context.sync();
});`;
  }

  /**
   * Get maximum stage number from operations
   */
  private getMaxStage(operations: ExecutionOperation[]): number {
    return Math.max(...operations.map(op => op.stage), 1);
  }

  /**
   * Execute operations sequentially with error handling
   */
  async executeOperations(operations: ExecutionOperation[]): Promise<ExecutionResult[]> {
    console.log(`🚀 Starting sequential execution of ${operations.length} operations`);
    this.startTime = Date.now();
    this.operations = operations;
    this.results = [];
    this.currentStage = 1;

    const maxStage = this.getMaxStage(operations);
    
    for (let stage = 1; stage <= maxStage; stage++) {
      this.currentStage = stage;
      const stageOperations = operations.filter(op => op.stage === stage);
      
      console.log(`📋 Executing Stage ${stage}: ${stageOperations.length} operations`);
      
      for (const operation of stageOperations) {
        const result = await this.executeOperation(operation);
        this.results.push(result);
        
        // Update progress
        this.updateProgress();
        
        // Check if we should continue after error
        if (!result.success && !this.config.continueOnError && !operation.optional) {
          console.error(`❌ Critical operation failed: ${operation.id}, stopping execution`);
          return this.results;
        }
      }
      
      // Stage validation if enabled
      if (this.config.validateEachStage) {
        await this.validateStage(stage);
      }
    }

    console.log(`✅ Sequential execution completed in ${Date.now() - this.startTime}ms`);
    this.logExecutionSummary();
    return this.results;
  }

  /**
   * Execute single operation with retry logic
   */
  private async executeOperation(operation: ExecutionOperation): Promise<ExecutionResult> {
    console.log(`🔄 Executing: ${operation.description} (${operation.id})`);
    const startTime = Date.now();
    let retryCount = 0;
    let lastError: string = '';

    while (retryCount <= this.config.maxRetries) {
      try {
        // Execute the operation code
        await this.runOperationCode(operation.code);
        
        const executionTime = Date.now() - startTime;
        console.log(`✅ Operation ${operation.id} completed in ${executionTime}ms`);
        
        return {
          operationId: operation.id,
          success: true,
          executionTime,
          retryCount
        };
        
      } catch (error) {
        retryCount++;
        lastError = error instanceof Error ? error.message : String(error);
        console.warn(`⚠️ Operation ${operation.id} failed (attempt ${retryCount}):`, lastError);
        
        if (retryCount <= this.config.maxRetries) {
          // Apply error correction before retry
          operation.code = this.applyErrorCorrection(operation.code, lastError);
          await new Promise(resolve => setTimeout(resolve, 1000 * retryCount)); // Exponential backoff
        }
      }
    }

    const executionTime = Date.now() - startTime;
    console.error(`❌ Operation ${operation.id} failed after ${retryCount} attempts`);
    
    return {
      operationId: operation.id,
      success: false,
      error: lastError,
      executionTime,
      retryCount
    };
  }

  /**
   * Execute operation code in Excel context
   */
  private async runOperationCode(code: string): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        // Wrap in timeout
        const timeoutId = setTimeout(() => {
          reject(new Error(`Operation timeout after ${this.config.operationTimeout}ms`));
        }, this.config.operationTimeout);

        // Ensure code is wrapped in Excel.run if needed
        const wrappedCode = code.includes("Excel.run") ? code : `
          await Excel.run(async (context) => {
            ${code}
            await context.sync();
          });
        `;

        // Execute the code using eval (necessary for Excel.js in web context)
        const executeCode = async () => {
          try {
            await eval(`(async () => { ${wrappedCode} })()`);
            clearTimeout(timeoutId);
            resolve();
          } catch (error) {
            clearTimeout(timeoutId);
            reject(error);
          }
        };

        executeCode();
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Apply error correction to operation code
   */
  private applyErrorCorrection(code: string, error: string): string {
    console.log('🔧 Applying error correction for:', error.substring(0, 100));
    
    let correctedCode = code;
    
    // Array dimension corrections
    if (error.includes('dimension') || error.includes('array')) {
      correctedCode = this.fixArrayDimensions(correctedCode);
    }
    
    // Sheet existence corrections
    if (error.includes('worksheet') || error.includes('sheet')) {
      correctedCode = this.fixSheetReferences(correctedCode);
    }
    
    return correctedCode;
  }

  /**
   * Fix array dimension issues
   */
  private fixArrayDimensions(code: string): string {
    return code
      .replace(/\.values\s*=\s*\[([^\[\]]+)\]/g, '.values = [[$1]]')
      .replace(/\.formulas\s*=\s*\[([^\[\]]+)\]/g, '.formulas = [[$1]]')
      .replace(/\.values\s*=\s*"([^"]+)"/g, '.values = [["$1"]]')
      .replace(/\.formulas\s*=\s*"([^"]+)"/g, '.formulas = [["$1"]]');
  }

  /**
   * Fix sheet reference issues
   */
  private fixSheetReferences(code: string): string {
    // Add try/catch for sheet operations
    return code.replace(
      /(const\s+sheet\s*=\s*context\.workbook\.worksheets\.)getItem\("([^"]+)"\)/g,
      `$1getItem("$2") || $1add("$2")`
    );
  }

  /**
   * Validate stage completion
   */
  private async validateStage(stage: number): Promise<boolean> {
    console.log(`🔍 Validating stage ${stage} completion...`);
    // Implementation would check Excel state for expected results
    return true;
  }

  /**
   * Update execution progress
   */
  private updateProgress(): void {
    const progress: ExecutionProgress = {
      currentStage: this.currentStage,
      totalStages: this.getMaxStage(this.operations),
      currentOperation: this.results[this.results.length - 1]?.operationId || '',
      operationsCompleted: this.results.length,
      totalOperations: this.operations.length,
      errors: this.results.filter(r => !r.success),
      timeElapsed: Date.now() - this.startTime
    };

    if (this.config.progressCallback) {
      this.config.progressCallback(progress);
    }
  }

  /**
   * Log execution summary
   */
  private logExecutionSummary(): void {
    const successful = this.results.filter(r => r.success).length;
    const failed = this.results.filter(r => !r.success).length;
    const totalTime = Date.now() - this.startTime;

    console.log('📊 Execution Summary:');
    console.log(`   ✅ Successful: ${successful}`);
    console.log(`   ❌ Failed: ${failed}`);
    console.log(`   ⏱️  Total Time: ${totalTime}ms`);
    console.log(`   📈 Success Rate: ${((successful / this.results.length) * 100).toFixed(1)}%`);
  }

  /**
   * Get execution statistics
   */
  getExecutionStats() {
    const successful = this.results.filter(r => r.success).length;
    const failed = this.results.filter(r => !r.success).length;
    
    return {
      totalOperations: this.results.length,
      successful,
      failed,
      successRate: this.results.length > 0 ? successful / this.results.length : 0,
      totalTime: this.results.length > 0 ? Date.now() - this.startTime : 0,
      averageOperationTime: this.results.length > 0 ? 
        this.results.reduce((sum, r) => sum + r.executionTime, 0) / this.results.length : 0
    };
  }
}

// Export singleton instance
export const sequentialEngine = new SequentialExecutionEngine();