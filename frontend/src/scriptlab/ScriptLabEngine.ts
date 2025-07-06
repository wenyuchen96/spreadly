import { ISnippet, IExecutionResult, IScriptLabEngineOptions, ILibraryReference } from './interfaces';
import { TypeScriptCompiler } from './TypeScriptCompiler';
import { OfficeJsHost } from './officeJsHost';

export class ScriptLabEngine {
  private compiler: TypeScriptCompiler;
  private options: IScriptLabEngineOptions;
  private executionFrame: HTMLIFrameElement | null = null;
  private executionPromise: Promise<IExecutionResult> | null = null;

  constructor(options: IScriptLabEngineOptions = {}) {
    this.options = {
      timeout: 30000, // 30 seconds default timeout
      allowedLibraries: [],
      sandboxMode: true,
      ...options
    };

    this.compiler = new TypeScriptCompiler(options.compilerOptions);
  }

  async executeSnippet(snippet: ISnippet): Promise<IExecutionResult> {
    const startTime = Date.now();
    
    try {
      // Wait for Office to be ready
      await OfficeJsHost.waitForOfficeReady();

      // Try to validate the code, but don't fail if validation fails
      try {
        const validation = await this.compiler.validateSyntax(snippet.script);
        if (!validation.isValid) {
          console.warn('Code validation warnings:', validation.errors);
          // Continue anyway - sometimes valid Office.js code fails TypeScript validation
        }
      } catch (validationError) {
        console.warn('Code validation not available, proceeding without validation');
      }

      // Compile TypeScript to JavaScript if needed
      let executableCode = snippet.script;
      if (TypeScriptCompiler.isTypeScriptCode(snippet.script)) {
        try {
          const compileResult = await this.compiler.compile(snippet.script);
          if (compileResult.success && compileResult.result) {
            executableCode = compileResult.result;
          } else {
            console.warn('TypeScript compilation failed, using original code:', compileResult.errors);
            // Fallback: use original code (many TypeScript features work in modern browsers)
            executableCode = snippet.script;
          }
        } catch (compileError) {
          console.warn('TypeScript compiler not available, using original code');
          executableCode = snippet.script;
        }
      }

      // Execute the code
      const result = await this.executeCode(executableCode, snippet.template, snippet.style, snippet.libraries);
      
      return {
        ...result,
        executionTime: Date.now() - startTime
      };

    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown execution error',
        executionTime: Date.now() - startTime
      };
    }
  }

  async executeCode(
    code: string, 
    html: string = '', 
    css: string = '', 
    libraries: string[] = []
  ): Promise<IExecutionResult> {
    return new Promise((resolve) => {
      try {
        // Create execution environment
        const executionHtml = this.createExecutionHtml(code, html, css, libraries);
        
        // Create iframe for isolated execution
        this.executionFrame = document.createElement('iframe');
        this.executionFrame.style.display = 'none';
        this.executionFrame.sandbox.add('allow-scripts', 'allow-same-origin');
        
        // Set up communication with iframe
        const messageHandler = (event: MessageEvent) => {
          if (event.source === this.executionFrame?.contentWindow) {
            window.removeEventListener('message', messageHandler);
            this.cleanupExecutionFrame();
            
            if (event.data.type === 'execution-result') {
              resolve({
                success: event.data.success,
                result: event.data.result,
                error: event.data.error,
                logs: event.data.logs
              });
            } else if (event.data.type === 'execution-error') {
              resolve({
                success: false,
                error: event.data.error,
                logs: event.data.logs
              });
            }
          }
        };

        window.addEventListener('message', messageHandler);

        // Set timeout
        const timeoutId = setTimeout(() => {
          window.removeEventListener('message', messageHandler);
          this.cleanupExecutionFrame();
          resolve({
            success: false,
            error: 'Execution timeout'
          });
        }, this.options.timeout);

        // Load and execute
        document.body.appendChild(this.executionFrame);
        this.executionFrame.srcdoc = executionHtml;

        // Clear timeout when done
        const originalHandler = messageHandler;
        const wrappedHandler = (event: MessageEvent) => {
          clearTimeout(timeoutId);
          originalHandler(event);
        };
        window.removeEventListener('message', messageHandler);
        window.addEventListener('message', wrappedHandler);

      } catch (error) {
        this.cleanupExecutionFrame();
        resolve({
          success: false,
          error: error instanceof Error ? error.message : 'Unknown execution error'
        });
      }
    });
  }

  private createExecutionHtml(code: string, html: string, css: string, libraries: string[] = []): string {
    const libraryTags = libraries.map(lib => `<script src="${lib}"></script>`).join('\n');
    
    return `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ${libraryTags}
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        ${css}
    </style>
</head>
<body>
    ${html}
    <script>
        const logs = [];
        const originalConsoleLog = console.log;
        const originalConsoleError = console.error;
        const originalConsoleWarn = console.warn;
        
        console.log = (...args) => {
            logs.push({ type: 'log', message: args.join(' ') });
            originalConsoleLog.apply(console, args);
        };
        
        console.error = (...args) => {
            logs.push({ type: 'error', message: args.join(' ') });
            originalConsoleError.apply(console, args);
        };
        
        console.warn = (...args) => {
            logs.push({ type: 'warn', message: args.join(' ') });
            originalConsoleWarn.apply(console, args);
        };

        window.onerror = (message, source, lineno, colno, error) => {
            parent.postMessage({
                type: 'execution-error',
                error: message,
                logs: logs
            }, '*');
        };

        Office.onReady(() => {
            try {
                // Make Office APIs available globally
                window.Excel = Excel;
                window.Word = Word;
                window.PowerPoint = PowerPoint;
                window.OneNote = OneNote;
                window.Outlook = Outlook;
                
                // Execute the user code
                const executeUserCode = async () => {
                    try {
                        ${code}
                        
                        parent.postMessage({
                            type: 'execution-result',
                            success: true,
                            result: 'Code executed successfully',
                            logs: logs
                        }, '*');
                    } catch (error) {
                        parent.postMessage({
                            type: 'execution-error',
                            error: error.message,
                            logs: logs
                        }, '*');
                    }
                };
                
                executeUserCode();
            } catch (error) {
                parent.postMessage({
                    type: 'execution-error',
                    error: error.message,
                    logs: logs
                }, '*');
            }
        });
    </script>
</body>
</html>`;
  }

  private cleanupExecutionFrame(): void {
    if (this.executionFrame && this.executionFrame.parentNode) {
      this.executionFrame.parentNode.removeChild(this.executionFrame);
    }
    this.executionFrame = null;
  }

  async validateCode(code: string): Promise<{ isValid: boolean; errors?: string[] }> {
    return await this.compiler.validateSyntax(code);
  }

  createSnippet(
    name: string,
    script: string,
    template: string = '',
    style: string = '',
    libraries: string[] = [],
    description: string = ''
  ): ISnippet {
    return {
      id: this.generateId(),
      name,
      description,
      script,
      style,
      template,
      libraries,
      created_at: Date.now(),
      last_modified: Date.now(),
      host: 'Excel',
      api_set: 'ExcelApi 1.1'
    };
  }

  private generateId(): string {
    return 'snippet_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  // Static helper methods for common Excel operations
  static createExcelSnippet(code: string, description: string = ''): ISnippet {
    const engine = new ScriptLabEngine();
    return engine.createSnippet(
      'Excel Operation',
      code,
      '<div id="output"></div>',
      'body { padding: 20px; }',
      [],
      description
    );
  }

  // Dispose method to clean up resources
  dispose(): void {
    this.cleanupExecutionFrame();
    this.executionPromise = null;
  }
}