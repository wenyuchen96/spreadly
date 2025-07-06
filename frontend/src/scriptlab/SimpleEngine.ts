import { ISnippet, IExecutionResult } from './interfaces';
import { OfficeJsHost } from './officeJsHost';

/**
 * Simplified Script Lab engine for testing without TypeScript compilation
 * This version executes JavaScript/TypeScript code directly without compilation
 */
export class SimpleScriptLabEngine {
  private executionFrame: HTMLIFrameElement | null = null;

  async executeSnippet(snippet: ISnippet): Promise<IExecutionResult> {
    const startTime = Date.now();
    
    try {
      // Wait for Office to be ready
      await OfficeJsHost.waitForOfficeReady();

      // Execute the code directly (modern browsers support most TypeScript syntax)
      const result = await this.executeCode(snippet.script, snippet.template, snippet.style, snippet.libraries);
      
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
          console.log('Received message from iframe:', event.data);
          
          if (event.source === this.executionFrame?.contentWindow) {
            window.removeEventListener('message', messageHandler);
            this.cleanupExecutionFrame();
            
            if (event.data.type === 'execution-result') {
              console.log('Execution successful:', event.data);
              resolve({
                success: event.data.success,
                result: event.data.result,
                error: event.data.error,
                logs: event.data.logs
              });
            } else if (event.data.type === 'execution-error') {
              console.log('Execution error:', event.data);
              resolve({
                success: false,
                error: event.data.error,
                logs: event.data.logs
              });
            }
          }
        };

        window.addEventListener('message', messageHandler);

        // Set timeout (30 seconds for testing)
        const timeoutId = setTimeout(() => {
          console.log('Execution timed out after 30 seconds');
          window.removeEventListener('message', messageHandler);
          this.cleanupExecutionFrame();
          resolve({
            success: false,
            error: 'Execution timeout (30s) - check browser console for iframe errors'
          });
        }, 30000);

        // Load and execute
        console.log('Creating iframe for code execution...');
        document.body.appendChild(this.executionFrame);
        
        console.log('Setting iframe content...');
        this.executionFrame.srcdoc = executionHtml;
        
        console.log('Iframe created, waiting for execution...');

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
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 10px; }
        ${css}
    </style>
</head>
<body>
    ${html}
    <div id="output">Ready to execute...</div>
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

        // Add immediate debug message
        console.log('Iframe script loaded, waiting for Office.js...');
        
        // Check if Office is available
        if (typeof Office === 'undefined') {
            console.error('Office.js not available in iframe');
            parent.postMessage({
                type: 'execution-error',
                error: 'Office.js not available in iframe',
                logs: logs
            }, '*');
            return;
        }

        Office.onReady((info) => {
            try {
                console.log('Office.js ready in iframe, host:', info ? info.host : 'unknown');
                
                // Check Excel availability
                if (typeof Excel === 'undefined') {
                    console.error('Excel API not available');
                    parent.postMessage({
                        type: 'execution-error',
                        error: 'Excel API not available in iframe',
                        logs: logs
                    }, '*');
                    return;
                }
                
                console.log('Excel API available, executing code...');
                
                // Execute the user code
                const executeUserCode = async () => {
                    try {
                        const outputDiv = document.getElementById('output');
                        if (outputDiv) outputDiv.innerHTML = 'Executing code...';
                        
                        console.log('Starting code execution...');
                        
                        // Execute the code (wrapped in async function)
                        const asyncWrapper = async () => {
                            ${code}
                        };
                        
                        await asyncWrapper();
                        
                        console.log('Code execution completed successfully');
                        
                        if (outputDiv) outputDiv.innerHTML = 'Code executed successfully!';
                        
                        parent.postMessage({
                            type: 'execution-result',
                            success: true,
                            result: 'Code executed successfully',
                            logs: logs
                        }, '*');
                    } catch (error) {
                        console.error('Code execution error:', error);
                        const outputDiv = document.getElementById('output');
                        if (outputDiv) outputDiv.innerHTML = 'Error: ' + error.message;
                        
                        parent.postMessage({
                            type: 'execution-error',
                            error: error.message,
                            logs: logs
                        }, '*');
                    }
                };
                
                // Add small delay to ensure everything is ready
                setTimeout(() => {
                    executeUserCode();
                }, 1000);
                
            } catch (error) {
                console.error('Office.onReady error:', error);
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

  dispose(): void {
    this.cleanupExecutionFrame();
  }
}