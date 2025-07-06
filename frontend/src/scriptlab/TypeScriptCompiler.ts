import { ICompilerOptions } from './interfaces';

declare const ts: any;

export class TypeScriptCompiler {
  private compilerOptions: ICompilerOptions;
  private isTypeScriptLoaded: boolean = false;

  constructor(options: ICompilerOptions = {}) {
    this.compilerOptions = {
      target: 'ES2017',
      module: 'ES2015',
      lib: ['ES2017', 'DOM'],
      moduleResolution: 'node',
      allowJs: true,
      checkJs: false,
      noEmit: true,
      strict: false,
      ...options
    };
  }

  private async loadTypeScript(): Promise<void> {
    if (this.isTypeScriptLoaded) {
      return;
    }

    return new Promise((resolve, reject) => {
      if (typeof ts !== 'undefined') {
        this.isTypeScriptLoaded = true;
        resolve();
        return;
      }

      // Load TypeScript from CDN
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/typescript@4.9.5/lib/typescript.min.js';
      script.onload = () => {
        this.isTypeScriptLoaded = true;
        resolve();
      };
      script.onerror = (error) => {
        reject(new Error('Failed to load TypeScript compiler'));
      };
      document.head.appendChild(script);
    });
  }

  async compile(code: string): Promise<{ success: boolean; result?: string; errors?: string[] }> {
    try {
      await this.loadTypeScript();

      // Check if TypeScript is actually available
      if (typeof ts === 'undefined') {
        return {
          success: false,
          errors: ['TypeScript compiler not available']
        };
      }

      // Simple transpilation with minimal options
      const compilerOptions = {
        target: ts.ScriptTarget.ES2017 || ts.ScriptTarget.ES2015,
        module: ts.ModuleKind.None,
        allowJs: true,
        strict: false
      };

      const result = ts.transpile(code, compilerOptions);
      
      if (result && result.trim()) {
        return {
          success: true,
          result: result
        };
      } else {
        return {
          success: false,
          errors: ['Compilation produced empty result']
        };
      }
    } catch (error) {
      return {
        success: false,
        errors: [error instanceof Error ? error.message : 'Unknown compilation error']
      };
    }
  }

  async validateSyntax(code: string): Promise<{ isValid: boolean; errors?: string[] }> {
    try {
      // Simple syntax validation - just try to compile and catch errors
      const compileResult = await this.compile(code);
      
      if (compileResult.success) {
        return { isValid: true };
      } else {
        return {
          isValid: false,
          errors: compileResult.errors
        };
      }
    } catch (error) {
      return {
        isValid: false,
        errors: [error instanceof Error ? error.message : 'Unknown syntax validation error']
      };
    }
  }

  static isTypeScriptCode(code: string): boolean {
    // Simple heuristic to detect TypeScript-specific syntax
    const tsPatterns = [
      /:\s*(string|number|boolean|any|void|object|Function)/,
      /interface\s+\w+/,
      /type\s+\w+\s*=/,
      /enum\s+\w+/,
      /public\s+|private\s+|protected\s+/,
      /readonly\s+/,
      /<\w+>/,
      /as\s+\w+/
    ];

    return tsPatterns.some(pattern => pattern.test(code));
  }
}