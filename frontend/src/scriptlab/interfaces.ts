export interface ISnippet {
  id: string;
  name: string;
  description?: string;
  script: string;
  style?: string;
  template?: string;
  libraries?: string[];
  created_at: number;
  last_modified: number;
  host?: string;
  api_set?: string;
}

export interface IExecutionResult {
  success: boolean;
  result?: any;
  error?: string;
  logs?: string[];
  executionTime?: number;
}

export interface ICompilerOptions {
  target?: string;
  module?: string;
  lib?: string[];
  moduleResolution?: string;
  allowJs?: boolean;
  checkJs?: boolean;
  noEmit?: boolean;
  strict?: boolean;
}

export interface IScriptLabEngineOptions {
  compilerOptions?: ICompilerOptions;
  timeout?: number;
  allowedLibraries?: string[];
  sandboxMode?: boolean;
}

export interface IOfficeHost {
  type: 'Excel' | 'Word' | 'PowerPoint' | 'Outlook' | 'OneNote';
  version: string;
  platform: 'PC' | 'Mac' | 'iOS' | 'Android' | 'Universal' | 'Web';
}

export interface ILibraryReference {
  url: string;
  name: string;
  version?: string;
  description?: string;
}