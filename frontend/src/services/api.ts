/**
 * API client for communicating with the Spreadly backend
 */

interface UploadResponse {
  session_token: string;
  message: string;
  data: any;
}

interface AnalysisResponse {
  session_token: string;
  analysis: any;
  insights: string;
}

interface QueryResponse {
  session_token: string;
  query: string;
  result: any;
}

interface FormulaResponse {
  description: string;
  formulas: Array<{
    formula: string;
    description: string;
    difficulty: string;
    example: string;
  }>;
}

class SpreadlyAPI {
  private baseUrl: string;
  private sessionToken: string | null = null;

  constructor(baseUrl: string = 'http://127.0.0.1:8000') {
    this.baseUrl = baseUrl;
  }

  /**
   * Upload Excel data to backend for processing
   */
  async uploadExcelData(data: any[][], fileName: string): Promise<UploadResponse> {
    try {
      // Convert Excel data to a simple JSON format for the backend
      const payload = {
        file_name: fileName,
        data: data,
        sheets: [{
          name: 'Sheet1',
          data: data
        }]
      };

      const response = await fetch(`${this.baseUrl}/api/excel/upload`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      this.sessionToken = result.session_token;
      return result;
    } catch (error) {
      console.error('Error uploading Excel data:', error);
      throw error;
    }
  }

  /**
   * Get AI analysis of uploaded spreadsheet
   */
  async getAnalysis(sessionToken?: string): Promise<AnalysisResponse> {
    const token = sessionToken || this.sessionToken;
    if (!token) {
      throw new Error('No session token available');
    }

    try {
      const response = await fetch(`${this.baseUrl}/api/excel/analyze/${token}`);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error getting analysis:', error);
      throw error;
    }
  }

  /**
   * Send natural language query about spreadsheet
   */
  async sendQuery(query: string, sessionToken?: string): Promise<QueryResponse> {
    const token = sessionToken || this.sessionToken;
    if (!token) {
      throw new Error('No session token available');
    }

    try {
      const response = await fetch(`${this.baseUrl}/api/excel/query`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          session_token: token,
          query: query
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error sending query:', error);
      throw error;
    }
  }

  /**
   * Generate Excel formulas from natural language description
   */
  async generateFormulas(description: string, context?: string): Promise<FormulaResponse> {
    try {
      const params = new URLSearchParams({
        description: description
      });
      
      if (context) {
        params.append('context', context);
      }

      const response = await fetch(`${this.baseUrl}/api/excel/formulas?${params}`);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Error generating formulas:', error);
      throw error;
    }
  }

  /**
   * Check if backend is available
   */
  async healthCheck(): Promise<boolean> {
    try {
      console.log('Checking backend health at:', `${this.baseUrl}/health`);
      
      // Use a shorter timeout for Excel Add-in environment
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 5000); // 5 second timeout
      
      const response = await fetch(`${this.baseUrl}/health`, {
        method: 'GET',
        mode: 'cors',
        headers: {
          'Content-Type': 'application/json',
        },
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      console.log('Health check response:', response.status, response.ok);
      return response.ok;
    } catch (error) {
      console.error('Backend health check failed:', error);
      console.error('Error details:', {
        message: error instanceof Error ? error.message : 'Unknown error',
        baseUrl: this.baseUrl,
        name: error instanceof Error ? error.name : 'Unknown'
      });
      return false;
    }
  }

  /**
   * Get current session token
   */
  getSessionToken(): string | null {
    return this.sessionToken;
  }

  /**
   * Set session token manually
   */
  setSessionToken(token: string): void {
    this.sessionToken = token;
  }
}

export const spreadlyAPI = new SpreadlyAPI();
export type { UploadResponse, AnalysisResponse, QueryResponse, FormulaResponse };