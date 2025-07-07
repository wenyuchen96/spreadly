/**
 * API client using Office.js Dialog API to bypass network restrictions
 */

declare const Office: any;

interface DialogResult {
  success: boolean;
  data?: any;
  error?: string;
}

class DialogAPI {
  private dialogUrl: string;

  constructor() {
    // In production, this would be your deployed URL
    this.dialogUrl = window.location.origin;
  }

  /**
   * Make API call through dialog window
   */
  private async makeDialogCall(callType: string, data?: string): Promise<DialogResult> {
    return new Promise((resolve, reject) => {
      const url = `${this.dialogUrl}/dialog/api-bridge.html?call=${callType}&data=${encodeURIComponent(data || '')}`;
      
      Office.context.ui.displayDialogAsync(url, {
        height: 30,
        width: 20,
        displayInIframe: false
      }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        const dialog = result.value;
        
        // Handle message from dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          try {
            const response = JSON.parse(arg.message);
            dialog.close();
            resolve(response);
          } catch (error) {
            dialog.close();
            reject(new Error('Failed to parse dialog response'));
          }
        });

        // Handle dialog errors
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
          dialog.close();
          reject(new Error(`Dialog error: ${arg.error}`));
        });
      });
    });
  }

  /**
   * Health check
   */
  async healthCheck(): Promise<boolean> {
    try {
      const result = await this.makeDialogCall('health');
      return result.success;
    } catch (error) {
      console.error('Dialog health check failed:', error);
      return false;
    }
  }

  /**
   * Generate formulas
   */
  async generateFormulas(description: string): Promise<any> {
    try {
      const result = await this.makeDialogCall('formulas', description);
      if (result.success) {
        return result.data;
      } else {
        throw new Error(result.error || 'Formula generation failed');
      }
    } catch (error) {
      console.error('Dialog formula generation failed:', error);
      throw error;
    }
  }

  /**
   * Upload Excel data
   */
  async uploadData(data: any): Promise<any> {
    try {
      const result = await this.makeDialogCall('upload', JSON.stringify(data));
      if (result.success) {
        return result.data;
      } else {
        throw new Error(result.error || 'Data upload failed');
      }
    } catch (error) {
      console.error('Dialog data upload failed:', error);
      throw error;
    }
  }

  /**
   * Send query
   */
  async sendQuery(query: string, sessionToken: string): Promise<any> {
    try {
      const requestData = { query, session_token: sessionToken };
      const result = await this.makeDialogCall('query', JSON.stringify(requestData));
      if (result.success) {
        return result.data;
      } else {
        throw new Error(result.error || 'Query failed');
      }
    } catch (error) {
      console.error('Dialog query failed:', error);
      throw error;
    }
  }
}

export const dialogAPI = new DialogAPI();