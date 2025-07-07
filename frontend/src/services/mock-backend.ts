/**
 * Mock backend service for when real backend is not accessible
 * Provides realistic AI-like responses for development and testing
 */

export interface MockFormulaResponse {
  formulas: Array<{
    formula: string;
    description: string;
    difficulty: 'basic' | 'intermediate' | 'advanced';
    example?: string;
  }>;
}

export interface MockAnalysisResponse {
  insights: string[];
  suggestions: string[];
  formulas?: Array<{
    formula: string;
    description: string;
  }>;
}

export class MockBackendService {
  private sessionToken: string | null = null;

  /**
   * Simulate backend health check
   */
  async healthCheck(): Promise<boolean> {
    // Simulate network delay
    await this.delay(500);
    return true;
  }

  /**
   * Generate realistic formulas based on description
   */
  async generateFormulas(description: string): Promise<MockFormulaResponse> {
    await this.delay(1000); // Simulate AI processing time
    
    const lowerDesc = description.toLowerCase();
    
    if (lowerDesc.includes('percentage') || lowerDesc.includes('percent')) {
      return {
        formulas: [
          {
            formula: '=(B2-A2)/A2*100',
            description: 'Calculate percentage change between two values',
            difficulty: 'basic',
            example: 'If A2=100 and B2=120, result is 20%'
          },
          {
            formula: '=B2/A2*100',
            description: 'Calculate percentage of total',
            difficulty: 'basic',
            example: 'If A2=total and B2=part, shows percentage'
          },
          {
            formula: '=ROUND((B2-A2)/A2*100,2)',
            description: 'Percentage change rounded to 2 decimals',
            difficulty: 'intermediate',
            example: 'More precise percentage calculation'
          }
        ]
      };
    }
    
    if (lowerDesc.includes('growth') || lowerDesc.includes('rate')) {
      return {
        formulas: [
          {
            formula: '=POWER(B2/A2,1/C2)-1',
            description: 'Compound Annual Growth Rate (CAGR)',
            difficulty: 'advanced',
            example: 'A2=start value, B2=end value, C2=years'
          },
          {
            formula: '=(B2-A2)/A2',
            description: 'Simple growth rate',
            difficulty: 'basic',
            example: 'Shows decimal growth (0.2 = 20%)'
          }
        ]
      };
    }
    
    if (lowerDesc.includes('average') || lowerDesc.includes('mean')) {
      return {
        formulas: [
          {
            formula: '=AVERAGE(A1:A10)',
            description: 'Calculate average of a range',
            difficulty: 'basic',
            example: 'Average of numbers in A1 through A10'
          },
          {
            formula: '=AVERAGEIF(A1:A10,">0")',
            description: 'Average of positive values only',
            difficulty: 'intermediate',
            example: 'Excludes zero and negative values'
          }
        ]
      };
    }
    
    if (lowerDesc.includes('sum') || lowerDesc.includes('total')) {
      return {
        formulas: [
          {
            formula: '=SUM(A1:A10)',
            description: 'Sum values in a range',
            difficulty: 'basic',
            example: 'Adds all numbers from A1 to A10'
          },
          {
            formula: '=SUMIF(A1:A10,">100")',
            description: 'Sum values greater than 100',
            difficulty: 'intermediate',
            example: 'Conditional sum based on criteria'
          }
        ]
      };
    }
    
    // Default formulas for unknown descriptions
    return {
      formulas: [
        {
          formula: '=SUM(A1:A10)',
          description: 'Sum of values (generic formula)',
          difficulty: 'basic',
          example: 'Basic calculation formula'
        },
        {
          formula: '=AVERAGE(A1:A10)',
          description: 'Average of values',
          difficulty: 'basic',
          example: 'Calculate mean value'
        },
        {
          formula: '=MAX(A1:A10)-MIN(A1:A10)',
          description: 'Range (difference between max and min)',
          difficulty: 'intermediate',
          example: 'Measures data spread'
        }
      ]
    };
  }

  /**
   * Upload data simulation
   */
  async uploadData(data: any[], fileName: string): Promise<{ session_token: string; message: string }> {
    await this.delay(1500); // Simulate upload time
    
    this.sessionToken = 'mock_session_' + Date.now();
    
    return {
      session_token: this.sessionToken,
      message: `✅ Mock upload successful! Processed ${data.length} rows from ${fileName}`
    };
  }

  /**
   * Get analysis simulation
   */
  async getAnalysis(): Promise<{ analysis: MockAnalysisResponse }> {
    await this.delay(2000); // Simulate AI analysis time
    
    const insights = [
      'Your data shows a clear upward trend over time',
      'There are some potential outliers in column C that may need attention',
      'The correlation between columns A and B appears to be strong',
      'Consider using conditional formatting to highlight key values',
      'Your data structure is well-organized for further analysis'
    ];
    
    const suggestions = [
      'Create a pivot table to summarize your data by categories',
      'Add data validation to prevent future input errors',
      'Use charts to visualize trends more clearly',
      'Consider adding calculated columns for better insights',
      'Apply filters to focus on specific data subsets'
    ];
    
    return {
      analysis: {
        insights: insights.slice(0, 3), // Return 3 random insights
        suggestions: suggestions.slice(0, 3), // Return 3 random suggestions
        formulas: [
          {
            formula: '=TREND(B2:B10,A2:A10)',
            description: 'Calculate trend line for your data'
          }
        ]
      }
    };
  }

  /**
   * Process natural language query
   */
  async processQuery(query: string): Promise<{ result: { answer: string; formula?: string } }> {
    await this.delay(1200); // Simulate AI processing
    
    const lowerQuery = query.toLowerCase();
    
    if (lowerQuery.includes('highest') || lowerQuery.includes('maximum') || lowerQuery.includes('max')) {
      return {
        result: {
          answer: 'To find the highest value in your data, you can use the MAX function. This will return the largest number in the specified range.',
          formula: '=MAX(A1:A10)'
        }
      };
    }
    
    if (lowerQuery.includes('lowest') || lowerQuery.includes('minimum') || lowerQuery.includes('min')) {
      return {
        result: {
          answer: 'To find the lowest value, use the MIN function. This returns the smallest number in your range.',
          formula: '=MIN(A1:A10)'
        }
      };
    }
    
    if (lowerQuery.includes('count') || lowerQuery.includes('how many')) {
      return {
        result: {
          answer: 'To count non-empty cells, use COUNTA. For counting cells with numbers only, use COUNT.',
          formula: '=COUNTA(A1:A10)'
        }
      };
    }
    
    // Generic response
    return {
      result: {
        answer: `I understand you're asking about: "${query}". Based on your spreadsheet data, I recommend exploring this with formulas or charts. This is a mock AI response - in production, Claude AI would provide much more detailed and specific insights.`
      }
    };
  }

  /**
   * Get session token
   */
  getSessionToken(): string | null {
    return this.sessionToken;
  }

  /**
   * Simulate network delay
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

export const mockBackend = new MockBackendService();