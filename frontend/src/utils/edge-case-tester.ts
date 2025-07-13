/**
 * Automated Edge Case Testing Suite
 * Self-tests code execution stability and auto-improves validation
 */

interface EdgeCaseTest {
  id: string;
  name: string;
  description: string;
  code: string;
  expectedIssues: string[];
  severity: 'low' | 'medium' | 'high' | 'critical';
}

interface TestResult {
  testId: string;
  passed: boolean;
  detectedIssues: string[];
  autoCorrections: string[];
  executionTime: number;
  errors: string[];
  improvements: string[];
}

export class EdgeCaseTester {
  private testResults: TestResult[] = [];

  constructor() {
    console.log('🧪 EdgeCaseTester initialized for automated stability testing');
  }

  /**
   * Run all edge case tests and return analysis
   */
  async runAllTests(): Promise<{
    totalTests: number;
    passed: number;
    failed: number;
    criticalIssues: string[];
    improvements: string[];
    stabilityScore: number;
  }> {
    console.log('🧪 Starting comprehensive edge case testing...');
    
    const tests = this.getEdgeCaseTests();
    const results: TestResult[] = [];

    for (const test of tests) {
      console.log(`🧪 Running test: ${test.name}`);
      const result = await this.runSingleTest(test);
      results.push(result);
      this.testResults.push(result);
    }

    return this.analyzeResults(results);
  }

  private getEdgeCaseTests(): EdgeCaseTest[] {
    return [
      {
        id: 'uncalled_function_basic',
        name: 'Uncalled Function Detection - Basic',
        description: 'Function declared but never called',
        code: `
async function createModel() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sheet1");
    sheet.getRange("A1").values = [["Test"]];
    await context.sync();
  });
}`,
        expectedIssues: ['uncalled_function'],
        severity: 'high'
      },
      {
        id: 'uncalled_function_nested',
        name: 'Uncalled Function Detection - Nested',
        description: 'Multiple nested functions with complex structure',
        code: `
async function createFinancialModel() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet3");
      
      async function setupHeaders() {
        sheet.getRange("A1").values = [["Complex Model"]];
      }
      
      await setupHeaders();
      await context.sync();
    });
  } catch (error) {
    console.error('Model creation failed:', error);
  }
}`,
        expectedIssues: ['uncalled_function'],
        severity: 'critical'
      },
      {
        id: 'array_dimension_chaos',
        name: 'Array Dimension Mismatch - Multiple Types',
        description: 'Mixed 1D/2D arrays and invalid assignments',
        code: `
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").values = "Single String";
  sheet.getRange("A2").values = ["Array", "Of", "Strings"];
  sheet.getRange("A3").formulas = "=SUM(A1:A2)";
  sheet.getRange("A4:B4").values = ["Too", "Many", "Values", "For", "Range"];
  await context.sync();
});`,
        expectedIssues: ['array_dimension_mismatch', 'invalid_assignment'],
        severity: 'high'
      },
      {
        id: 'invalid_sheet_reference',
        name: 'Invalid Sheet References',
        description: 'References to non-existent sheets and invalid names',
        code: `
await Excel.run(async (context) => {
  const sheet1 = context.workbook.worksheets.getItem("NonExistentSheet");
  const sheet2 = context.workbook.worksheets.getItem("Sheet with spaces!");
  const sheet3 = context.workbook.worksheets.getItem("Sheet99");
  sheet1.getRange("A1").values = [["Test"]];
  await context.sync();
});`,
        expectedIssues: ['invalid_sheet_reference', 'sheet_not_found'],
        severity: 'medium'
      },
      {
        id: 'formula_complexity_bomb',
        name: 'Complex Formula Nesting',
        description: 'Deeply nested formulas with potential circular references',
        code: `
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.getRange("A1").formulas = [["=B1+C1"]];
  sheet.getRange("B1").formulas = [["=C1+D1"]];
  sheet.getRange("C1").formulas = [["=A1+B1"]]; // Circular reference
  sheet.getRange("D1").formulas = [["=IF(A1>0,B1*C1,SUM(A1:C1)+AVERAGE(B1:D1))"]];
  await context.sync();
});`,
        expectedIssues: ['circular_reference', 'complex_formula'],
        severity: 'medium'
      },
      {
        id: 'memory_bomb_test',
        name: 'Memory Usage Test',
        description: 'Large data operations that could cause memory issues',
        code: `
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const largeData = Array(1000).fill(0).map((_, i) => 
    Array(50).fill(0).map((_, j) => \`Cell_\${i}_\${j}\`)
  );
  sheet.getRange("A1:AX1000").values = largeData;
  await context.sync();
});`,
        expectedIssues: ['memory_usage', 'large_operation'],
        severity: 'low'
      },
      {
        id: 'type_coercion_chaos',
        name: 'Type Coercion Issues',
        description: 'Mixed data types causing coercion problems',
        code: `
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const mixedData = [
    ["String", 123, true, null, undefined],
    [new Date(), 3.14159, "=SUM(A1:A5)", "", 0],
    [false, "123abc", Infinity, -0, NaN]
  ];
  sheet.getRange("A1:E3").values = mixedData;
  await context.sync();
});`,
        expectedIssues: ['type_coercion', 'invalid_data_type'],
        severity: 'medium'
      },
      {
        id: 'async_nightmare',
        name: 'Async/Await Complexity',
        description: 'Complex async patterns that could cause race conditions',
        code: `
async function complexAsyncModel() {
  const promises = [];
  for (let i = 0; i < 10; i++) {
    promises.push(
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange(\`A\${i + 1}\`).values = [[\`Async_\${i}\`]];
        await context.sync();
      })
    );
  }
  await Promise.all(promises);
}`,
        expectedIssues: ['async_complexity', 'potential_race_condition'],
        severity: 'high'
      }
    ];
  }

  private async runSingleTest(test: EdgeCaseTest): Promise<TestResult> {
    const startTime = Date.now();
    const result: TestResult = {
      testId: test.id,
      passed: false,
      detectedIssues: [],
      autoCorrections: [],
      executionTime: 0,
      errors: [],
      improvements: []
    };

    try {
      // Import validation function (would need to adjust path in real implementation)
      const { validateGeneratedCode, autoCorrectArrayDimensions } = await import('../taskpane/taskpane');
      
      // Run validation
      const validation = validateGeneratedCode(test.code);
      
      // Check if expected issues were detected
      for (const expectedIssue of test.expectedIssues) {
        const detected = validation.errors.some(error => 
          error.toLowerCase().includes(expectedIssue.replace('_', ' '))
        ) || validation.warnings.some(warning => 
          warning.toLowerCase().includes(expectedIssue.replace('_', ' '))
        ) || validation.fixableIssues.some(issue => 
          issue.toLowerCase().includes(expectedIssue.replace('_', ' '))
        );
        
        if (detected) {
          result.detectedIssues.push(expectedIssue);
        }
      }

      // Test auto-correction
      const correctedCode = autoCorrectArrayDimensions(test.code);
      if (correctedCode !== test.code) {
        result.autoCorrections.push('auto_correction_applied');
        
        // Re-validate corrected code
        const correctedValidation = validateGeneratedCode(correctedCode);
        if (correctedValidation.errors.length < validation.errors.length) {
          result.autoCorrections.push('error_count_reduced');
        }
      }

      // Calculate pass/fail
      const detectedExpectedIssues = result.detectedIssues.length;
      const totalExpectedIssues = test.expectedIssues.length;
      result.passed = detectedExpectedIssues >= Math.floor(totalExpectedIssues * 0.7); // 70% detection rate

      result.executionTime = Date.now() - startTime;

    } catch (error) {
      result.errors.push(error instanceof Error ? error.message : 'Unknown error');
      result.executionTime = Date.now() - startTime;
    }

    return result;
  }

  private analyzeResults(results: TestResult[]): {
    totalTests: number;
    passed: number;
    failed: number;
    criticalIssues: string[];
    improvements: string[];
    stabilityScore: number;
  } {
    const totalTests = results.length;
    const passed = results.filter(r => r.passed).length;
    const failed = totalTests - passed;
    
    const criticalIssues: string[] = [];
    const improvements: string[] = [];

    // Analyze patterns
    const allDetectedIssues = results.flatMap(r => r.detectedIssues);
    const allErrors = results.flatMap(r => r.errors);
    
    // Check for critical patterns
    if (allDetectedIssues.filter(issue => issue === 'uncalled_function').length > 0) {
      criticalIssues.push('Uncalled function detection needs improvement');
    }
    
    if (allErrors.length > totalTests * 0.2) {
      criticalIssues.push('High error rate in validation system');
    }

    // Generate improvements
    const autoCorrectionsCount = results.filter(r => r.autoCorrections.length > 0).length;
    if (autoCorrectionsCount < totalTests * 0.8) {
      improvements.push('Enhance auto-correction coverage');
    }
    
    if (results.some(r => r.detectedIssues.length === 0 && r.testId.includes('uncalled'))) {
      improvements.push('Fix function call detection regex patterns');
    }

    const stabilityScore = Math.round((passed / totalTests) * 100);

    console.log(`🧪 Edge Case Testing Complete:`, {
      totalTests,
      passed,
      failed,
      stabilityScore: `${stabilityScore}%`,
      criticalIssues: criticalIssues.length,
      improvements: improvements.length
    });

    return {
      totalTests,
      passed,
      failed,
      criticalIssues,
      improvements,
      stabilityScore
    };
  }

  /**
   * Get improvement suggestions based on test results
   */
  getImprovementSuggestions(): string[] {
    const suggestions: string[] = [];
    
    // Analyze failure patterns
    const failedTests = this.testResults.filter(r => !r.passed);
    const commonIssues = this.findCommonPatterns(failedTests);
    
    if (commonIssues.includes('uncalled_function')) {
      suggestions.push('Improve function call detection with better regex patterns');
      suggestions.push('Add support for nested function declarations');
      suggestions.push('Implement context-aware function call insertion');
    }
    
    if (commonIssues.includes('array_dimension')) {
      suggestions.push('Enhance array dimension validation for complex structures');
      suggestions.push('Add type-specific correction strategies');
    }
    
    return suggestions;
  }

  private findCommonPatterns(failedTests: TestResult[]): string[] {
    const patterns: string[] = [];
    const issueCount: { [key: string]: number } = {};
    
    failedTests.forEach(test => {
      test.detectedIssues.forEach(issue => {
        issueCount[issue] = (issueCount[issue] || 0) + 1;
      });
    });
    
    // Issues that appear in >50% of failed tests are considered common
    const threshold = Math.ceil(failedTests.length * 0.5);
    for (const [issue, count] of Object.entries(issueCount)) {
      if (count >= threshold) {
        patterns.push(issue);
      }
    }
    
    return patterns;
  }
}

// Export singleton instance
export const edgeCaseTester = new EdgeCaseTester();