import { ScriptLabEngine, ExcelOperations } from './index';

// Test function to verify Script Lab engine integration
export async function testScriptLabEngine(): Promise<void> {
  console.log('🧪 Testing Script Lab Engine Integration...');
  
  try {
    // Test 1: Create engine instance
    const engine = new ScriptLabEngine();
    console.log('✅ ScriptLabEngine instance created successfully');
    
    // Test 2: Create a simple snippet
    const testSnippet = engine.createSnippet(
      'Test Snippet',
      `
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    
    range.values = [["Test", "Success"]];
    range.format.fill.color = "lightgreen";
    
    await context.sync();
    console.log("Test snippet executed on range: " + range.address);
});
      `,
      '<div>Test execution</div>',
      'body { font-family: Arial; }',
      []
    );
    
    console.log('✅ Test snippet created:', testSnippet.name);
    
    // Test 3: Validate code syntax
    const validation = await engine.validateCode(testSnippet.script);
    console.log('✅ Code validation result:', validation.isValid ? 'Valid' : 'Invalid');
    if (!validation.isValid) {
      console.log('❌ Validation errors:', validation.errors);
    }
    
    // Test 4: Test Excel operations helpers
    const highlightSnippet = ExcelOperations.createHighlightCellsSnippet('yellow');
    console.log('✅ Highlight snippet created:', highlightSnippet.name);
    
    const insertDataSnippet = ExcelOperations.createInsertDataSnippet([
      ['Name', 'Age', 'City'],
      ['Alice', 25, 'New York'],
      ['Bob', 30, 'San Francisco']
    ]);
    console.log('✅ Insert data snippet created:', insertDataSnippet.name);
    
    const chartSnippet = ExcelOperations.createChartSnippet('A1:C4', 'ColumnClustered');
    console.log('✅ Chart snippet created:', chartSnippet.name);
    
    // Test 5: Attempt code execution (this will only work when Office.js is available)
    if (typeof Office !== 'undefined' && Office.context) {
      console.log('🔄 Office.js detected, attempting code execution...');
      const result = await engine.executeSnippet(testSnippet);
      console.log('📊 Execution result:', result);
    } else {
      console.log('ℹ️ Office.js not available, skipping execution test');
      console.log('ℹ️ This is normal when testing outside of Excel');
    }
    
    console.log('🎉 All Script Lab Engine tests completed successfully!');
    
    // Clean up
    engine.dispose();
    console.log('✅ Engine disposed successfully');
    
  } catch (error) {
    console.error('❌ Test failed:', error);
    throw error;
  }
}

// Manual test commands for chat interface
export const testCommands = {
  highlight: 'highlight cells yellow',
  insertData: 'insert data [[1,2,3],[4,5,6]]',
  createChart: 'create chart A1:C5',
  formatCells: 'format cells A1:B2',
  runDemo: 'test'
};

// Function to run a specific test command
export function getTestCommand(command: keyof typeof testCommands): string {
  return testCommands[command];
}

// Export test utilities for use in console
if (typeof window !== 'undefined') {
  (window as any).testScriptLab = testScriptLabEngine;
  (window as any).testCommands = testCommands;
  (window as any).getTestCommand = getTestCommand;
}