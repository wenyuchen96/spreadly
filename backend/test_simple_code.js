// Simple Excel.js test code that should work in Script Lab
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Test basic operations
    sheet.getRange("A1").values = [["Test Value"]];
    sheet.getRange("A2").values = [["Another Value"]];
    
    await context.sync();
});