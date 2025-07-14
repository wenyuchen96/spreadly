// This is the exact output from our improved AI system
// Test this in Script Lab to verify it works

await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    sheet.getRange("A3").values = [["Company Valuation"]];
    sheet.getRange("A5").values = [["Revenue Projections"]];
    sheet.getRange("A6").values = [["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]];
    sheet.getRange("A8").values = [["Assumptions"]];
    sheet.getRange("A9").values = [["Revenue Growth Rate"]];
    
    await context.sync();
});