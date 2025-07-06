// Main exports for the ScriptLab engine
export { ScriptLabEngine } from './ScriptLabEngine';
export { TypeScriptCompiler } from './TypeScriptCompiler';
export { OfficeJsHost } from './officeJsHost';
export * from './interfaces';

// Common Excel operations as helper functions
export const ExcelOperations = {
  // Create a snippet to highlight selected cells
  createHighlightCellsSnippet: (color: string = 'yellow') => {
    const code = `
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = "${color}";
    await context.sync();
    console.log("Selected cells highlighted with ${color}");
});
`;
    return ScriptLabEngine.createExcelSnippet(code, `Highlight selected cells with ${color} color`);
  },

  // Create a snippet to insert data into cells
  createInsertDataSnippet: (data: any[][], startCell: string = 'A1') => {
    const code = `
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = worksheet.getRange("${startCell}");
    const resizedRange = range.getResizedRange(${data.length - 1}, ${data[0]?.length - 1 || 0});
    resizedRange.values = ${JSON.stringify(data)};
    await context.sync();
    console.log("Data inserted starting at ${startCell}");
});
`;
    return ScriptLabEngine.createExcelSnippet(code, `Insert data starting at ${startCell}`);
  },

  // Create a snippet to create a chart
  createChartSnippet: (dataRange: string, chartType: string = 'ColumnClustered') => {
    const code = `
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = worksheet.getRange("${dataRange}");
    const chart = worksheet.charts.add(Excel.ChartType.${chartType}, range);
    chart.title.text = "Generated Chart";
    await context.sync();
    console.log("Chart created from range ${dataRange}");
});
`;
    return ScriptLabEngine.createExcelSnippet(code, `Create ${chartType} chart from range ${dataRange}`);
  },

  // Create a snippet to format cells
  createFormatCellsSnippet: (range: string, format: any) => {
    const formatCode = Object.entries(format).map(([key, value]) => {
      if (key === 'fill') {
        return `range.format.fill.color = "${value}";`;
      } else if (key === 'font') {
        return Object.entries(value as any).map(([fontKey, fontValue]) => 
          `range.format.font.${fontKey} = ${typeof fontValue === 'string' ? `"${fontValue}"` : fontValue};`
        ).join('\n    ');
      }
      return `range.format.${key} = ${typeof value === 'string' ? `"${value}"` : value};`;
    }).join('\n    ');

    const code = `
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = worksheet.getRange("${range}");
    ${formatCode}
    await context.sync();
    console.log("Range ${range} formatted");
});
`;
    return ScriptLabEngine.createExcelSnippet(code, `Format range ${range}`);
  }
};

// Import the main engine class for static methods
import { ScriptLabEngine } from './ScriptLabEngine';