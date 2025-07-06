/**
 * Excel data extraction utilities
 */

export interface ExcelDataSummary {
  range: string;
  rowCount: number;
  columnCount: number;
  data: any[][];
  hasHeaders: boolean;
  dataTypes: string[];
}

/**
 * Get data from the currently selected range
 */
export async function getSelectedRangeData(): Promise<ExcelDataSummary> {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "values", "rowCount", "columnCount"]);
    await context.sync();
    
    const data = range.values as any[][];
    
    return {
      range: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      data: data,
      hasHeaders: detectHeaders(data),
      dataTypes: analyzeDataTypes(data)
    };
  });
}

/**
 * Get data from the entire used range of the active worksheet
 */
export async function getWorksheetData(): Promise<ExcelDataSummary> {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = worksheet.getUsedRange();
    range.load(["address", "values", "rowCount", "columnCount"]);
    await context.sync();
    
    if (!range) {
      return {
        range: "A1",
        rowCount: 0,
        columnCount: 0,
        data: [],
        hasHeaders: false,
        dataTypes: []
      };
    }
    
    const data = range.values as any[][];
    
    return {
      range: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      data: data,
      hasHeaders: detectHeaders(data),
      dataTypes: analyzeDataTypes(data)
    };
  });
}

/**
 * Get data from a specific range address
 */
export async function getRangeData(rangeAddress: string): Promise<ExcelDataSummary> {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = worksheet.getRange(rangeAddress);
    range.load(["address", "values", "rowCount", "columnCount"]);
    await context.sync();
    
    const data = range.values as any[][];
    
    return {
      range: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      data: data,
      hasHeaders: detectHeaders(data),
      dataTypes: analyzeDataTypes(data)
    };
  });
}

/**
 * Detect if the first row contains headers
 */
function detectHeaders(data: any[][]): boolean {
  if (!data || data.length === 0) return false;
  
  const firstRow = data[0];
  if (!firstRow) return false;
  
  // Check if first row contains mostly strings and subsequent rows contain numbers
  const firstRowStrings = firstRow.filter(cell => typeof cell === 'string').length;
  const firstRowNumbers = firstRow.filter(cell => typeof cell === 'number').length;
  
  if (data.length > 1) {
    const secondRow = data[1];
    const secondRowNumbers = secondRow.filter(cell => typeof cell === 'number').length;
    const secondRowStrings = secondRow.filter(cell => typeof cell === 'string').length;
    
    // If first row is mostly strings and second row is mostly numbers, likely headers
    return firstRowStrings > firstRowNumbers && secondRowNumbers > secondRowStrings;
  }
  
  // If only one row, assume headers if mostly strings
  return firstRowStrings > firstRowNumbers;
}

/**
 * Analyze data types for each column
 */
function analyzeDataTypes(data: any[][]): string[] {
  if (!data || data.length === 0) return [];
  
  const columnCount = Math.max(...data.map(row => row.length));
  const dataTypes: string[] = [];
  
  for (let col = 0; col < columnCount; col++) {
    const columnValues = data.map(row => row[col]).filter(val => val !== null && val !== undefined && val !== '');
    
    if (columnValues.length === 0) {
      dataTypes.push('empty');
      continue;
    }
    
    const numberCount = columnValues.filter(val => typeof val === 'number' || !isNaN(Number(val))).length;
    const stringCount = columnValues.filter(val => typeof val === 'string' && isNaN(Number(val))).length;
    const dateCount = columnValues.filter(val => isDateString(val)).length;
    
    if (dateCount > columnValues.length * 0.5) {
      dataTypes.push('date');
    } else if (numberCount > columnValues.length * 0.7) {
      dataTypes.push('number');
    } else if (stringCount > columnValues.length * 0.7) {
      dataTypes.push('text');
    } else {
      dataTypes.push('mixed');
    }
  }
  
  return dataTypes;
}

/**
 * Check if a value looks like a date string
 */
function isDateString(value: any): boolean {
  if (typeof value !== 'string') return false;
  
  const date = new Date(value);
  return !isNaN(date.getTime()) && value.match(/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/);
}

/**
 * Get worksheet information
 */
export async function getWorksheetInfo() {
  return await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load(["name", "id"]);
    
    const workbook = context.workbook;
    workbook.worksheets.load(["name"]);
    
    await context.sync();
    
    return {
      activeSheet: {
        name: worksheet.name,
        id: worksheet.id
      },
      allSheets: workbook.worksheets.items.map(sheet => sheet.name)
    };
  });
}