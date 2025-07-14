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

/**
 * Comprehensive workbook data structure
 */
export interface ComprehensiveWorkbookData {
  metadata: {
    totalSheets: number;
    activeSheetName: string;
    lastModified: string;
  };
  sheets: SheetData[];
  tables: TableData[];
  namedRanges: NamedRangeData[];
  summary: {
    totalCells: number;
    totalUsedCells: number;
    hasFormulas: boolean;
    hasCharts: boolean;
  };
}

export interface SheetData {
  name: string;
  id: string;
  isActive: boolean;
  usedRange: {
    address: string;
    rowCount: number;
    columnCount: number;
  } | null;
  data: any[][];
  formulas: string[][];
  dataTypes: string[];
  hasHeaders: boolean;
  tableCount: number;
  chartCount: number;
}

export interface TableData {
  name: string;
  sheetName: string;
  range: string;
  headers: string[];
  rowCount: number;
  columnCount: number;
}

export interface NamedRangeData {
  name: string;
  formula: string;
  value: any;
  scope: string;
}

/**
 * Ultra-simple data reading for debugging
 */
export async function getDirectActiveSheetData(): Promise<{data: any[][], info: string}> {
  console.log('🔍 Direct active sheet data reading...');
  
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load(['name']);
    
    // Try A1:J20 range
    const range = sheet.getRange('A1:J20');
    range.load(['values', 'address']);
    await context.sync();
    
    const allData = range.values as any[][];
    const nonEmptyRows = allData.filter(row => 
      row.some(cell => cell !== null && cell !== '' && cell !== undefined)
    );
    
    const info = `Sheet: ${sheet.name}, Range: ${range.address}, Total rows: ${allData.length}, Non-empty rows: ${nonEmptyRows.length}`;
    console.log('🔍 Direct data result:', info, nonEmptyRows.slice(0, 3));
    
    return { data: nonEmptyRows, info };
  });
}

/**
 * Simple, reliable data extraction for active sheet
 */
export async function getActiveSheetDataReliably(): Promise<any[][]> {
  console.log('🔍 Getting active sheet data with reliable method...');
  
  return await Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    
    try {
      // Method 1: Try used range
      const usedRange = activeSheet.getUsedRange();
      usedRange.load(["address", "values", "rowCount", "columnCount"]);
      await context.sync();
      
      if (usedRange.address && usedRange.values) {
        const data = usedRange.values as any[][];
        console.log('✅ Used range method worked:', { address: usedRange.address, data: data.slice(0, 3) });
        return data;
      }
    } catch (error) {
      console.warn('⚠️ Used range method failed:', error);
    }
    
    try {
      // Method 2: Try specific range A1:Z100
      const specificRange = activeSheet.getRange("A1:Z100");
      specificRange.load(["values"]);
      await context.sync();
      
      const data = specificRange.values as any[][];
      // Filter out empty rows
      const nonEmptyData = data.filter(row => row.some(cell => cell !== null && cell !== "" && cell !== undefined));
      
      if (nonEmptyData.length > 0) {
        console.log('✅ Specific range method worked:', { rows: nonEmptyData.length, data: nonEmptyData.slice(0, 3) });
        return nonEmptyData;
      }
    } catch (error) {
      console.warn('⚠️ Specific range method failed:', error);
    }
    
    console.log('❌ Both methods failed, returning empty array');
    return [];
  });
}

/**
 * Get comprehensive workbook data including all sheets, tables, and ranges
 */
export async function getComprehensiveWorkbookData(): Promise<ComprehensiveWorkbookData> {
  console.log('🔍 Starting comprehensive workbook data extraction...');
  
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const activeWorksheet = workbook.worksheets.getActiveWorksheet();
    
    // Load workbook-level data
    workbook.worksheets.load(["name", "id"]);
    workbook.tables.load(["name", "worksheet"]);
    workbook.names.load(["name", "formula", "value", "scope"]);
    activeWorksheet.load(["name", "id"]);
    
    await context.sync();
    
    console.log(`🔍 Found ${workbook.worksheets.items.length} worksheets`);
    
    const sheets: SheetData[] = [];
    const tables: TableData[] = [];
    const namedRanges: NamedRangeData[] = [];
    
    let totalCells = 0;
    let totalUsedCells = 0;
    let hasFormulas = false;
    let hasCharts = false;
    
    // Process each worksheet
    for (const worksheet of workbook.worksheets.items) {
      console.log(`🔍 Processing sheet: ${worksheet.name}`);
      
      try {
        // Get used range for this worksheet
        const usedRange = worksheet.getUsedRange();
        worksheet.tables.load(["name", "range"]);
        worksheet.charts.load(["name"]);
        
        // Load range data if it exists
        let rangeData: any[][] = [];
        let rangeFormulas: string[][] = [];
        let usedRangeInfo = null;
        
        try {
          usedRange.load(["address", "values", "formulas", "rowCount", "columnCount"]);
          await context.sync();
          
          console.log(`🔍 After sync for ${worksheet.name} - usedRange loaded:`, {
            address: usedRange.address,
            hasValues: usedRange.values !== undefined,
            valuesType: typeof usedRange.values,
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount
          });
          
          if (usedRange.address) {
            rangeData = usedRange.values as any[][];
            rangeFormulas = usedRange.formulas as string[][];
            
            // If usedRange method failed to get data, try the direct approach that worked
            if (rangeData.length === 0 && usedRange.rowCount > 0) {
              console.log(`🔍 Used range returned empty data for ${worksheet.name}, trying direct method...`);
              try {
                // Use the same successful approach as getDirectActiveSheetData
                const directRange = worksheet.getRange('A1:J20');
                directRange.load(['values']);
                await context.sync();
                
                const allData = directRange.values as any[][];
                const nonEmptyRows = allData.filter(row => 
                  row.some(cell => cell !== null && cell !== '' && cell !== undefined)
                );
                
                if (nonEmptyRows.length > 0) {
                  rangeData = nonEmptyRows;
                  console.log(`✅ Direct method successful for ${worksheet.name}:`, { rows: nonEmptyRows.length });
                }
              } catch (directError) {
                console.warn(`❌ Direct method failed for ${worksheet.name}:`, directError);
              }
            }
            usedRangeInfo = {
              address: usedRange.address,
              rowCount: rangeData.length > 0 ? rangeData.length : usedRange.rowCount,  // Use actual data rows
              columnCount: rangeData.length > 0 && rangeData[0] ? rangeData[0].length : usedRange.columnCount  // Use actual data cols
            };
            
            // Debug: Log actual data extracted
            console.log(`🔍 Data extracted from ${worksheet.name}:`, {
              address: usedRange.address,
              dimensions: `${usedRange.rowCount}x${usedRange.columnCount}`,
              sampleData: rangeData.slice(0, 3),  // First 3 rows
              dataLength: rangeData.length,
              firstRowLength: rangeData.length > 0 ? rangeData[0].length : 0,
              allDataFlat: rangeData.flat()
            });
            
            // If rangeData is empty but we have a used range, try alternative loading
            if (rangeData.length === 0 && usedRange.rowCount > 0) {
              console.log(`⚠️ Empty data but non-zero dimensions for ${worksheet.name}, trying alternative loading...`);
              
              try {
                // Try loading the range again with different approach
                const alternativeRange = worksheet.getRange(usedRange.address);
                alternativeRange.load(["values", "formulas"]);
                await context.sync();
                
                rangeData = alternativeRange.values as any[][];
                rangeFormulas = alternativeRange.formulas as string[][];
                
                console.log(`🔍 Alternative loading result for ${worksheet.name}:`, {
                  newDataLength: rangeData.length,
                  newSampleData: rangeData.slice(0, 2)
                });
              } catch (altError) {
                console.warn(`❌ Alternative loading failed for ${worksheet.name}:`, altError);
                
                // Last resort: Try reading a smaller range directly
                console.log(`🔍 Trying direct range reading for ${worksheet.name}...`);
                try {
                  // Try reading just the first few rows/columns directly
                  const sampleRows = Math.min(5, usedRange.rowCount);
                  const sampleCols = Math.min(5, usedRange.columnCount);
                  const endCol = String.fromCharCode(64 + sampleCols); // A=65, so A=65-1+1
                  const sampleAddress = `A1:${endCol}${sampleRows}`;
                  
                  console.log(`🔍 Trying to read sample range: ${sampleAddress}`);
                  const sampleRange = worksheet.getRange(sampleAddress);
                  sampleRange.load(["values"]);
                  await context.sync();
                  
                  if (sampleRange.values && (sampleRange.values as any[][]).length > 0) {
                    rangeData = sampleRange.values as any[][];
                    console.log(`✅ Sample range reading successful for ${worksheet.name}:`, rangeData);
                  }
                } catch (sampleError) {
                  console.warn(`❌ Sample range reading failed for ${worksheet.name}:`, sampleError);
                }
              }
            }
            
            totalCells += usedRange.rowCount * usedRange.columnCount;
            
            // More sophisticated counting of used cells
            const nonEmptyCells = rangeData.flat().filter(cell => {
              return cell !== null && cell !== undefined && cell !== '' && cell !== 0;
            });
            totalUsedCells += nonEmptyCells.length;
            
            // Debug: Show what we consider "used" vs "empty"
            console.log(`🔍 Cell analysis for ${worksheet.name}:`, {
              totalCells: usedRange.rowCount * usedRange.columnCount,
              nonEmptyCells: nonEmptyCells.length,
              sampleNonEmptyValues: nonEmptyCells.slice(0, 5)
            });
            
            // Check for formulas
            if (rangeFormulas.some(row => row.some(cell => cell.startsWith('=')))) {
              hasFormulas = true;
            }
          }
        } catch (error) {
          console.log(`⚠️ No used range found for sheet: ${worksheet.name}`);
          rangeData = [];
          rangeFormulas = [];
        }
        
        await context.sync();
        
        // Check for charts
        if (worksheet.charts.items.length > 0) {
          hasCharts = true;
        }
        
        // Additional check: Verify if the data is actually meaningful
        const hasActualData = rangeData.some(row => 
          row.some(cell => cell !== null && cell !== undefined && cell !== '' && cell !== 0)
        );
        
        console.log(`🔍 Final check for ${worksheet.name}:`, {
          hasUsedRange: !!usedRangeInfo,
          hasActualData,
          dataRowCount: rangeData.length,
          usedRangeRowCount: usedRangeInfo?.rowCount || 0
        });

        const sheetData: SheetData = {
          name: worksheet.name,
          id: worksheet.id,
          isActive: worksheet.name === activeWorksheet.name,
          usedRange: usedRangeInfo,
          data: rangeData,
          formulas: rangeFormulas,
          dataTypes: analyzeDataTypes(rangeData),
          hasHeaders: detectHeaders(rangeData),
          tableCount: worksheet.tables.items.length,
          chartCount: worksheet.charts.items.length
        };
        
        sheets.push(sheetData);
        
        // Process tables in this worksheet
        for (const table of worksheet.tables.items) {
          const tableRange = table.getRange();
          const tableHeaders = table.getHeaderRowRange();
          
          tableRange.load(["address", "rowCount", "columnCount"]);
          tableHeaders.load(["values"]);
          
          await context.sync();
          
          const tableData: TableData = {
            name: table.name,
            sheetName: worksheet.name,
            range: tableRange.address,
            headers: (tableHeaders.values as any[][])[0] || [],
            rowCount: tableRange.rowCount,
            columnCount: tableRange.columnCount
          };
          
          tables.push(tableData);
        }
        
      } catch (error) {
        console.error(`❌ Error processing sheet ${worksheet.name}:`, error);
        
        // Add minimal sheet data for failed sheets
        sheets.push({
          name: worksheet.name,
          id: worksheet.id,
          isActive: worksheet.name === activeWorksheet.name,
          usedRange: null,
          data: [],
          formulas: [],
          dataTypes: [],
          hasHeaders: false,
          tableCount: 0,
          chartCount: 0
        });
      }
    }
    
    // Process named ranges
    for (const namedRange of workbook.names.items) {
      try {
        namedRanges.push({
          name: namedRange.name,
          formula: namedRange.formula,
          value: namedRange.value,
          scope: namedRange.scope
        });
      } catch (error) {
        console.warn(`⚠️ Could not process named range: ${namedRange.name}`);
      }
    }
    
    const comprehensiveData: ComprehensiveWorkbookData = {
      metadata: {
        totalSheets: workbook.worksheets.items.length,
        activeSheetName: activeWorksheet.name,
        lastModified: new Date().toISOString()
      },
      sheets,
      tables,
      namedRanges,
      summary: {
        totalCells,
        totalUsedCells,
        hasFormulas,
        hasCharts
      }
    };
    
    console.log('✅ Comprehensive workbook data extraction completed');
    console.log(`📊 Summary: ${sheets.length} sheets, ${tables.length} tables, ${namedRanges.length} named ranges`);
    console.log(`📊 Data: ${totalUsedCells}/${totalCells} used cells, formulas: ${hasFormulas}, charts: ${hasCharts}`);
    
    return comprehensiveData;
  });
}

/**
 * Get lightweight workbook context (for performance-sensitive operations)
 */
export async function getLightweightWorkbookContext(): Promise<Partial<ComprehensiveWorkbookData>> {
  console.log('🔍 Getting lightweight workbook context...');
  
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const activeWorksheet = workbook.worksheets.getActiveWorksheet();
    
    workbook.worksheets.load(["name", "id"]);
    workbook.tables.load(["name"]);
    workbook.names.load(["name"]);
    activeWorksheet.load(["name", "id"]);
    
    await context.sync();
    
    // Get basic info for each sheet without full data
    const sheets: Partial<SheetData>[] = [];
    
    for (const worksheet of workbook.worksheets.items) {
      try {
        const usedRange = worksheet.getUsedRange();
        worksheet.tables.load(["name"]);
        worksheet.charts.load(["name"]);
        
        let usedRangeInfo = null;
        
        try {
          usedRange.load(["address", "rowCount", "columnCount"]);
          await context.sync();
          
          if (usedRange.address) {
            usedRangeInfo = {
              address: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount
            };
          }
        } catch (error) {
          // No used range
        }
        
        await context.sync();
        
        sheets.push({
          name: worksheet.name,
          id: worksheet.id,
          isActive: worksheet.name === activeWorksheet.name,
          usedRange: usedRangeInfo,
          tableCount: worksheet.tables.items.length,
          chartCount: worksheet.charts.items.length
        });
        
      } catch (error) {
        console.warn(`⚠️ Error getting info for sheet ${worksheet.name}:`, error);
      }
    }
    
    return {
      metadata: {
        totalSheets: workbook.worksheets.items.length,
        activeSheetName: activeWorksheet.name,
        lastModified: new Date().toISOString()
      },
      sheets,
      tables: workbook.tables.items.map(table => ({ name: table.name })),
      namedRanges: workbook.names.items.map(range => ({ name: range.name }))
    };
  });
}