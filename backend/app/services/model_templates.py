"""
Financial Model Templates and Examples for AI Code Generation
These templates ensure consistent, high-quality output
"""

DCF_TEMPLATE = """
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // === ASSUMPTIONS SECTION ===
    sheet.getRange("A1:B1").values = [["DCF VALUATION MODEL", ""]];
    sheet.getRange("A1:B1").format.font.bold = true;
    sheet.getRange("A1:B1").format.font.size = 14;
    
    sheet.getRange("A3:B3").values = [["ASSUMPTIONS", ""]];
    sheet.getRange("A3:B3").format.fill.color = "#4472C4";
    sheet.getRange("A3:B3").format.font.bold = true;
    
    sheet.getRange("A4:B8").values = [
        ["Discount Rate (WACC)", "10%"],
        ["Terminal Growth Rate", "2%"],
        ["Tax Rate", "25%"],
        ["Years of Projection", "5"],
        ["Terminal Value Multiple", ""]
    ];
    sheet.getRange("B4:B8").format.fill.color = "#E7F3FF";
    
    // === CASH FLOW PROJECTIONS ===
    sheet.getRange("D3:I3").values = [["CASH FLOW PROJECTIONS", "", "", "", "", ""]];
    sheet.getRange("D3:I3").format.font.bold = true;
    sheet.getRange("D3:I3").format.fill.color = "#4472C4";
    
    sheet.getRange("D4:I4").values = [["Year", "0", "1", "2", "3", "4", "5"]];
    sheet.getRange("D4:I4").format.font.bold = true;
    
    sheet.getRange("D5:I10").values = [
        ["Revenue", "", "100000", "110000", "121000", "133100", "146410"],
        ["Operating Expenses", "", "-60000", "-66000", "-72600", "-79860", "-87846"],
        ["EBITDA", "", "=F5+F6", "=G5+G6", "=H5+H6", "=I5+I6", "=J5+J6"],
        ["Depreciation", "", "-5000", "-5500", "-6050", "-6655", "-7321"],
        ["EBIT", "", "=F7+F8", "=G7+G8", "=H7+H8", "=I7+I8", "=J7+J8"],
        ["Tax", "", "=F9*$B$6", "=G9*$B$6", "=H9*$B$6", "=I9*$B$6", "=J9*$B$6"]
    ];
    
    // === VALUATION CALCULATIONS ===
    sheet.getRange("D12:I12").values = [["Free Cash Flow", "", "=F9+F8+F10", "=G9+G8+G10", "=H9+H8+H10", "=I9+I8+I10", "=J9+J8+J10"]];
    sheet.getRange("D13:I13").values = [["Discount Factor", "", "=1/POWER(1+$B$4,F4)", "=1/POWER(1+$B$4,G4)", "=1/POWER(1+$B$4,H4)", "=1/POWER(1+$B$4,I4)", "=1/POWER(1+$B$4,J4)"]];
    sheet.getRange("D14:I14").values = [["Present Value", "", "=F12*F13", "=G12*G13", "=H12*H13", "=I12*I13", "=J12*J13"]];
    
    // === RESULTS ===
    sheet.getRange("A12:B16").values = [
        ["VALUATION RESULTS", ""],
        ["Sum of PV Cash Flows", "=SUM(G14:J14)"],
        ["Terminal Value", "=J12*(1+$B$5)/($B$4-$B$5)"],
        ["PV of Terminal Value", "=B14*J13"],
        ["Enterprise Value", "=B13+B15"]
    ];
    sheet.getRange("A12:B16").format.fill.color = "#D4EDDA";
    sheet.getRange("A12:A16").format.font.bold = true;
    
    // === FORMATTING ===
    sheet.getRange("F5:J16").format.numberFormat = "$#,##0";
    sheet.getRange("B4:B5").format.numberFormat = "0%";
    
    await context.sync();
});
"""

NPV_TEMPLATE = """
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // === PROJECT ASSUMPTIONS ===
    sheet.getRange("A1:B1").values = [["NPV ANALYSIS", ""]];
    sheet.getRange("A1:B1").format.font.bold = true;
    sheet.getRange("A1:B1").format.font.size = 14;
    
    sheet.getRange("A3:B3").values = [["INPUT ASSUMPTIONS", ""]];
    sheet.getRange("A3:B3").format.fill.color = "#4472C4";
    sheet.getRange("A3:B3").format.font.bold = true;
    
    sheet.getRange("A4:B7").values = [
        ["Initial Investment", "-100000"],
        ["Discount Rate", "12%"],
        ["Project Life (Years)", "5"],
        ["Annual Cash Flow", "25000"]
    ];
    sheet.getRange("B4:B7").format.fill.color = "#E7F3FF";
    
    // === CASH FLOW TABLE ===
    sheet.getRange("D3:H3").values = [["CASH FLOW ANALYSIS", "", "", "", ""]];
    sheet.getRange("D3:H3").format.font.bold = true;
    sheet.getRange("D3:H3").format.fill.color = "#4472C4";
    
    sheet.getRange("D4:H4").values = [["Year", "Cash Flow", "Discount Factor", "Present Value", "Cumulative NPV"]];
    sheet.getRange("D4:H4").format.font.bold = true;
    
    sheet.getRange("D5:H10").values = [
        ["0", "=$B$4", "1", "=E5*F5", "=G5"],
        ["1", "=$B$7", "=1/POWER(1+$B$5,D6)", "=E6*F6", "=H5+G6"],
        ["2", "=$B$7", "=1/POWER(1+$B$5,D7)", "=E7*F7", "=H6+G7"],
        ["3", "=$B$7", "=1/POWER(1+$B$5,D8)", "=E8*F8", "=H7+G8"],
        ["4", "=$B$7", "=1/POWER(1+$B$5,D9)", "=E9*F9", "=H8+G9"],
        ["5", "=$B$7", "=1/POWER(1+$B$5,D10)", "=E10*F10", "=H9+G10"]
    ];
    
    // === RESULTS & METRICS ===
    sheet.getRange("A10:B15").values = [
        ["PROJECT METRICS", ""],
        ["Net Present Value", "=H10"],
        ["Internal Rate of Return", "=IRR(E5:E10)"],
        ["Payback Period (Years)", "=MATCH(TRUE,H5:H10>0,0)-1"],
        ["Profitability Index", "=1+(B12/-$B$4)"]
    ];
    sheet.getRange("A10:B15").format.fill.color = "#D4EDDA";
    sheet.getRange("A10:A15").format.font.bold = true;
    
    // === SENSITIVITY ANALYSIS ===
    sheet.getRange("J3:N3").values = [["SENSITIVITY ANALYSIS", "", "", "", ""]];
    sheet.getRange("J3:N3").format.font.bold = true;
    sheet.getRange("J3:N3").format.fill.color = "#FFA500";
    
    sheet.getRange("J4:N8").values = [
        ["Discount Rate", "10%", "12%", "14%", "16%"],
        ["NPV @10%", "=NPV(J5,$E$6:$E$10)+$E$5", "", "", ""],
        ["NPV @12%", "", "=$B$12", "", ""],
        ["NPV @14%", "", "", "=NPV(M5,$E$6:$E$10)+$E$5", ""],
        ["NPV @16%", "", "", "", "=NPV(N5,$E$6:$E$10)+$E$5"]
    ];
    
    // === FORMATTING ===
    sheet.getRange("E5:H10").format.numberFormat = "$#,##0";
    sheet.getRange("B4").format.numberFormat = "$#,##0";
    sheet.getRange("B5").format.numberFormat = "0%";
    sheet.getRange("B12:B15").format.numberFormat = "$#,##0";
    sheet.getRange("B13").format.numberFormat = "0%";
    sheet.getRange("K5:N8").format.numberFormat = "$#,##0";
    
    await context.sync();
});
"""

def get_template_for_model(model_type: str) -> str:
    """Return appropriate template based on model type"""
    model_type_lower = model_type.lower()
    
    if 'dcf' in model_type_lower or 'discounted cash flow' in model_type_lower:
        return DCF_TEMPLATE
    elif 'npv' in model_type_lower:
        return NPV_TEMPLATE
    else:
        return NPV_TEMPLATE  # Default to NPV as it's simpler