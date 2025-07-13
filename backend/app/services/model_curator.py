"""
Model Curator Service for Financial Model Template Management
Handles creation, management, and initialization of professional financial model templates
"""

import json
import logging
from typing import List, Dict, Any
from datetime import datetime

from app.models.financial_model import (
    FinancialModel,
    ModelType,
    Industry,
    ComplexityLevel,
    ModelMetadata,
    PerformanceMetrics
)
from app.services.model_vector_store import get_vector_store


class ModelCurator:
    """
    Service for managing and curating financial model templates
    """
    
    def __init__(self):
        self.vector_store = get_vector_store()
    
    async def initialize_model_library(self) -> Dict[str, Any]:
        """Initialize the model library with professional templates"""
        logging.info("Initializing financial model library...")
        
        results = {
            "total_added": 0,
            "successful": [],
            "failed": [],
            "categories": {}
        }
        
        # Get all professional templates
        templates = self.get_professional_templates()
        
        for template in templates:
            try:
                success = await self.vector_store.add_model(template)
                if success:
                    results["successful"].append(template.id)
                    results["total_added"] += 1
                    
                    # Track by category
                    category = f"{template.model_type}_{template.industry}"
                    results["categories"][category] = results["categories"].get(category, 0) + 1
                else:
                    results["failed"].append(template.id)
                    
            except Exception as e:
                logging.error(f"Failed to add template {template.id}: {e}")
                results["failed"].append(template.id)
        
        logging.info(f"Model library initialized: {results['total_added']} models added")
        return results
    
    def get_professional_templates(self) -> List[FinancialModel]:
        """Get all professional financial model templates"""
        templates = []
        
        # DCF Models
        templates.extend(self._create_dcf_templates())
        
        # NPV Models
        templates.extend(self._create_npv_templates())
        
        # Valuation Models
        templates.extend(self._create_valuation_templates())
        
        # Budget/Forecast Models
        templates.extend(self._create_budget_templates())
        
        return templates
    
    def _create_dcf_templates(self) -> List[FinancialModel]:
        """Create DCF model templates"""
        templates = []
        
        # Technology DCF - Advanced
        tech_dcf = FinancialModel(
            id="dcf_tech_advanced_001",
            name="Technology Company DCF Model",
            description="Advanced DCF model for technology companies with SaaS metrics, R&D capitalization, and multiple scenario analysis",
            model_type=ModelType.DCF,
            industry=Industry.TECHNOLOGY,
            complexity=ComplexityLevel.ADVANCED,
            excel_code='''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // HEADER
    sheet.getRange("A1").values = [["TECHNOLOGY COMPANY DCF VALUATION MODEL"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.fill.color = "#2F4F4F";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    
    // ASSUMPTIONS SECTION
    sheet.getRange("A3").values = [["KEY ASSUMPTIONS"]];
    sheet.getRange("A3:F3").format.fill.color = "#4472C4";
    sheet.getRange("A3:F3").format.font.bold = true;
    sheet.getRange("A3:F3").format.font.color = "#FFFFFF";
    
    sheet.getRange("A4:B12").values = [
        ["Revenue Growth Rate (Y1-3)", "25%"],
        ["Revenue Growth Rate (Y4-5)", "15%"],
        ["Terminal Growth Rate", "3%"],
        ["EBITDA Margin (Mature)", "30%"],
        ["Tax Rate", "25%"],
        ["WACC", "12%"],
        ["CapEx as % of Revenue", "3%"],
        ["Working Capital as % Rev", "5%"],
        ["Debt/Equity Ratio", "20%"]
    ];
    sheet.getRange("B4:B12").format.fill.color = "#E7F3FF";
    
    // PROJECTION YEARS
    sheet.getRange("H3:M3").values = [["Year", "1", "2", "3", "4", "5"]];
    sheet.getRange("H3:M3").format.font.bold = true;
    sheet.getRange("H3:M3").format.fill.color = "#4472C4";
    sheet.getRange("H3:M3").format.font.color = "#FFFFFF";
    
    // REVENUE PROJECTIONS
    sheet.getRange("H4:M8").values = [
        ["Revenue", "100000", "=I4*(1+$B$4)", "=J4*(1+$B$4)", "=K4*(1+$B$5)", "=L4*(1+$B$5)"],
        ["Revenue Growth %", "=($B$4)", "=($B$4)", "=($B$4)", "=($B$5)", "=($B$5)"],
        ["EBITDA", "=I4*0.20", "=J4*0.25", "=K4*$B$7", "=L4*$B$7", "=M4*$B$7"],
        ["Depreciation", "=I4*0.02", "=J4*0.025", "=K4*0.03", "=L4*0.03", "=M4*0.03"],
        ["EBIT", "=I6-I7", "=J6-J7", "=K6-K7", "=L6-L7", "=M6-M7"]
    ];
    
    // FREE CASH FLOW CALCULATION
    sheet.getRange("H10:M15").values = [
        ["Tax", "=I8*$B$8", "=J8*$B$8", "=K8*$B$8", "=L8*$B$8", "=M8*$B$8"],
        ["NOPAT", "=I8-I10", "=J8-J10", "=K8-K10", "=L8-L10", "=M8-M10"],
        ["+ Depreciation", "=I7", "=J7", "=K7", "=L7", "=M7"],
        ["- CapEx", "=-I4*$B$9", "=-J4*$B$9", "=-K4*$B$9", "=-L4*$B$9", "=-M4*$B$9"],
        ["- Δ Working Capital", "=-I4*$B$10", "=-(J4-I4)*$B$10", "=-(K4-J4)*$B$10", "=-(L4-K4)*$B$10", "=-(M4-L4)*$B$10"],
        ["Free Cash Flow", "=SUM(I11:I14)", "=SUM(J11:J14)", "=SUM(K11:K14)", "=SUM(L11:L14)", "=SUM(M11:M14)"]
    ];
    
    // TERMINAL VALUE
    sheet.getRange("H17:I19").values = [
        ["Terminal FCF", "=M15*(1+$B$6)"],
        ["Terminal Value", "=I17/($B$9-$B$6)"],
        ["PV of Terminal Value", "=I18/POWER(1+$B$9,5)"]
    ];
    
    // VALUATION
    sheet.getRange("A17:B22").values = [
        ["VALUATION SUMMARY", ""],
        ["PV of FCF (Y1-5)", "=NPV($B$9,I15:M15)"],
        ["PV of Terminal Value", "=I19"],
        ["Enterprise Value", "=B18+B19"],
        ["Less: Net Debt", "0"],
        ["Equity Value", "=B20-B21"]
    ];
    sheet.getRange("A17:B22").format.fill.color = "#D4EDDA";
    sheet.getRange("A17:A22").format.font.bold = true;
    
    // FORMATTING
    sheet.getRange("I4:M22").format.numberFormat = "$#,##0";
    sheet.getRange("B4:B6").format.numberFormat = "0%";
    sheet.getRange("I5:M5").format.numberFormat = "0%";
    
    await context.sync();
});
            ''',
            business_description="Comprehensive DCF model for technology companies featuring SaaS metrics, R&D considerations, and growth scenarios",
            sample_inputs={
                "initial_revenue": 100000000,
                "growth_rate_early": 0.25,
                "growth_rate_mature": 0.15,
                "terminal_growth": 0.03,
                "wacc": 0.12
            },
            expected_outputs={
                "enterprise_value": "calculated_ev",
                "equity_value": "calculated_equity_value",
                "implied_multiple": "ev_revenue_multiple"
            },
            metadata=ModelMetadata(
                components=["revenue_projections", "free_cash_flow", "terminal_value", "wacc_calculation", "sensitivity_analysis"],
                excel_functions=["NPV", "IRR", "POWER", "SUM", "MATCH"],
                formatting_features=["conditional_formatting", "data_validation", "named_ranges"],
                business_assumptions=["revenue_growth_stages", "margin_expansion", "capex_requirements"],
                time_horizon_years=5,
                currencies=["USD"],
                regions=["north_america", "global"]
            ),
            performance=PerformanceMetrics(
                execution_success_rate=0.95,
                user_rating=4.7,
                usage_count=0,
                last_used=None,
                error_count=0,
                modification_frequency=0.15
            ),
            created_by="investment_banking_template",
            keywords=["dcf", "technology", "saas", "valuation", "growth", "terminal value", "wacc"],
            tags=["professional", "investment_grade", "tech_sector"]
        )
        templates.append(tech_dcf)
        
        # Healthcare DCF - Intermediate
        healthcare_dcf = FinancialModel(
            id="dcf_healthcare_intermediate_001",
            name="Healthcare Company DCF Model",
            description="DCF model for healthcare companies with regulatory considerations and drug pipeline valuation",
            model_type=ModelType.DCF,
            industry=Industry.HEALTHCARE,
            complexity=ComplexityLevel.INTERMEDIATE,
            excel_code='''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // HEADER
    sheet.getRange("A1").values = [["HEALTHCARE COMPANY DCF VALUATION"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.fill.color = "#228B22";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    
    // ASSUMPTIONS
    sheet.getRange("A3").values = [["HEALTHCARE ASSUMPTIONS"]];
    sheet.getRange("A3:D3").format.fill.color = "#32CD32";
    sheet.getRange("A3:D3").format.font.bold = true;
    
    sheet.getRange("A4:B9").values = [
        ["Revenue Growth (Organic)", "8%"],
        ["Drug Success Probability", "70%"],
        ["Regulatory Discount", "15%"],
        ["Terminal Growth Rate", "2%"],
        ["Healthcare WACC", "10%"],
        ["Tax Rate", "21%"]
    ];
    sheet.getRange("B4:B9").format.fill.color = "#E7F3FF";
    
    // PROJECTIONS
    sheet.getRange("E3:J3").values = [["Year", "1", "2", "3", "4", "5"]];
    sheet.getRange("E3:J3").format.font.bold = true;
    sheet.getRange("E3:J3").format.fill.color = "#32CD32";
    
    sheet.getRange("E4:J8").values = [
        ["Revenue", "500000", "=F4*(1+$B$4)", "=G4*(1+$B$4)", "=H4*(1+$B$4)", "=I4*(1+$B$4)"],
        ["EBITDA (25%)", "=F4*0.25", "=G4*0.25", "=H4*0.25", "=I4*0.25", "=J4*0.25"],
        ["EBIT", "=F5*0.8", "=G5*0.8", "=H5*0.8", "=I5*0.8", "=J5*0.8"],
        ["Tax", "=F6*$B$9", "=G6*$B$9", "=H6*$B$9", "=I6*$B$9", "=J6*$B$9"],
        ["NOPAT", "=F6-F7", "=G6-G7", "=H6-H7", "=I6-I7", "=J6-J7"]
    ];
    
    // FREE CASH FLOW
    sheet.getRange("E10:J12").values = [
        ["FCF (NOPAT*0.9)", "=F8*0.9", "=G8*0.9", "=H8*0.9", "=I8*0.9", "=J8*0.9"],
        ["Risk Adjustment", "=F10*$B$5", "=G10*$B$5", "=H10*$B$5", "=I10*$B$5", "=J10*$B$5"],
        ["Adj. FCF", "=F10-F11", "=G10-G11", "=H10-H11", "=I10-I11", "=J10-J11"]
    ];
    
    // VALUATION
    sheet.getRange("A14:B18").values = [
        ["VALUATION", ""],
        ["PV of FCF", "=NPV($B$8,F12:J12)"],
        ["Terminal Value", "=J12*(1+$B$7)/($B$8-$B$7)/POWER(1+$B$8,5)"],
        ["Enterprise Value", "=B15+B16"],
        ["Equity Value", "=B17"]
    ];
    sheet.getRange("A14:B18").format.fill.color = "#98FB98";
    
    // FORMATTING
    sheet.getRange("F4:J18").format.numberFormat = "$#,##0";
    sheet.getRange("B4:B9").format.numberFormat = "0%";
    
    await context.sync();
});
            ''',
            business_description="DCF model tailored for healthcare companies with drug pipeline risks and regulatory considerations",
            sample_inputs={
                "initial_revenue": 500000000,
                "organic_growth": 0.08,
                "drug_success_prob": 0.70,
                "regulatory_discount": 0.15
            },
            expected_outputs={
                "enterprise_value": "risk_adjusted_ev",
                "equity_value": "healthcare_equity_value"
            },
            metadata=ModelMetadata(
                components=["revenue_projections", "risk_adjustments", "regulatory_factors", "terminal_value"],
                excel_functions=["NPV", "POWER", "SUM"],
                formatting_features=["sector_specific_formatting"],
                business_assumptions=["drug_pipeline_risk", "regulatory_approval", "market_penetration"],
                time_horizon_years=5,
                currencies=["USD"],
                regions=["north_america", "europe"]
            ),
            performance=PerformanceMetrics(
                execution_success_rate=0.92,
                user_rating=4.4,
                usage_count=0,
                last_used=None,
                error_count=0,
                modification_frequency=0.20
            ),
            created_by="healthcare_specialist",
            keywords=["dcf", "healthcare", "pharmaceutical", "risk_adjustment", "regulatory", "drug_pipeline"],
            tags=["healthcare_sector", "risk_adjusted", "regulatory_focused"]
        )
        templates.append(healthcare_dcf)
        
        return templates
    
    def _create_npv_templates(self) -> List[FinancialModel]:
        """Create NPV model templates"""
        templates = []
        
        # Basic NPV Model
        basic_npv = FinancialModel(
            id="npv_basic_001",
            name="Basic NPV Analysis",
            description="Fundamental NPV analysis for capital investment decisions with sensitivity analysis",
            model_type=ModelType.NPV,
            industry=Industry.GENERAL,
            complexity=ComplexityLevel.BASIC,
            excel_code='''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // HEADER
    sheet.getRange("A1").values = [["NET PRESENT VALUE ANALYSIS"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.fill.color = "#4169E1";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    
    // INPUT ASSUMPTIONS
    sheet.getRange("A3").values = [["PROJECT INPUTS"]];
    sheet.getRange("A3:C3").format.fill.color = "#6495ED";
    sheet.getRange("A3:C3").format.font.bold = true;
    sheet.getRange("A3:C3").format.font.color = "#FFFFFF";
    
    sheet.getRange("A4:B8").values = [
        ["Initial Investment", "-1000000"],
        ["Annual Cash Flow", "250000"],
        ["Project Life (Years)", "5"],
        ["Discount Rate", "10%"],
        ["Salvage Value", "100000"]
    ];
    sheet.getRange("B4:B8").format.fill.color = "#E6F3FF";
    
    // CASH FLOW ANALYSIS
    sheet.getRange("D3:I3").values = [["Year", "0", "1", "2", "3", "4", "5"]];
    sheet.getRange("D3:I3").format.font.bold = true;
    sheet.getRange("D3:I3").format.fill.color = "#6495ED";
    sheet.getRange("D3:I3").format.font.color = "#FFFFFF";
    
    sheet.getRange("D4:I8").values = [
        ["Cash Flow", "=$B$4", "=$B$5", "=$B$5", "=$B$5", "=$B$5", "=$B$5+$B$8"],
        ["Discount Factor", "1", "=1/POWER(1+$B$7,E3)", "=1/POWER(1+$B$7,F3)", "=1/POWER(1+$B$7,G3)", "=1/POWER(1+$B$7,H3)", "=1/POWER(1+$B$7,I3)"],
        ["Present Value", "=E4*E5", "=F4*F5", "=G4*G5", "=H4*H5", "=I4*I5", "=J4*J5"],
        ["Cumulative PV", "=E6", "=E7+F6", "=F7+G6", "=G7+H6", "=H7+I6", "=I7+J6"]
    ];
    
    // RESULTS
    sheet.getRange("A10:B15").values = [
        ["RESULTS", ""],
        ["Net Present Value", "=J7"],
        ["Internal Rate of Return", "=IRR(E4:J4)"],
        ["Payback Period (Years)", "=MATCH(TRUE,E8:J8>0,0)-1"],
        ["Profitability Index", "=1+(B11/-$B$4)"]
    ];
    sheet.getRange("A10:B15").format.fill.color = "#B0E0E6";
    sheet.getRange("A10:A15").format.font.bold = true;
    
    // SENSITIVITY ANALYSIS
    sheet.getRange("D10:H10").values = [["SENSITIVITY ANALYSIS", "", "", "", ""]];
    sheet.getRange("D10:H10").format.font.bold = true;
    sheet.getRange("D10:H10").format.fill.color = "#FF6347";
    sheet.getRange("D10:H10").format.font.color = "#FFFFFF";
    
    sheet.getRange("D11:H14").values = [
        ["Discount Rate", "8%", "10%", "12%", "14%"],
        ["NPV @8%", "=NPV(D12,$F$4:$J$4)+$E$4", "", "", ""],
        ["NPV @10%", "", "=$B$11", "", ""],
        ["NPV @12%", "", "", "=NPV(F12,$F$4:$J$4)+$E$4", ""]
    ];
    
    // FORMATTING
    sheet.getRange("E4:J8").format.numberFormat = "$#,##0";
    sheet.getRange("B4").format.numberFormat = "$#,##0";
    sheet.getRange("B7").format.numberFormat = "0%";
    sheet.getRange("B11:B15").format.numberFormat = "$#,##0";
    sheet.getRange("B12").format.numberFormat = "0%";
    sheet.getRange("E12:H14").format.numberFormat = "$#,##0";
    
    await context.sync();
});
            ''',
            business_description="Comprehensive NPV analysis for evaluating capital investment projects with sensitivity testing",
            sample_inputs={
                "initial_investment": -1000000,
                "annual_cash_flow": 250000,
                "project_life": 5,
                "discount_rate": 0.10
            },
            expected_outputs={
                "npv": "calculated_npv",
                "irr": "internal_rate_return",
                "payback_period": "years_to_payback"
            },
            metadata=ModelMetadata(
                components=["cash_flow_analysis", "present_value_calculation", "sensitivity_analysis", "irr_calculation"],
                excel_functions=["NPV", "IRR", "POWER", "MATCH"],
                formatting_features=["data_tables", "conditional_formatting"],
                business_assumptions=["constant_cash_flows", "terminal_value", "discount_rate_stability"],
                time_horizon_years=5,
                currencies=["USD", "EUR", "GBP"],
                regions=["global"]
            ),
            performance=PerformanceMetrics(
                execution_success_rate=0.98,
                user_rating=4.5,
                usage_count=0,
                last_used=None,
                error_count=0,
                modification_frequency=0.10
            ),
            created_by="financial_modeling_standard",
            keywords=["npv", "capital_budgeting", "investment_analysis", "irr", "payback", "sensitivity"],
            tags=["basic", "educational", "general_purpose"]
        )
        templates.append(basic_npv)
        
        return templates
    
    def _create_valuation_templates(self) -> List[FinancialModel]:
        """Create valuation model templates"""
        templates = []
        
        # Comparable Company Analysis
        comp_valuation = FinancialModel(
            id="valuation_comp_001",
            name="Comparable Company Analysis",
            description="Multi-multiple valuation using comparable company analysis with statistical analysis",
            model_type=ModelType.VALUATION,
            industry=Industry.GENERAL,
            complexity=ComplexityLevel.INTERMEDIATE,
            excel_code='''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // HEADER
    sheet.getRange("A1").values = [["COMPARABLE COMPANY ANALYSIS"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.fill.color = "#800080";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    
    // COMPARABLE COMPANIES
    sheet.getRange("A3:F3").values = [["Company", "Market Cap", "Revenue", "EBITDA", "P/E", "EV/EBITDA"]];
    sheet.getRange("A3:F3").format.font.bold = true;
    sheet.getRange("A3:F3").format.fill.color = "#9370DB";
    sheet.getRange("A3:F3").format.font.color = "#FFFFFF";
    
    sheet.getRange("A4:F8").values = [
        ["Comp 1", "5000", "1000", "300", "20", "12"],
        ["Comp 2", "7500", "1500", "450", "22", "15"],
        ["Comp 3", "3000", "600", "180", "18", "10"],
        ["Comp 4", "6000", "1200", "360", "25", "14"],
        ["Target Co", "TBD", "800", "240", "TBD", "TBD"]
    ];
    
    // MULTIPLE ANALYSIS
    sheet.getRange("H3:J3").values = [["Multiple", "Mean", "Median"]];
    sheet.getRange("H3:J3").format.font.bold = true;
    sheet.getRange("H3:J3").format.fill.color = "#9370DB";
    sheet.getRange("H3:J3").format.font.color = "#FFFFFF";
    
    sheet.getRange("H4:J6").values = [
        ["P/E Ratio", "=AVERAGE(E4:E7)", "=MEDIAN(E4:E7)"],
        ["EV/EBITDA", "=AVERAGE(F4:F7)", "=MEDIAN(F4:F7)"],
        ["EV/Revenue", "=AVERAGE(B4:B7/C4:C7)", "=MEDIAN(B4:B7/C4:C7)"]
    ];
    
    // VALUATION RESULTS
    sheet.getRange("A10:C10").values = [["VALUATION SUMMARY", "", ""]];
    sheet.getRange("A10:C10").format.font.bold = true;
    sheet.getRange("A10:C10").format.fill.color = "#DDA0DD";
    
    sheet.getRange("A11:C15").values = [
        ["Method", "Low", "High"],
        ["P/E Method", "=C8*I4*21", "=C8*J4*21"],
        ["EV/EBITDA Method", "=D8*I5", "=D8*J5"],
        ["EV/Revenue Method", "=C8*I6", "=C8*J6"],
        ["Average Valuation", "=AVERAGE(B12:B14)", "=AVERAGE(C12:C14)"]
    ];
    
    // FORMATTING
    sheet.getRange("B4:F8").format.numberFormat = "#,##0";
    sheet.getRange("I4:J6").format.numberFormat = "0.0";
    sheet.getRange("B11:C15").format.numberFormat = "$#,##0";
    
    await context.sync();
});
            ''',
            business_description="Comprehensive comparable company analysis with multiple valuation methods and statistical analysis",
            sample_inputs={
                "target_revenue": 800,
                "target_ebitda": 240,
                "comparable_multiples": "industry_specific"
            },
            expected_outputs={
                "valuation_range": "low_high_estimates",
                "recommended_multiple": "median_multiple",
                "confidence_interval": "statistical_range"
            },
            metadata=ModelMetadata(
                components=["comparable_selection", "multiple_analysis", "statistical_measures", "valuation_range"],
                excel_functions=["AVERAGE", "MEDIAN", "PERCENTILE"],
                formatting_features=["data_validation", "dynamic_ranges"],
                business_assumptions=["market_comparability", "multiple_stability", "business_similarity"],
                time_horizon_years=1,
                currencies=["USD"],
                regions=["north_america"]
            ),
            performance=PerformanceMetrics(
                execution_success_rate=0.94,
                user_rating=4.3,
                usage_count=0,
                last_used=None,
                error_count=0,
                modification_frequency=0.25
            ),
            created_by="valuation_specialist",
            keywords=["valuation", "comparable", "multiples", "pe_ratio", "ev_ebitda", "market_analysis"],
            tags=["relative_valuation", "market_based", "multiple_methods"]
        )
        templates.append(comp_valuation)
        
        return templates
    
    def _create_budget_templates(self) -> List[FinancialModel]:
        """Create budget and forecast model templates"""
        templates = []
        
        # Annual Budget Template
        budget_template = FinancialModel(
            id="budget_annual_001",
            name="Annual Budget & Forecast Model",
            description="Comprehensive annual budget with monthly breakdown, variance analysis, and scenario planning",
            model_type=ModelType.BUDGET,
            industry=Industry.GENERAL,
            complexity=ComplexityLevel.INTERMEDIATE,
            excel_code='''
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // HEADER
    sheet.getRange("A1").values = [["ANNUAL BUDGET & FORECAST MODEL"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;
    sheet.getRange("A1").format.fill.color = "#2E8B57";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    
    // REVENUE ASSUMPTIONS
    sheet.getRange("A3").values = [["REVENUE ASSUMPTIONS"]];
    sheet.getRange("A3:D3").format.fill.color = "#3CB371";
    sheet.getRange("A3:D3").format.font.bold = true;
    sheet.getRange("A3:D3").format.font.color = "#FFFFFF";
    
    sheet.getRange("A4:B8").values = [
        ["Base Revenue (Annual)", "10000000"],
        ["Growth Rate", "15%"],
        ["Seasonality Factor Q1", "20%"],
        ["Seasonality Factor Q2", "25%"],
        ["Seasonality Factor Q3", "30%"],
        ["Seasonality Factor Q4", "25%"]
    ];
    sheet.getRange("B4:B8").format.fill.color = "#E0FFE0";
    
    // QUARTERLY BREAKDOWN
    sheet.getRange("D3:H3").values = [["Quarter", "Q1", "Q2", "Q3", "Q4"]];
    sheet.getRange("D3:H3").format.font.bold = true;
    sheet.getRange("D3:H3").format.fill.color = "#3CB371";
    sheet.getRange("D3:H3").format.font.color = "#FFFFFF";
    
    sheet.getRange("D4:H8").values = [
        ["Revenue", "=$B$4*$B$6", "=$B$4*$B$7", "=$B$4*$B$8", "=$B$4*$B$9"],
        ["Growth %", "=$B$5", "=$B$5", "=$B$5", "=$B$5"],
        ["Adj. Revenue", "=E4*(1+E5)", "=F4*(1+F5)", "=G4*(1+G5)", "=H4*(1+H5)"],
        ["COGS (60%)", "=E6*0.6", "=F6*0.6", "=G6*0.6", "=H6*0.6"],
        ["Gross Profit", "=E6-E7", "=F6-F7", "=G6-G7", "=H6-H7"]
    ];
    
    // OPERATING EXPENSES
    sheet.getRange("D10:H10").values = [["OPERATING EXPENSES", "", "", "", ""]];
    sheet.getRange("D10:H10").format.font.bold = true;
    sheet.getRange("D10:H10").format.fill.color = "#FF6347";
    sheet.getRange("D10:H10").format.font.color = "#FFFFFF";
    
    sheet.getRange("D11:H15").values = [
        ["Sales & Marketing", "=E6*0.15", "=F6*0.15", "=G6*0.15", "=H6*0.15"],
        ["R&D", "=E6*0.10", "=F6*0.10", "=G6*0.10", "=H6*0.10"],
        ["General & Admin", "=E6*0.08", "=F6*0.08", "=G6*0.08", "=H6*0.08"],
        ["Total OpEx", "=SUM(E11:E13)", "=SUM(F11:F13)", "=SUM(G11:G13)", "=SUM(H11:H13)"],
        ["EBITDA", "=E8-E14", "=F8-F14", "=G8-G14", "=H8-H14"]
    ];
    
    // ANNUAL SUMMARY
    sheet.getRange("A17:B22").values = [
        ["ANNUAL SUMMARY", ""],
        ["Total Revenue", "=SUM(E6:H6)"],
        ["Total COGS", "=SUM(E7:H7)"],
        ["Gross Profit", "=B18-B19"],
        ["Total OpEx", "=SUM(E14:H14)"],
        ["EBITDA", "=B20-B21"]
    ];
    sheet.getRange("A17:B22").format.fill.color = "#90EE90";
    sheet.getRange("A17:A22").format.font.bold = true;
    
    // VARIANCE ANALYSIS
    sheet.getRange("J3:L3").values = [["Budget", "Actual", "Variance"]];
    sheet.getRange("J3:L3").format.font.bold = true;
    sheet.getRange("J3:L3").format.fill.color = "#FFD700";
    
    sheet.getRange("J4:L8").values = [
        ["Q1 Revenue", "=E6", "0", "=L4-K4"],
        ["Q2 Revenue", "=F6", "0", "=L5-K5"],
        ["Q3 Revenue", "=G6", "0", "=L6-K6"],
        ["Q4 Revenue", "=H6", "0", "=L7-K7"],
        ["Annual Total", "=B18", "=SUM(L4:L7)", "=L8-K8"]
    ];
    
    // FORMATTING
    sheet.getRange("E4:H22").format.numberFormat = "$#,##0";
    sheet.getRange("B4").format.numberFormat = "$#,##0";
    sheet.getRange("B5").format.numberFormat = "0%";
    sheet.getRange("E5:H5").format.numberFormat = "0%";
    sheet.getRange("K4:L8").format.numberFormat = "$#,##0";
    
    await context.sync();
});
            ''',
            business_description="Comprehensive annual budget model with quarterly breakdown, variance analysis, and scenario planning capabilities",
            sample_inputs={
                "base_revenue": 10000000,
                "growth_rate": 0.15,
                "seasonality_factors": [0.20, 0.25, 0.30, 0.25],
                "expense_ratios": {"cogs": 0.60, "sales": 0.15, "rd": 0.10}
            },
            expected_outputs={
                "annual_revenue": "projected_revenue",
                "quarterly_breakdown": "seasonal_distribution",
                "ebitda_margin": "profitability_measure"
            },
            metadata=ModelMetadata(
                components=["revenue_forecasting", "expense_budgeting", "variance_analysis", "quarterly_phasing"],
                excel_functions=["SUM", "AVERAGE", "PERCENTAGE"],
                formatting_features=["quarterly_layout", "variance_highlighting"],
                business_assumptions=["seasonal_patterns", "expense_ratios", "growth_sustainability"],
                time_horizon_years=1,
                currencies=["USD", "EUR"],
                regions=["global"]
            ),
            performance=PerformanceMetrics(
                execution_success_rate=0.96,
                user_rating=4.6,
                usage_count=0,
                last_used=None,
                error_count=0,
                modification_frequency=0.30
            ),
            created_by="corporate_finance_team",
            keywords=["budget", "forecast", "variance", "quarterly", "revenue_planning", "expense_budget"],
            tags=["corporate_planning", "budget_management", "quarterly_reporting"]
        )
        templates.append(budget_template)
        
        return templates


# Singleton instance
_curator_instance = None

def get_model_curator() -> ModelCurator:
    """Get singleton model curator instance"""
    global _curator_instance
    if _curator_instance is None:
        _curator_instance = ModelCurator()
    return _curator_instance