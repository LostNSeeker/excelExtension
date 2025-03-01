// dcf.js - Discounted Cash Flow Model Template

/**
 * Creates a Discounted Cash Flow (DCF) model in Excel
 * @param {object} context - The Excel context
 * @param {object} params - Parameters for the DCF model
 * @returns {Promise<object>} - Information about the created model
 */
async function createDCFModel(context, params = {}) {
    // Default parameters
    const defaults = {
        companyName: "Sample Company",
        historicalYears: 3,
        projectionYears: 5,
        revenueGrowthRate: 0.05,
        ebitdaMargin: 0.25,
        taxRate: 0.25,
        depreciationRate: 0.05,
        capexPercentOfRevenue: 0.1,
        workingCapitalPercentOfRevenue: 0.15,
        discountRate: 0.10,
        perpetualGrowthRate: 0.02,
        exitMultiple: 8
    };

    // Merge defaults with provided parameters
    const modelParams = { ...defaults, ...params };
    
    // Create a new worksheet for the model if it doesn't exist
    let sheet;
    try {
        sheet = context.workbook.worksheets.getItem("DCF Model");
    } catch (error) {
        sheet = context.workbook.worksheets.add("DCF Model");
    }
    
    // Activate the worksheet
    sheet.activate();
    
    // Setup the model structure
    await setupModelStructure(sheet, modelParams);
    
    // Add headers and labels
    await addHeadersAndLabels(sheet, modelParams);
    
    // Create historical data section
    await createHistoricalSection(sheet, modelParams);
    
    // Create projection section
    await createProjectionSection(sheet, modelParams);
    
    // Create DCF valuation section
    await createValuationSection(sheet, modelParams);
    
    // Format the worksheet
    await formatWorksheet(sheet, modelParams);
    
    return {
        sheetName: sheet.name,
        modelType: "DCF",
        parameters: modelParams
    };
}

/**
 * Sets up the basic structure of the DCF model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function setupModelStructure(sheet, params) {
    // Clear the sheet
    sheet.getRange().clear();
    
    // Set column widths
    sheet.getRange("A:A").columnWidth = 250; // Labels
    sheet.getRange("B:Z").columnWidth = 120; // Data columns
    
    // Set title
    sheet.getRange("A1:G1").values = [[`${params.companyName} - Discounted Cash Flow (DCF) Model`, "", "", "", "", "", ""]];
    sheet.getRange("A1:G1").format.font.bold = true;
    sheet.getRange("A1:G1").format.font.size = 16;
    sheet.getRange("A1:G1").merge();
    
    // Set section headers
    sheet.getRange("A3").values = [["Model Assumptions"]];
    sheet.getRange("A3").format.font.bold = true;
    sheet.getRange("A3").format.font.size = 14;
    
    sheet.getRange("A10").values = [["Historical & Projected Financials"]];
    sheet.getRange("A10").format.font.bold = true;
    sheet.getRange("A10").format.font.size = 14;
    
    sheet.getRange("A30").values = [["DCF Valuation"]];
    sheet.getRange("A30").format.font.bold = true;
    sheet.getRange("A30").format.font.size = 14;
}

/**
 * Adds headers and labels to the DCF model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function addHeadersAndLabels(sheet, params) {
    // Assumption labels
    sheet.getRange("A4:A9").values = [
        ["Revenue Growth Rate"],
        ["EBITDA Margin"],
        ["Tax Rate"],
        ["Discount Rate (WACC)"],
        ["Perpetual Growth Rate"],
        ["Exit Multiple (EV/EBITDA)"]
    ];
    
    // Assumption values
    sheet.getRange("B4:B9").values = [
        [params.revenueGrowthRate],
        [params.ebitdaMargin],
        [params.taxRate],
        [params.discountRate],
        [params.perpetualGrowthRate],
        [params.exitMultiple]
    ];
    sheet.getRange("B4:B9").numberFormat = "0.0%";
    sheet.getRange("B9").numberFormat = "0.0";
    
    // Financial statement labels
    sheet.getRange("A12:A29").values = [
        ["Income Statement"],
        ["Revenue"],
        ["Growth Rate"],
        ["EBITDA"],
        ["EBITDA Margin"],
        ["Depreciation & Amortization"],
        ["EBIT"],
        ["EBIT Margin"],
        ["Taxes"],
        ["NOPAT"],
        [""],
        ["Cash Flow Statement"],
        ["NOPAT"],
        ["Add: Depreciation & Amortization"],
        ["Less: Capital Expenditures"],
        ["Less: Change in Working Capital"],
        ["Unlevered Free Cash Flow"],
        [""]
    ];
    
    // DCF valuation labels
    sheet.getRange("A32:A43").values = [
        ["Unlevered Free Cash Flow"],
        ["Discount Period"],
        ["Discount Factor"],
        ["Present Value of FCF"],
        [""],
        ["Sum of PV of FCF"],
        ["Terminal Value"],
        ["PV of Terminal Value"],
        [""],
        ["Enterprise Value"],
        ["Less: Net Debt"],
        ["Equity Value"]
    ];
}

/**
 * Creates the historical data section of the DCF model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createHistoricalSection(sheet, params) {
    // Create year headers for historical data
    const currentYear = new Date().getFullYear();
    const yearHeaders = [];
    
    for (let i = 0; i < params.historicalYears; i++) {
        yearHeaders.push([`FY ${currentYear - params.historicalYears + i}`]);
    }
    
    // Add headers to the sheet
    sheet.getRange(11, 2, 1, params.historicalYears).values = [yearHeaders.map(item => item[0])];
    sheet.getRange(11, 2, 1, params.historicalYears).format.font.bold = true;
    
    // Simulate historical data (in a real implementation, this would come from user input or data source)
    let baseRevenue = 1000; // Starting revenue in millions
    
    for (let year = 0; year < params.historicalYears; year++) {
        const col = year + 2; // Column B is index 2
        const growthRate = 0.08 + Math.random() * 0.04; // Random growth between 8-12%
        
        if (year > 0) {
            baseRevenue = baseRevenue * (1 + growthRate);
        }
        
        // Revenue
        sheet.getRange(13, col).values = [[baseRevenue]];
        
        // Growth Rate (not applicable for first historical year)
        if (year > 0) {
            sheet.getRange(14, col).formulas = [[`=(${getColumnLetter(col)}13/${getColumnLetter(col-1)}13-1)`]];
        }
        
        // EBITDA
        const ebitdaMargin = 0.2 + Math.random() * 0.1; // Random margin between 20-30%
        sheet.getRange(15, col).formulas = [[`=${getColumnLetter(col)}13*${ebitdaMargin}`]];
        
        // EBITDA Margin
        sheet.getRange(16, col).formulas = [[`=${getColumnLetter(col)}15/${getColumnLetter(col)}13`]];
        
        // Depreciation & Amortization
        sheet.getRange(17, col).formulas = [[`=${getColumnLetter(col)}13*${params.depreciationRate}`]];
        
        // EBIT
        sheet.getRange(18, col).formulas = [[`=${getColumnLetter(col)}15-${getColumnLetter(col)}17`]];
        
        // EBIT Margin
        sheet.getRange(19, col).formulas = [[`=${getColumnLetter(col)}18/${getColumnLetter(col)}13`]];
        
        // Taxes
        sheet.getRange(20, col).formulas = [[`=IF(${getColumnLetter(col)}18>0,${getColumnLetter(col)}18*$B$6,0)`]];
        
        // NOPAT
        sheet.getRange(21, col).formulas = [[`=${getColumnLetter(col)}18-${getColumnLetter(col)}20`]];
        
        // Cash Flow items
        // NOPAT (same as above)
        sheet.getRange(24, col).formulas = [[`=${getColumnLetter(col)}21`]];
        
        // Add: Depreciation & Amortization
        sheet.getRange(25, col).formulas = [[`=${getColumnLetter(col)}17`]];
        
        // Less: Capital Expenditures
        sheet.getRange(26, col).formulas = [[`=${getColumnLetter(col)}13*${params.capexPercentOfRevenue}`]];
        
        // Less: Change in Working Capital
        if (year > 0) {
            sheet.getRange(27, col).formulas = [[`=(${getColumnLetter(col)}13-${getColumnLetter(col-1)}13)*${params.workingCapitalPercentOfRevenue}`]];
        } else {
            sheet.getRange(27, col).values = [[0]];
        }
        
        // Unlevered Free Cash Flow
        sheet.getRange(28, col).formulas = [[`=${getColumnLetter(col)}24+${getColumnLetter(col)}25-${getColumnLetter(col)}26-${getColumnLetter(col)}27`]];
    }
}

/**
 * Creates the projection section of the DCF model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createProjectionSection(sheet, params) {
    // Create year headers for projections
    const currentYear = new Date().getFullYear();
    const yearHeaders = [];
    
    for (let i = 0; i < params.projectionYears; i++) {
        yearHeaders.push([`FY ${currentYear + i + 1}`]);
    }
    
    // Add headers to the sheet
    const startCol = 2 + params.historicalYears; // Column after historical data
    sheet.getRange(11, startCol, 1, params.projectionYears).values = [yearHeaders.map(item => item[0])];
    sheet.getRange(11, startCol, 1, params.projectionYears).format.font.bold = true;
    
    // Create projection formulas
    for (let year = 0; year < params.projectionYears; year++) {
        const col = startCol + year;
        const prevCol = col - 1;
        
        // Revenue
        if (year === 0) {
            // First projection year based on last historical year
            sheet.getRange(13, col).formulas = [[`=${getColumnLetter(prevCol)}13*(1+$B$4)`]];
        } else {
            // Subsequent projection years
            sheet.getRange(13, col).formulas = [[`=${getColumnLetter(prevCol)}13*(1+$B$4)`]];
        }
        
        // Growth Rate
        sheet.getRange(14, col).formulas = [[`=(${getColumnLetter(col)}13/${getColumnLetter(prevCol)}13)-1`]];
        
        // EBITDA
        sheet.getRange(15, col).formulas = [[`=${getColumnLetter(col)}13*$B$5`]];
        
        // EBITDA Margin
        sheet.getRange(16, col).formulas = [[`=${getColumnLetter(col)}15/${getColumnLetter(col)}13`]];
        
        // Depreciation & Amortization
        sheet.getRange(17, col).formulas = [[`=${getColumnLetter(col)}13*${params.depreciationRate}`]];
        
        // EBIT
        sheet.getRange(18, col).formulas = [[`=${getColumnLetter(col)}15-${getColumnLetter(col)}17`]];
        
        // EBIT Margin
        sheet.getRange(19, col).formulas = [[`=${getColumnLetter(col)}18/${getColumnLetter(col)}13`]];
        
        // Taxes
        sheet.getRange(20, col).formulas = [[`=IF(${getColumnLetter(col)}18>0,${getColumnLetter(col)}18*$B$6,0)`]];
        
        // NOPAT
        sheet.getRange(21, col).formulas = [[`=${getColumnLetter(col)}18-${getColumnLetter(col)}20`]];
        
        // Cash Flow items
        // NOPAT (same as above)
        sheet.getRange(24, col).formulas = [[`=${getColumnLetter(col)}21`]];
        
        // Add: Depreciation & Amortization
        sheet.getRange(25, col).formulas = [[`=${getColumnLetter(col)}17`]];
        
        // Less: Capital Expenditures
        sheet.getRange(26, col).formulas = [[`=${getColumnLetter(col)}13*${params.capexPercentOfRevenue}`]];
        
        // Less: Change in Working Capital
        sheet.getRange(27, col).formulas = [[`=(${getColumnLetter(col)}13-${getColumnLetter(prevCol)}13)*${params.workingCapitalPercentOfRevenue}`]];
        
        // Unlevered Free Cash Flow
        sheet.getRange(28, col).formulas = [[`=${getColumnLetter(col)}24+${getColumnLetter(col)}25-${getColumnLetter(col)}26-${getColumnLetter(col)}27`]];
    }
}

/**
 * Creates the DCF valuation section of the model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createValuationSection(sheet, params) {
    const startCol = 2 + params.historicalYears; // Column after historical data
    const endCol = startCol + params.projectionYears - 1;
    
    // Copy unlevered FCF from above
    for (let year = 0; year < params.projectionYears; year++) {
        const col = startCol + year;
        
        // Unlevered FCF
        sheet.getRange(32, col).formulas = [[`=${getColumnLetter(col)}28`]];
        
        // Discount Period (midyear convention)
        sheet.getRange(33, col).values = [[year + 0.5]];
        
        // Discount Factor
        sheet.getRange(34, col).formulas = [[`=1/POWER(1+$B$7,${getColumnLetter(col)}33)`]];
        
        // Present Value of FCF
        sheet.getRange(35, col).formulas = [[`=${getColumnLetter(col)}32*${getColumnLetter(col)}34`]];
    }
    
    // Sum of PV of FCF
    sheet.getRange(37, 2).formulas = [[`=SUM(${getColumnLetter(startCol)}35:${getColumnLetter(endCol)}35)`]];
    
    // Last year EBITDA for terminal value calculation
    sheet.getRange(38, 2).formulas = [[`=${getColumnLetter(endCol)}15*(1+$B$8)`]];
    
    // Terminal Value - Exit Multiple Method
    sheet.getRange(38, 3).formulas = [[`=$B38*$B$9`]];
    
    // Terminal Value - Perpetuity Growth Method
    sheet.getRange(38, 4).formulas = [[`=${getColumnLetter(endCol)}28*(1+$B$8)/(($B$7-$B$8))`]];
    
    // PV of Terminal Value
    sheet.getRange(39, 2).formulas = [[`=$C$38*${getColumnLetter(endCol)}34`]];
    
    // Enterprise Value
    sheet.getRange(41, 2).formulas = [[`=$B$37+$B$39`]];
    
    // Net Debt placeholder (to be filled by user)
    sheet.getRange(42, 2).values = [[0]];
    
    // Equity Value
    sheet.getRange(43, 2).formulas = [[`=$B$41-$B$42`]];
    
    // Add terminal value method note
    sheet.getRange(38, 5).values = [["<-- Terminal Value Calculations"]];
    sheet.getRange(38, 5).format.font.italic = true;
    sheet.getRange(38, 5).format.font.color = "#666666";
    
    sheet.getRange(39, 3).values = [["Exit Multiple Method"]];
    sheet.getRange(39, 3).format.font.italic = true;
    
    sheet.getRange(39, 4).values = [["Perpetuity Growth Method"]];
    sheet.getRange(39, 4).format.font.italic = true;
}

/**
 * Formats the DCF worksheet
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function formatWorksheet(sheet, params) {
    const lastCol = 2 + params.historicalYears + params.projectionYears - 1;
    
    // Format headers
    sheet.getRange(11, 2, 1, lastCol - 1).format.font.bold = true;
    
    // Format numbers
    // Currency formatting for monetary values
    sheet.getRange(13, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(15, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(17, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(18, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(20, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(21, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(24, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(25, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(26, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(27, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(28, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(32, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange(35, 2, 1, lastCol - 1).numberFormat = "$#,##0.0,, \"M\"";
    
    // Percentage formatting
    sheet.getRange(14, 2, 1, lastCol - 1).numberFormat = "0.0%";
    sheet.getRange(16, 2, 1, lastCol - 1).numberFormat = "0.0%";
    sheet.getRange(19, 2, 1, lastCol - 1).numberFormat = "0.0%";
    sheet.getRange(34, 2, 1, lastCol - 1).numberFormat = "0.000";
    
    // Format valuation section
    sheet.getRange("B37:B43").numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange("B38").numberFormat = "$#,##0.0,, \"M\"";
    sheet.getRange("C38:D38").numberFormat = "$#,##0.0,, \"M\"";
    
    // Highlight important cells
    sheet.getRange("B41").format.fill.color = "#E6F0FF";
    sheet.getRange("B43").format.fill.color = "#E6F0FF";
    sheet.getRange("B41").format.font.bold = true;
    sheet.getRange("B43").format.font.bold = true;
    
    // Add section borders
    sheet.getRange(10, 1, 1, lastCol).format.borders.bottom.style = "Continuous";
    sheet.getRange(30, 1, 1, lastCol).format.borders.bottom.style = "Continuous";
    
    // Add title and subtitle for historical and projection sections
    sheet.getRange(11, 2, 1, params.historicalYears).format.fill.color = "#F0F0F0";
    sheet.getRange(11, startCol, 1, params.projectionYears).format.fill.color = "#E6F0FF";
    
    // Add historical/projection labels
    sheet.getRange("B10").values = [["Historical"]];
    sheet.getRange("B10").format.font.italic = true;
    sheet.getRange(`${getColumnLetter(2 + params.historicalYears)}10`).values = [["Projection"]];
    sheet.getRange(`${getColumnLetter(2 + params.historicalYears)}10`).format.font.italic = true;
}

/**
 * Utility function to convert column index to Excel column letter
 * @param {number} column - 1-based column index
 * @returns {string} - Excel column letter (A, B, C, ...)
 */
function getColumnLetter(column) {
    let dividend = column;
    let columnName = '';
    let modulo;
    
    while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    
    return columnName;
}

module.exports = {
    createDCFModel
};