// lbo.js - Leveraged Buyout Model Template

/**
 * Creates a Leveraged Buyout (LBO) model in Excel
 * @param {object} context - The Excel context
 * @param {object} params - Parameters for the LBO model
 * @returns {Promise<object>} - Information about the created model
 */
async function createLBOModel(context, params = {}) {
    // Default parameters
    const defaults = {
        companyName: "Target Company",
        purchasePrice: 1000, // in millions
        entryMultiple: 8.0,  // EV/EBITDA multiple
        entryYear: new Date().getFullYear(),
        projectionYears: 5,
        exitMultiple: 9.0,
        debtToEbitda: 4.0,
        interestRate: 0.06,
        taxRate: 0.25,
        revenueGrowthRate: 0.05,
        ebitdaMargin: 0.30,
        capexPercentOfRevenue: 0.04,
        depreciationPercentOfRevenue: 0.03,
        workingCapitalPercentOfRevenue: 0.10,
        debtRepaymentPercentOfEBITDA: 0.50
    };

    // Merge defaults with provided parameters
    const modelParams = { ...defaults, ...params };
    
    // Create a new worksheet for the model if it doesn't exist
    let sheet;
    try {
        sheet = context.workbook.worksheets.getItem("LBO Model");
    } catch (error) {
        sheet = context.workbook.worksheets.add("LBO Model");
    }
    
    // Activate the worksheet
    sheet.activate();
    
    // Setup the model structure
    await setupModelStructure(sheet, modelParams);
    
    // Create transaction structure section
    await createTransactionSection(sheet, modelParams);
    
    // Create financial projections section
    await createProjectionsSection(sheet, modelParams);
    
    // Create debt schedule section
    await createDebtScheduleSection(sheet, modelParams);
    
    // Create returns analysis section
    await createReturnsSection(sheet, modelParams);
    
    // Format the worksheet
    await formatWorksheet(sheet, modelParams);
    
    return {
        sheetName: sheet.name,
        modelType: "LBO",
        parameters: modelParams
    };
}

/**
 * Sets up the basic structure of the LBO model
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
    sheet.getRange("A1:G1").values = [[`${params.companyName} - Leveraged Buyout (LBO) Model`, "", "", "", "", "", ""]];
    sheet.getRange("A1:G1").format.font.bold = true;
    sheet.getRange("A1:G1").format.font.size = 16;
    sheet.getRange("A1:G1").merge();
}

/**
 * Creates the transaction structure section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createTransactionSection(sheet, params) {
    // Transaction section header
    sheet.getRange("A3").values = [["Transaction Structure"]];
    sheet.getRange("A3").format.font.bold = true;
    sheet.getRange("A3").format.font.size = 14;
    
    // Year headers
    const entryYear = params.entryYear;
    sheet.getRange("B4").values = [["Entry"]];
    sheet.getRange("B4").format.font.bold = true;
    
    // Transaction structure details
    sheet.getRange("A5:A16").values = [
        ["Purchase Metrics"],
        ["Purchase Price ($M)"],
        ["LTM EBITDA ($M)"],
        ["Entry Multiple (EV/EBITDA)"],
        [""],
        ["Sources & Uses"],
        ["Sources"],
        ["  Debt"],
        ["  Equity"],
        ["  Total Sources"],
        ["Uses"],
        ["  Purchase Equity"]
    ];
    
    // Calculate LTM EBITDA based on purchase price and entry multiple
    const ltmEbitda = params.purchasePrice / params.entryMultiple;
    
    // Calculate debt based on debt to EBITDA ratio
    const debt = ltmEbitda * params.debtToEbitda;
    
    // Calculate equity contribution
    const equity = params.purchasePrice - debt;
    
    // Transaction values
    sheet.getRange("B6").values = [[params.purchasePrice]];
    sheet.getRange("B7").values = [[ltmEbitda]];
    sheet.getRange("B8").values = [[params.entryMultiple]];
    
    // Sources & Uses
    sheet.getRange("B11").values = [[debt]];
    sheet.getRange("B12").values = [[equity]];
    sheet.getRange("B13").formulas = [["=SUM(B11:B12)"]];
    sheet.getRange("B15").values = [[params.purchasePrice]];
    
    // Format values
    sheet.getRange("B6:B8").numberFormat = "#,##0.0";
    sheet.getRange("B11:B13").numberFormat = "#,##0.0";
    sheet.getRange("B15").numberFormat = "#,##0.0";
}

/**
 * Creates the projections section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createProjectionsSection(sheet, params) {
    // Projections section header
    sheet.getRange("A18").values = [["Financial Projections"]];
    sheet.getRange("A18").format.font.bold = true;
    sheet.getRange("A18").format.font.size = 14;
    
    // Year headers
    const years = [];
    for (let i = 0; i <= params.projectionYears; i++) {
        years.push([`Year ${i}`]);
    }
    sheet.getRange(19, 2, 1, params.projectionYears + 1).values = [years.map(y => y[0])];
    sheet.getRange(19, 2, 1, params.projectionYears + 1).format.font.bold = true;
    
    // Income statement labels
    sheet.getRange("A20:A30").values = [
        ["Income Statement"],
        ["Revenue"],
        ["Growth %"],
        ["EBITDA"],
        ["EBITDA Margin"],
        ["Depreciation & Amortization"],
        ["EBIT"],
        ["Interest Expense"],
        ["EBT"],
        ["Taxes"],
        ["Net Income"]
    ];
    
    // Cash flow labels
    sheet.getRange("A32:A38").values = [
        ["Cash Flow"],
        ["Net Income"],
        ["Add: Depreciation & Amortization"],
        ["Less: Capital Expenditures"],
        ["Less: Change in Working Capital"],
        ["Less: Debt Repayment"],
        ["Free Cash Flow to Equity"]
    ];
    
    // Calculate base revenue (based on LTM EBITDA and EBITDA margin)
    const baseRevenue = params.purchasePrice / params.entryMultiple / params.ebitdaMargin;
    
    // Year 0 (entry year) financials
    sheet.getRange("B22").values = [[baseRevenue]];
    sheet.getRange("B23").values = [["--"]];
    sheet.getRange("B24").formulas = [["=B22*" + params.ebitdaMargin]];
    sheet.getRange("B25").formulas = [["=B24/B22"]];
    sheet.getRange("B26").formulas = [["=B22*" + params.depreciationPercentOfRevenue]];
    sheet.getRange("B27").formulas = [["=B24-B26"]];
    sheet.getRange("B28").formulas = [["=B11*" + params.interestRate]];
    sheet.getRange("B29").formulas = [["=B27-B28"]];
    sheet.getRange("B30").formulas = [["=IF(B29>0,B29*" + params.taxRate + ",0)"]];
    sheet.getRange("B31").formulas = [["=B29-B30"]];
    
    // Year 0 cash flow
    sheet.getRange("B33").formulas = [["=B31"]];
    sheet.getRange("B34").formulas = [["=B26"]];
    sheet.getRange("B35").formulas = [["=B22*" + params.capexPercentOfRevenue]];
    sheet.getRange("B36").values = [[0]]; // No change in WC for entry year
    sheet.getRange("B37").formulas = [["=B24*" + params.debtRepaymentPercentOfEBITDA]];
    sheet.getRange("B38").formulas = [["=B33+B34-B35-B36-B37"]];
    
    // Projection years
    for (let year = 1; year <= params.projectionYears; year++) {
        const col = year + 2; // Column C is year 1, D is year 2, etc.
        const prevCol = col - 1;
        
        // Income statement projections
        sheet.getRange(22, col).formulas = [[`=${getColumnLetter(prevCol)}22*(1+${params.revenueGrowthRate})`]];
        sheet.getRange(23, col).formulas = [[`=(${getColumnLetter(col)}22/${getColumnLetter(prevCol)}22)-1`]];
        sheet.getRange(24, col).formulas = [[`=${getColumnLetter(col)}22*${params.ebitdaMargin}`]];
        sheet.getRange(25, col).formulas = [[`=${getColumnLetter(col)}24/${getColumnLetter(col)}22`]];
        sheet.getRange(26, col).formulas = [[`=${getColumnLetter(col)}22*${params.depreciationPercentOfRevenue}`]];
        sheet.getRange(27, col).formulas = [[`=${getColumnLetter(col)}24-${getColumnLetter(col)}26`]];
        
        // Interest expense will be calculated after debt schedule is created
        sheet.getRange(28, col).formulas = [[`=0`]]; // Placeholder
        
        sheet.getRange(29, col).formulas = [[`=${getColumnLetter(col)}27-${getColumnLetter(col)}28`]];
        sheet.getRange(30, col).formulas = [[`=IF(${getColumnLetter(col)}29>0,${getColumnLetter(col)}29*${params.taxRate},0)`]];
        sheet.getRange(31, col).formulas = [[`=${getColumnLetter(col)}29-${getColumnLetter(col)}30`]];
        
        // Cash flow projections
        sheet.getRange(33, col).formulas = [[`=${getColumnLetter(col)}31`]];
        sheet.getRange(34, col).formulas = [[`=${getColumnLetter(col)}26`]];
        sheet.getRange(35, col).formulas = [[`=${getColumnLetter(col)}22*${params.capexPercentOfRevenue}`]];
        sheet.getRange(36, col).formulas = [[`=(${getColumnLetter(col)}22-${getColumnLetter(prevCol)}22)*${params.workingCapitalPercentOfRevenue}`]];
        sheet.getRange(37, col).formulas = [[`=${getColumnLetter(col)}24*${params.debtRepaymentPercentOfEBITDA}`]];
        sheet.getRange(38, col).formulas = [[`=${getColumnLetter(col)}33+${getColumnLetter(col)}34-${getColumnLetter(col)}35-${getColumnLetter(col)}36-${getColumnLetter(col)}37`]];
    }
    
    // Format numbers
    sheet.getRange(22, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(23, 2, 1, params.projectionYears + 1).numberFormat = "0.0%";
    sheet.getRange(24, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(25, 2, 1, params.projectionYears + 1).numberFormat = "0.0%";
    sheet.getRange(26, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(27, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(28, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(29, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(30, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(31, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(33, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(34, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(35, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(36, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(37, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
    sheet.getRange(38, 2, 1, params.projectionYears + 1).numberFormat = "#,##0.0";
}

/**
 * Creates the debt schedule section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createDebtScheduleSection(sheet, params) {
    // Debt schedule section header
    sheet.getRange("A40").values = [["Debt Schedule"]];
    sheet.getRange("A40").format.font.bold = true;
    sheet.getRange("A40").format.font.size = 14;
    
    // Debt schedule labels
    sheet.getRange("A41:A45").values = [
        ["Beginning Balance"],
        ["Repayments"],
        ["New Borrowings"],
        ["Ending Balance"],
        ["Interest Expense"]
    ];
    
    // Calculate debt amount
    const initialDebt = params.purchasePrice / params.entryMultiple * params.debtToEbitda;
    
    // Year 0 (entry year) debt schedule
    sheet.getRange("B41").values = [[initialDebt]];
    sheet.getRange("B42").formulas = [["=B37"]]; // Debt repayment from cash flow
    sheet.getRange("B43").values = [[0]]; // No new borrowings
    sheet.getRange("B44").formulas = [["=B41-B42+B43"]];
    sheet.getRange("B45").formulas = [["=B41*" + params.interestRate]];
    
    // Update interest expense in income statement
    sheet.getRange("B28").formulas = [["=B45"]];
    
    // Projection years
    for (let year = 1; year <= params.projectionYears; year++) {
        const col = year + 2; // Column C is year 1, D is year 2, etc.
        const prevCol = col - 1;
        
        // Debt schedule projections
        sheet.getRange(41, col).formulas = [[`=${getColumnLetter(prevCol)}44`]]; // Beginning balance is previous ending balance
        sheet.getRange(42, col).formulas = [[`=${getColumnLetter(col)}37`]]; // Debt repayment from cash flow
        sheet.getRange(43, col).values = [[0]]; // No new borrowings
        sheet.getRange(44, col).formulas = [[`=${getColumnLetter(col)}41-${getColumnLetter(col)}42+${getColumnLetter(col)}43`]];
        sheet.getRange(45, col).formulas = [[`=${getColumnLetter(col)}41*${params.interestRate}`]];
        
        // Update interest expense in income statement
        sheet.getRange(28, col).formulas = [[`=${getColumnLetter(col)}45`]];
    }
    
    // Format numbers
    sheet.getRange(41, 2, 5, params.projectionYears + 1).numberFormat = "#,##0.0";
}

/**
 * Creates the returns analysis section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createReturnsSection(sheet, params) {
    // Returns section header
    sheet.getRange("A47").values = [["Returns Analysis"]];
    sheet.getRange("A47").format.font.bold = true;
    sheet.getRange("A47").format.font.size = 14;
    
    // Returns labels
    sheet.getRange("A48:A55").values = [
        ["Exit Year EBITDA"],
        ["Exit Multiple"],
        ["Enterprise Value at Exit"],
        ["Exit Year Net Debt"],
        ["Implied Equity Value at Exit"],
        ["Initial Equity Investment"],
        ["Multiple of Money (MoM)"],
        ["Internal Rate of Return (IRR)"]
    ];
    
    // Exit year (last projection year)
    const exitCol = 2 + params.projectionYears;
    
    // Exit year calculations
    sheet.getRange("B49").values = [[params.exitMultiple]];
    sheet.getRange("B48").formulas = [[`=${getColumnLetter(exitCol)}24`]]; // Exit year EBITDA
    sheet.getRange("B50").formulas = [["=B48*B49"]]; // Exit enterprise value
    sheet.getRange("B51").formulas = [[`=${getColumnLetter(exitCol)}44`]]; // Exit year net debt
    sheet.getRange("B52").formulas = [["=B50-B51"]]; // Implied equity value
    sheet.getRange("B53").formulas = [["=B12"]]; // Initial equity investment
    sheet.getRange("B54").formulas = [["=B52/B53"]]; // Multiple of money
    
    // IRR calculation - needs to create cash flow array
    // Year 0 cash flow is negative initial equity
    let irrFormula = "=IRR({-B53,";
    
    // Intermediate cash flows are zeros (no dividends assumed)
    for (let year = 1; year < params.projectionYears; year++) {
        irrFormula += "0,";
    }
    
    // Final cash flow is exit equity value
    irrFormula += "B52})";
    
    sheet.getRange("B55").formulas = [[irrFormula]];
    
    // Format numbers
    sheet.getRange("B48").numberFormat = "#,##0.0";
    sheet.getRange("B49").numberFormat = "0.0";
    sheet.getRange("B50:B53").numberFormat = "#,##0.0";
    sheet.getRange("B54").numberFormat = "0.0x";
    sheet.getRange("B55").numberFormat = "0.0%";
}

/**
 * Formats the LBO worksheet
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function formatWorksheet(sheet, params) {
    // Add section borders
    sheet.getRange("A3:G3").format.borders.bottom.style = "Continuous";
    sheet.getRange("A18:G18").format.borders.bottom.style = "Continuous";
    sheet.getRange("A40:G40").format.borders.bottom.style = "Continuous";
    sheet.getRange("A47:G47").format.borders.bottom.style = "Continuous";
    
    // Highlight key outputs
    sheet.getRange("B54:B55").format.fill.color = "#E6F0FF";
    sheet.getRange("B54:B55").format.font.bold = true;
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
    createLBOModel
};