// merger.js - Merger Model Template

/**
 * Creates a Merger & Acquisition model in Excel
 * @param {object} context - The Excel context
 * @param {object} params - Parameters for the merger model
 * @returns {Promise<object>} - Information about the created model
 */
async function createMergerModel(context, params = {}) {
    // Default parameters
    const defaults = {
        acquirerName: "Acquirer Corp",
        targetName: "Target Corp",
        acquirerSharePrice: 50.0,
        targetSharePrice: 30.0,
        acquirerShares: 100, // in millions
        targetShares: 50, // in millions
        acquirerNetDebt: 500, // in millions
        targetNetDebt: 200, // in millions
        acquirerEPS: 3.50,
        targetEPS: 2.00,
        offerPremium: 0.30, // 30% premium
        cashConsideration: 0.40, // 40% cash, 60% stock
        synergies: 100, // in millions
        taxRate: 0.25,
        transactionFees: 50 // in millions
    };

    // Merge defaults with provided parameters
    const modelParams = { ...defaults, ...params };
    
    // Create a new worksheet for the model if it doesn't exist
    let sheet;
    try {
        sheet = context.workbook.worksheets.getItem("Merger Model");
    } catch (error) {
        sheet = context.workbook.worksheets.add("Merger Model");
    }
    
    // Activate the worksheet
    sheet.activate();
    
    // Setup the model structure
    await setupModelStructure(sheet, modelParams);
    
    // Create company information section
    await createCompanyInfoSection(sheet, modelParams);
    
    // Create transaction details section
    await createTransactionSection(sheet, modelParams);
    
    // Create pro forma analysis section
    await createProFormaSection(sheet, modelParams);
    
    // Format the worksheet
    await formatWorksheet(sheet, modelParams);
    
    return {
        sheetName: sheet.name,
        modelType: "Merger",
        parameters: modelParams
    };
}

/**
 * Sets up the basic structure of the merger model
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function setupModelStructure(sheet, params) {
    // Clear the sheet
    sheet.getRange().clear();
    
    // Set column widths
    sheet.getRange("A:A").columnWidth = 250; // Labels
    sheet.getRange("B:E").columnWidth = 120; // Data columns
    
    // Set title
    sheet.getRange("A1:E1").values = [[`${params.acquirerName} / ${params.targetName} - Merger Model`, "", "", "", ""]];
    sheet.getRange("A1:E1").format.font.bold = true;
    sheet.getRange("A1:E1").format.font.size = 16;
    sheet.getRange("A1:E1").merge();
}

/**
 * Creates the company information section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createCompanyInfoSection(sheet, params) {
    // Section headers
    sheet.getRange("A3").values = [["Company Information"]];
    sheet.getRange("A3").format.font.bold = true;
    sheet.getRange("A3").format.font.size = 14;
    
    // Column headers
    sheet.getRange("B4:D4").values = [["Acquirer", "Target", "Pro Forma"]];
    sheet.getRange("B4:D4").format.font.bold = true;
    
    // Company information labels
    sheet.getRange("A5:A14").values = [
        ["Share Price ($)"],
        ["Premium (%)"],
        ["Offer Price ($)"],
        ["Shares Outstanding (M)"],
        ["Market Capitalization ($M)"],
        ["Net Debt ($M)"],
        ["Enterprise Value ($M)"],
        ["EPS ($)"],
        ["P/E Ratio"],
        ["EV/EBITDA"]
    ];
    
    // Acquirer values
    sheet.getRange("B5").values = [[params.acquirerSharePrice]];
    sheet.getRange("B6").values = [["--"]];
    sheet.getRange("B7").values = [["--"]];
    sheet.getRange("B8").values = [[params.acquirerShares]];
    sheet.getRange("B9").formulas = [["=B5*B8"]];
    sheet.getRange("B10").values = [[params.acquirerNetDebt]];
    sheet.getRange("B11").formulas = [["=B9+B10"]];
    sheet.getRange("B12").values = [[params.acquirerEPS]];
    sheet.getRange("B13").formulas = [["=B5/B12"]];
    sheet.getRange("B14").values = [["--"]]; // Would need EBITDA input for EV/EBITDA
    
    // Target values
    sheet.getRange("C5").values = [[params.targetSharePrice]];
    sheet.getRange("C6").values = [[params.offerPremium]];
    sheet.getRange("C7").formulas = [["=C5*(1+C6)"]];
    sheet.getRange("C8").values = [[params.targetShares]];
    sheet.getRange("C9").formulas = [["=C5*C8"]];
    sheet.getRange("C10").values = [[params.targetNetDebt]];
    sheet.getRange("C11").formulas = [["=C9+C10"]];
    sheet.getRange("C12").values = [[params.targetEPS]];
    sheet.getRange("C13").formulas = [["=C5/C12"]];
    sheet.getRange("C14").values = [["--"]]; // Would need EBITDA input for EV/EBITDA
    
    // Pro forma values will be calculated in the pro forma section
    
    // Format cells
    sheet.getRange("B5:D5").numberFormat = "$0.00";
    sheet.getRange("B6:D6").numberFormat = "0.0%";
    sheet.getRange("B7:D7").numberFormat = "$0.00";
    sheet.getRange("B8:D8").numberFormat = "#,##0.0";
    sheet.getRange("B9:D11").numberFormat = "$#,##0.0";
    sheet.getRange("B12:D12").numberFormat = "$0.00";
    sheet.getRange("B13:D13").numberFormat = "0.0";
    sheet.getRange("B14:D14").numberFormat = "0.0";
}

/**
 * Creates the transaction details section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createTransactionSection(sheet, params) {
    // Section header
    sheet.getRange("A16").values = [["Transaction Details"]];
    sheet.getRange("A16").format.font.bold = true;
    sheet.getRange("A16").format.font.size = 14;
    
    // Transaction details labels
    sheet.getRange("A17:A26").values = [
        ["Offer Price per Share ($)"],
        ["Equity Purchase Price ($M)"],
        ["% Cash Consideration"],
        ["% Stock Consideration"],
        ["Cash Consideration ($M)"],
        ["Stock Consideration ($M)"],
        ["Exchange Ratio (Target/Acquirer)"],
        ["New Shares Issued (M)"],
        ["Transaction Fees ($M)"],
        ["Pro Forma Shares Outstanding (M)"]
    ];
    
    // Transaction values
    sheet.getRange("B17").formulas = [["=C7"]]; // Offer price per share
    sheet.getRange("B18").formulas = [["=C7*C8"]]; // Equity purchase price
    sheet.getRange("B19").values = [[params.cashConsideration]];
    sheet.getRange("B20").formulas = [["=1-B19"]];
    sheet.getRange("B21").formulas = [["=B18*B19"]]; // Cash consideration
    sheet.getRange("B22").formulas = [["=B18*B20"]]; // Stock consideration
    sheet.getRange("B23").formulas = [["=B22/(B5*C8)"]]; // Exchange ratio
    sheet.getRange("B24").formulas = [["=B22/B5"]]; // New shares issued
    sheet.getRange("B25").values = [[params.transactionFees]];
    sheet.getRange("B26").formulas = [["=B8+B24"]]; // Pro forma shares outstanding
    
    // Format cells
    sheet.getRange("B17").numberFormat = "$0.00";
    sheet.getRange("B18").numberFormat = "$#,##0.0";
    sheet.getRange("B19:B20").numberFormat = "0.0%";
    sheet.getRange("B21:B22").numberFormat = "$#,##0.0";
    sheet.getRange("B23").numberFormat = "0.000";
    sheet.getRange("B24").numberFormat = "#,##0.0";
    sheet.getRange("B25").numberFormat = "$#,##0.0";
    sheet.getRange("B26").numberFormat = "#,##0.0";
}

/**
 * Creates the pro forma analysis section
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function createProFormaSection(sheet, params) {
    // Section header
    sheet.getRange("A28").values = [["Pro Forma Analysis"]];
    sheet.getRange("A28").format.font.bold = true;
    sheet.getRange("A28").format.font.size = 14;
    
    // Pro forma analysis labels
    sheet.getRange("A29:A40").values = [
        ["Net Income - Acquirer ($M)"],
        ["Net Income - Target ($M)"],
        ["Synergies ($M)"],
        ["Tax Effect of Synergies ($M)"],
        ["After-Tax Synergies ($M)"],
        ["Incremental Interest Expense ($M)"],
        ["Tax Effect of Interest ($M)"],
        ["After-Tax Interest Expense ($M)"],
        ["Pro Forma Net Income ($M)"],
        ["Pro Forma EPS ($)"],
        ["Accretion / (Dilution) ($)"],
        ["Accretion / (Dilution) (%)"]
    ];
    
    // Pro forma calculations
    sheet.getRange("B29").formulas = [["=B12*B8"]]; // Acquirer net income
    sheet.getRange("B30").formulas = [["=C12*C8"]]; // Target net income
    sheet.getRange("B31").values = [[params.synergies]]; 
    sheet.getRange("B32").formulas = [["=B31*" + params.taxRate]];
    sheet.getRange("B33").formulas = [["=B31-B32"]]; // After-tax synergies
    
    // Assume incremental interest expense from new debt for cash portion
    sheet.getRange("B34").formulas = [["=B21*0.05"]]; // Assuming 5% interest rate
    sheet.getRange("B35").formulas = [["=B34*" + params.taxRate]];
    sheet.getRange("B36").formulas = [["=B34-B35"]]; // After-tax interest expense
    
    sheet.getRange("B37").formulas = [["=B29+B30+B33-B36"]]; // Pro forma net income
    sheet.getRange("B38").formulas = [["=B37/B26"]]; // Pro forma EPS
    sheet.getRange("B39").formulas = [["=B38-B12"]]; // Accretion/dilution per share
    sheet.getRange("B40").formulas = [["=B39/B12"]]; // Accretion/dilution percentage
    
    // Update pro forma values in company information section
    sheet.getRange("D9").formulas = [["=B9+C9"]]; // Pro forma market cap
    sheet.getRange("D10").formulas = [["=B10+C10+B21"]]; // Pro forma net debt
    sheet.getRange("D11").formulas = [["=D9+D10"]]; // Pro forma enterprise value
    sheet.getRange("D12").formulas = [["=B38"]]; // Pro forma EPS
    sheet.getRange("D13").formulas = [["=B5/D12"]]; // Pro forma P/E ratio
    
    // Format cells
    sheet.getRange("B29:B37").numberFormat = "$#,##0.0";
    sheet.getRange("B38").numberFormat = "$0.00";
    sheet.getRange("B39").numberFormat = "$0.00";
    sheet.getRange("B40").numberFormat = "0.0%";
}

/**
 * Formats the merger model worksheet
 * @param {object} sheet - The Excel worksheet
 * @param {object} params - Model parameters
 */
async function formatWorksheet(sheet, params) {
    // Add section borders
    sheet.getRange("A3:E3").format.borders.bottom.style = "Continuous";
    sheet.getRange("A16:E16").format.borders.bottom.style = "Continuous";
    sheet.getRange("A28:E28").format.borders.bottom.style = "Continuous";
    
    // Highlight key outputs
    sheet.getRange("B38:B40").format.fill.color = "#E6F0FF";
    sheet.getRange("B38:B40").format.font.bold = true;
    
    // Add conditional formatting for accretion/dilution
    const conditionalFormat = sheet.getRange("B40").conditionalFormats.add("CellValue");
    conditionalFormat.cellValue.format.font.color = "#107C10"; // Green for accretion
    conditionalFormat.cellValue.rule = { formula1: "0", operator: "GreaterThan" };
    
    const conditionalFormat2 = sheet.getRange("B40").conditionalFormats.add("CellValue");
    conditionalFormat2.cellValue.format.font.color = "#A4262C"; // Red for dilution
    conditionalFormat2.cellValue.rule = { formula1: "0", operator: "LessThan" };
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
    createMergerModel
};