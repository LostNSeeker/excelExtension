// Wait for Office to initialize
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Initialize the application
        initApp();
    }
});

// API endpoint (in a real implementation, this would point to your actual API)
const API_BASE_URL = 'https://api.example.com';

function initApp() {
    // Setup tab navigation
    setupTabs();
    
    // Initialize formula assistant
    document.getElementById('get-formula').addEventListener('click', handleFormulaRequest);
    
    // Initialize chat functionality
    document.getElementById('send-chat').addEventListener('click', handleChatMessage);
    document.getElementById('chat-input').addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            handleChatMessage();
        }
    });
    
    // Initialize financial model buttons
    setupModelButtons();
    
    // Initialize PDF extraction
    document.getElementById('extract-pdf').addEventListener('click', handlePdfExtraction);
}

function setupTabs() {
    const tabs = document.querySelectorAll('.ms-Pivot-link');
    const contents = document.querySelectorAll('.ms-Pivot-content');
    
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // Remove selected class from all tabs
            tabs.forEach(t => t.classList.remove('ms-Pivot-link--selected'));
            // Add selected class to clicked tab
            tab.classList.add('ms-Pivot-link--selected');
            
            // Hide all content sections
            contents.forEach(content => content.classList.add('hidden'));
            
            // Show the corresponding content section
            const tabId = tab.id.split('-')[1];
            document.getElementById(`${tabId}-content`).classList.remove('hidden');
        });
    });
}

function setupModelButtons() {
    const modelButtons = document.querySelectorAll('.model-card button');
    modelButtons.forEach(button => {
        button.addEventListener('click', function() {
            const modelType = this.closest('.model-card').dataset.model;
            createFinancialModel(modelType);
        });
    });
}

// Formula Assistant
async function handleFormulaRequest() {
    const query = document.getElementById('formula-query').value;
    if (!query.trim()) return;
    
    const responseDiv = document.getElementById('formula-response');
    responseDiv.innerHTML = '<p>Thinking...</p>';
    
    try {
        // In a real implementation, this would call an AI service
        const formulaResponse = await getFormulaFromAI(query);
        
        responseDiv.innerHTML = `
            <h3>Suggested Formula:</h3>
            <pre>${formulaResponse.formula}</pre>
            <h3>Explanation:</h3>
            <p>${formulaResponse.explanation}</p>
            <button id="insert-formula" class="ms-Button">
                <span class="ms-Button-label">Insert Formula</span>
            </button>
        `;
        
        document.getElementById('insert-formula').addEventListener('click', () => {
            insertFormulaIntoCell(formulaResponse.formula);
        });
    } catch (error) {
        responseDiv.innerHTML = `<p class="error">Error: ${error.message}</p>`;
    }
}

// This is a mock function - in a real implementation, it would call an AI API
async function getFormulaFromAI(query) {
    // Simulate API call delay
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // Mock responses based on query keywords
    if (query.toLowerCase().includes('npv')) {
        return {
            formula: '=NPV(rate, value1, [value2], ...)',
            explanation: 'The NPV function calculates the net present value of an investment using a discount rate and a series of future payments (negative values) and income (positive values).'
        };
    } else if (query.toLowerCase().includes('irr')) {
        return {
            formula: '=IRR(values, [guess])',
            explanation: 'The IRR function returns the internal rate of return for a series of cash flows represented by the numbers in values.'
        };
    } else if (query.toLowerCase().includes('wacc') || query.toLowerCase().includes('weighted average cost of capital')) {
        return {
            formula: '=(E/(D+E))*Re + (D/(D+E))*Rd*(1-T)',
            explanation: 'The Weighted Average Cost of Capital (WACC) formula weighs a company\'s cost of debt and equity according to their proportions in the company\'s capital structure. Where E is equity value, D is debt value, Re is cost of equity, Rd is cost of debt, and T is tax rate.'
        };
    } else if (query.toLowerCase().includes('cagr')) {
        return {
            formula: '=(Ending Value/Beginning Value)^(1/Number of Years)-1',
            explanation: 'The Compound Annual Growth Rate (CAGR) measures the annual growth rate of an investment over a specified period of time longer than one year.'
        };
    } else {
        return {
            formula: '=SUM(number1, [number2], ...)',
            explanation: 'The SUM function adds all the numbers that you specify as arguments. Each argument can be a range, a cell reference, an array, a constant, a formula, or the result from another function.'
        };
    }
}

async function insertFormulaIntoCell(formula) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.formulas = [[formula]];
            await context.sync();
        });
    } catch (error) {
        console.error("Error inserting formula:", error);
        // Show error message to user
        document.getElementById('formula-response').innerHTML += `
            <p class="error">Error inserting formula: ${error.message}</p>
        `;
    }
}

// Chat Assistant
let chatHistory = [];

async function handleChatMessage() {
    const chatInput = document.getElementById('chat-input');
    const chatMessages = document.getElementById('chat-messages');
    const message = chatInput.value;
    
    if (!message.trim()) return;
    
    // Add user message to chat
    chatMessages.innerHTML += `
        <div class="chat-message user">
            <p>${message}</p>
        </div>
    `;
    
    // Clear input
    chatInput.value = '';
    
    // Scroll to bottom of chat
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    // Add user message to history
    chatHistory.push({ role: "user", content: message });
    
    try {
        // In a real implementation, this would call an AI service
        const response = await getChatResponseFromAI(message);
        
        // Add AI response to chat
        chatMessages.innerHTML += `
            <div class="chat-message assistant">
                <p>${response}</p>
            </div>
        `;
        
        // Add assistant response to history
        chatHistory.push({ role: "assistant", content: response });
        
        // Scroll to bottom of chat
        chatMessages.scrollTop = chatMessages.scrollHeight;
    } catch (error) {
        chatMessages.innerHTML += `
            <div class="chat-message assistant">
                <p class="error">Error: ${error.message}</p>
            </div>
        `;
    }
}

// This is a mock function - in a real implementation, it would call an AI API
async function getChatResponseFromAI(message) {
    // Simulate API call delay
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // Mock responses based on message keywords
    const lowerMessage = message.toLowerCase();
    
    if (lowerMessage.includes('hello') || lowerMessage.includes('hi')) {
        return "Hello! I'm your AI financial assistant. How can I help you with your financial modeling or Excel tasks today?";
    } else if (lowerMessage.includes('wacc') || lowerMessage.includes('weighted average cost of capital')) {
        return "The Weighted Average Cost of Capital (WACC) represents a firm's average cost of capital from all sources, including common stock, preferred stock, bonds, and other long-term debt. To calculate it, use this formula: WACC = (E/V × Re) + (D/V × Rd × (1 - T)), where E is equity value, D is debt value, V is total value (E+D), Re is cost of equity, Rd is cost of debt, and T is tax rate. Would you like me to help you implement this in your spreadsheet?";
    } else if (lowerMessage.includes('pivot table') || lowerMessage.includes('pivottable')) {
        return "To create a pivot table, select your data range, go to Insert > PivotTable. For financial analysis, you might want to group by time periods, aggregate by sum or average, and add calculated fields for ratios. Would you like me to help you create one for your current dataset?";
    } else if (lowerMessage.includes('forecast') || lowerMessage.includes('predict')) {
        return "For forecasting in Excel, you can use functions like FORECAST.LINEAR for linear trends, FORECAST.ETS for time series with seasonality, or create a regression model for multiple factors. Excel also offers the Forecast Sheet tool for visual analysis. Which approach would best suit your data?";
    } else if (lowerMessage.includes('dcf') || lowerMessage.includes('discounted cash flow')) {
        return "A Discounted Cash Flow (DCF) model values a company based on its expected future cash flows, discounted to present value. Key components include projecting cash flows, determining terminal value, and calculating the appropriate discount rate (usually WACC). Would you like me to help you create a DCF model template?";
    } else {
        return "I understand you're asking about " + message + ". To provide you with the most accurate assistance, could you please provide more details about your specific financial analysis needs?";
    }
}

// Financial Model Creator
async function createFinancialModel(modelType) {
    try {
        await Excel.run(async (context) => {
            // Add a new worksheet
            const sheet = context.workbook.worksheets.add(modelType + " Model");
            
            // Set up basic structure based on model type
            if (modelType === "dcf") {
                setupDCFModel(sheet);
            } else if (modelType === "lbo") {
                setupLBOModel(sheet);
            } else if (modelType === "merger") {
                setupMergerModel(sheet);
            } else {
                setupCustomModel(sheet);
            }
            
            sheet.activate();
            await context.sync();
            
            // Notify user
            alert(`${capitalizeFirstLetter(modelType)} model has been created in a new worksheet.`);
        });
    } catch (error) {
        console.error("Error creating model:", error);
        alert(`Error creating model: ${error.message}`);
    }
}

function setupDCFModel(sheet) {
    // Set up headers
    sheet.getRange("A1:J1").values = [["Discounted Cash Flow Model", "", "", "", "", "", "", "", "", ""]];
    sheet.getRange("A1:J1").format.font.bold = true;
    sheet.getRange("A1:J1").merge();
    
    // Historical and Forecast years
    sheet.getRange("B3:F3").values = [["Historical", "", "", "Forecast", ""]];
    sheet.getRange("B4:F4").values = [["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]];
    
    // Income Statement
    sheet.getRange("A6").values = [["Income Statement"]];
    sheet.getRange("A6").format.font.bold = true;
    
    sheet.getRange("A7:A16").values = [
        ["Revenue"],
        ["Growth %"],
        ["COGS"],
        ["Gross Profit"],
        ["Operating Expenses"],
        ["EBITDA"],
        ["Depreciation & Amortization"],
        ["EBIT"],
        ["Taxes"],
        ["Net Income"]
    ];
    
    // DCF Section
    sheet.getRange("A18").values = [["DCF Valuation"]];
    sheet.getRange("A18").format.font.bold = true;
    
    sheet.getRange("A19:A26").values = [
        ["EBIT"],
        ["Taxes"],
        ["NOPAT"],
        ["Add: D&A"],
        ["Less: CapEx"],
        ["Less: Change in WC"],
        ["Free Cash Flow"],
        ["Terminal Value"]
    ];
    
    sheet.getRange("A28:A32").values = [
        ["Discount Rate (WACC)"],
        ["Discounted FCF"],
        ["Sum of Discounted FCF"],
        ["Terminal Value"],
        ["Enterprise Value"]
    ];
    
    // Format the sheet
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
}

function setupLBOModel(sheet) {
    // Set up headers
    sheet.getRange("A1:J1").values = [["Leveraged Buyout (LBO) Model", "", "", "", "", "", "", "", "", ""]];
    sheet.getRange("A1:J1").format.font.bold = true;
    sheet.getRange("A1:J1").merge();
    
    // Years
    sheet.getRange("B3:F3").values = [["Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]];
    
    // Transaction Assumptions
    sheet.getRange("A5").values = [["Transaction Assumptions"]];
    sheet.getRange("A5").format.font.bold = true;
    
    sheet.getRange("A6:A10").values = [
        ["Purchase Price"],
        ["Debt"],
        ["Equity"],
        ["Debt / EBITDA"],
        ["Transaction Fees"]
    ];
    
    // Financial Projections
    sheet.getRange("A12").values = [["Financial Projections"]];
    sheet.getRange("A12").format.font.bold = true;
    
    sheet.getRange("A13:A20").values = [
        ["Revenue"],
        ["EBITDA"],
        ["EBITDA Margin"],
        ["Depreciation & Amortization"],
        ["EBIT"],
        ["Interest Expense"],
        ["EBT"],
        ["Net Income"]
    ];
    
    // Debt Schedule
    sheet.getRange("A22").values = [["Debt Schedule"]];
    sheet.getRange("A22").format.font.bold = true;
    
    sheet.getRange("A23:A26").values = [
        ["Beginning Balance"],
        ["Repayments"],
        ["New Issuances"],
        ["Ending Balance"]
    ];
    
    // Returns Analysis
    sheet.getRange("A28").values = [["Returns Analysis"]];
    sheet.getRange("A28").format.font.bold = true;
    
    sheet.getRange("A29:A33").values = [
        ["Exit Enterprise Value"],
        ["Exit Multiple"],
        ["Equity Value"],
        ["Initial Equity Investment"],
        ["Multiple of Money (MoM)"]
    ];
    
    // Format the sheet
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
}

function setupMergerModel(sheet) {
    // Set up headers
    sheet.getRange("A1:J1").values = [["Merger Model", "", "", "", "", "", "", "", "", ""]];
    sheet.getRange("A1:J1").format.font.bold = true;
    sheet.getRange("A1:J1").merge();
    
    // Company sections
    sheet.getRange("A3").values = [["Acquirer Information"]];
    sheet.getRange("A3").format.font.bold = true;
    
    sheet.getRange("A4:A8").values = [
        ["Share Price"],
        ["Shares Outstanding"],
        ["Market Capitalization"],
        ["Net Debt"],
        ["Enterprise Value"]
    ];
    
    sheet.getRange("A10").values = [["Target Information"]];
    sheet.getRange("A10").format.font.bold = true;
    
    sheet.getRange("A11:A15").values = [
        ["Share Price"],
        ["Shares Outstanding"],
        ["Market Capitalization"],
        ["Net Debt"],
        ["Enterprise Value"]
    ];
    
    // Transaction Details
    sheet.getRange("A17").values = [["Transaction Details"]];
    sheet.getRange("A17").format.font.bold = true;
    
    sheet.getRange("A18:A22").values = [
        ["Offer Price per Share"],
        ["Premium"],
        ["% Cash Consideration"],
        ["% Stock Consideration"],
        ["Transaction Value"]
    ];
    
    // Pro Forma Analysis
    sheet.getRange("A24").values = [["Pro Forma Analysis"]];
    sheet.getRange("A24").format.font.bold = true;
    
    sheet.getRange("A25:A31").values = [
        ["Acquirer EPS"],
        ["Target EPS"],
        ["Pro Forma EPS"],
        ["Accretion / (Dilution)"],
        ["Synergies Required for Breakeven"],
        ["Pro Forma Ownership - Acquirer"],
        ["Pro Forma Ownership - Target"]
    ];
    
    // Format the sheet
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
}

function setupCustomModel(sheet) {
    // Set up headers
    sheet.getRange("A1:J1").values = [["Custom Financial Model", "", "", "", "", "", "", "", "", ""]];
    sheet.getRange("A1:J1").format.font.bold = true;
    sheet.getRange("A1:J1").merge();
    
    // Provide guidance
    sheet.getRange("A3:F3").values = [["Use this template to create your custom financial model with AI assistance.", "", "", "", "", ""]];
    sheet.getRange("A3:F3").merge();
    
    // Basic structure
    sheet.getRange("A5").values = [["Model Assumptions"]];
    sheet.getRange("A5").format.font.bold = true;
    
    sheet.getRange("A6:A10").values = [
        ["Start Date"],
        ["Projection Period (Years)"],
        ["Growth Rate (%)"],
        ["Discount Rate (%)"],
        ["Tax Rate (%)"]
    ];
    
    sheet.getRange("A12").values = [["Projections"]];
    sheet.getRange("A12").format.font.bold = true;
    
    // Year headers
    sheet.getRange("B13:F13").values = [["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]];
    
    // Basic projection categories
    sheet.getRange("A14:A20").values = [
        ["Revenue"],
        ["Expenses"],
        ["EBITDA"],
        ["Capital Expenditures"],
        ["Changes in Working Capital"],
        ["Taxes"],
        ["Free Cash Flow"]
    ];
    
    // Format the sheet
    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();
}

// PDF Data Extraction
async function handlePdfExtraction() {
    const fileInput = document.getElementById('pdf-file');
    const resultDiv = document.getElementById('pdf-result');
    
    if (!fileInput.files || fileInput.files.length === 0) {
        resultDiv.innerHTML = '<p class="error">Please select a PDF file.</p>';
        return;
    }
    
    const file = fileInput.files[0];
    resultDiv.innerHTML = '<p>Processing PDF...</p>';
    
    try {
        // In a real implementation, this would call a service to extract data from the PDF
        // For the MVP, we'll simulate this with a mock response
        const extractedData = await extractDataFromPDF(file);
        
        resultDiv.innerHTML = `
            <h3>Extracted Data:</h3>
            <pre>${JSON.stringify(extractedData, null, 2)}</pre>
            <button id="import-data" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Import to Excel</span>
            </button>
        `;
        
        document.getElementById('import-data').addEventListener('click', () => {
            importDataToExcel(extractedData);
        });
    } catch (error) {
        resultDiv.innerHTML = `<p class="error">Error: ${error.message}</p>`;
    }
}

// This is a mock function - in a real implementation, it would use a PDF parsing service
async function extractDataFromPDF(file) {
    // Simulate API call delay
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Mock extracted data
    return {
        "company": "Sample Corp.",
        "period": "FY 2023",
        "financial_data": {
            "income_statement": {
                "revenue": 1250000,
                "cost_of_revenue": 750000,
                "gross_profit": 500000,
                "operating_expenses": 300000,
                "operating_income": 200000,
                "net_income": 150000
            },
            "balance_sheet": {
                "total_assets": 2000000,
                "total_liabilities": 800000,
                "total_equity": 1200000
            },
            "cash_flow": {
                "operating_cash_flow": 180000,
                "investing_cash_flow": -120000,
                "financing_cash_flow": -30000,
                "net_change_in_cash": 30000
            }
        },
        "key_ratios": {
            "profit_margin": 0.12,
            "return_on_assets": 0.075,
            "debt_to_equity": 0.667,
            "current_ratio": 1.8
        }
    };
}

async function importDataToExcel(data) {
    try {
        await Excel.run(async (context) => {
            // Add a new worksheet
            const sheet = context.workbook.worksheets.add("Imported Financial Data");
            
            // Set up company and period info
            sheet.getRange("A1:B1").values = [["Company:", data.company]];
            sheet.getRange("A2:B2").values = [["Period:", data.period]];
            
            // Income Statement
            sheet.getRange("A4").values = [["Income Statement"]];
            sheet.getRange("A4").format.font.bold = true;
            
            sheet.getRange("A5:B5").values = [["Revenue", data.financial_data.income_statement.revenue]];
            sheet.getRange("A6:B6").values = [["Cost of Revenue", data.financial_data.income_statement.cost_of_revenue]];
            sheet.getRange("A7:B7").values = [["Gross Profit", data.financial_data.income_statement.gross_profit]];
            sheet.getRange("A8:B8").values = [["Operating Expenses", data.financial_data.income_statement.operating_expenses]];
            sheet.getRange("A9:B9").values = [["Operating Income", data.financial_data.income_statement.operating_income]];
            sheet.getRange("A10:B10").values = [["Net Income", data.financial_data.income_statement.net_income]];
            
            // Balance Sheet
            sheet.getRange("A12").values = [["Balance Sheet"]];
            sheet.getRange("A12").format.font.bold = true;
            
            sheet.getRange("A13:B13").values = [["Total Assets", data.financial_data.balance_sheet.total_assets]];
            sheet.getRange("A14:B14").values = [["Total Liabilities", data.financial_data.balance_sheet.total_liabilities]];
            sheet.getRange("A15:B15").values = [["Total Equity", data.financial_data.balance_sheet.total_equity]];
            
            // Cash Flow
            sheet.getRange("A17").values = [["Cash Flow"]];
            sheet.getRange("A17").format.font.bold = true;
            
            sheet.getRange("A18:B18").values = [["Operating Cash Flow", data.financial_data.cash_flow.operating_cash_flow]];
            sheet.getRange("A19:B19").values = [["Investing Cash Flow", data.financial_data.cash_flow.investing_cash_flow]];
            sheet.getRange("A20:B20").values = [["Financing Cash Flow", data.financial_data.cash_flow.financing_cash_flow]];
            sheet.getRange("A21:B21").values = [["Net Change in Cash", data.financial_data.cash_flow.net_change_in_cash]];
            
            // Key Ratios
            sheet.getRange("A23").values = [["Key Ratios"]];
            sheet.getRange("A23").format.font.bold = true;
            
            sheet.getRange("A24:B24").values = [["Profit Margin", data.key_ratios.profit_margin]];
            sheet.getRange("A25:B25").values = [["Return on Assets", data.key_ratios.return_on_assets]];
            sheet.getRange("A26:B26").values = [["Debt to Equity", data.key_ratios.debt_to_equity]];
            sheet.getRange("A27:B27").values = [["Current Ratio", data.key_ratios.current_ratio]];
            
            // Format numbers
            sheet.getRange("B5:B10").numberFormat = "#,##0.00";
            sheet.getRange("B13:B15").numberFormat = "#,##0.00";
            sheet.getRange("B18:B21").numberFormat = "#,##0.00";
            sheet.getRange("B24:B27").numberFormat = "0.00";
            
            // Format the sheet
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();
            
            sheet.activate();
            await context.sync();
            
            // Notify user
            alert("Financial data has been imported to a new worksheet.");
        });
    } catch (error) {
        console.error("Error importing data:", error);
        alert(`Error importing data: ${error.message}`);
    }
}

// Utility function
function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}
