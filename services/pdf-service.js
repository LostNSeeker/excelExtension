/**
 * PDF Processing Service
 * 
 * This service handles extraction of financial data from PDF documents
 * using text analysis and pattern recognition techniques.
 */

const pdfParse = require('pdf-parse');
const fs = require('fs');
const path = require('path');

class PDFService {
    constructor() {
        // Common patterns for financial data extraction
        this.patterns = {
            // Income Statement
            revenue: /(?:revenue|sales|net\s+sales)[\s:]*[$]?([\d,.]+)/i,
            costOfRevenue: /(?:cost\s+of\s+(?:revenue|sales|goods\s+sold)|cogs)[\s:]*[$]?([\d,.]+)/i,
            grossProfit: /(?:gross\s+profit|gross\s+margin)[\s:]*[$]?([\d,.]+)/i,
            operatingExpenses: /(?:operating\s+expenses|total\s+expenses)[\s:]*[$]?([\d,.]+)/i,
            operatingIncome: /(?:operating\s+income|operating\s+profit|ebit)[\s:]*[$]?([\d,.]+)/i,
            netIncome: /(?:net\s+income|net\s+earnings|net\s+profit)[\s:]*[$]?([\d,.]+)/i,
            
            // Balance Sheet
            totalAssets: /(?:total\s+assets)[\s:]*[$]?([\d,.]+)/i,
            totalLiabilities: /(?:total\s+liabilities)[\s:]*[$]?([\d,.]+)/i,
            totalEquity: /(?:(?:total\s+|total\s+shareholders['']?\s+)?equity|net\s+assets)[\s:]*[$]?([\d,.]+)/i,
            cash: /(?:cash(?:\s+and\s+cash\s+equivalents)?)[\s:]*[$]?([\d,.]+)/i,
            
            // Cash Flow
            operatingCashFlow: /(?:(?:net\s+)?cash\s+(?:provided\s+by|from)\s+operating\s+activities)[\s:]*[$]?([\d,.]+)/i,
            investingCashFlow: /(?:(?:net\s+)?cash\s+(?:used\s+in|provided\s+by|from)\s+investing\s+activities)[\s:]*[$]?([\d,.]+)/i,
            financingCashFlow: /(?:(?:net\s+)?cash\s+(?:used\s+in|provided\s+by|from)\s+financing\s+activities)[\s:]*[$]?([\d,.]+)/i,
            
            // Company and Period Information
            companyName: /([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*(?:\s+[A-Z]+)?(?:\s+Inc\.?|Corp\.?|Corporation|Company|Co\.?))/,
            period: /(?:(?:fiscal|year|period)[\s:]+((?:ended|ending)\s+)?(?:(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},?\s+)?\d{4})/i
        };
    }

    /**
     * Extract financial data from a PDF buffer
     * @param {Buffer} pdfBuffer - PDF file buffer
     * @returns {Promise<Object>} - Extracted financial data
     */
    async extractFinancialData(pdfBuffer) {
        try {
            // Parse PDF to text
            const pdfData = await pdfParse(pdfBuffer);
            const text = pdfData.text;
            
            // Extract data using patterns
            const extractedData = this.extractDataFromText(text);
            
            // Identify document type
            const documentType = this.identifyDocumentType(text);
            
            // Add metadata
            extractedData.metadata = {
                pageCount: pdfData.numpages,
                documentType,
                extractionTimestamp: new Date().toISOString()
            };
            
            return extractedData;
        } catch (error) {
            console.error('Error extracting data from PDF:', error);
            throw new Error('Failed to extract data from PDF');
        }
    }

    /**
     * Extract financial data from text using regular expressions
     * @param {string} text - Text content to search
     * @returns {Object} - Extracted financial data
     */
    extractDataFromText(text) {
        const data = {
            company: this.extractPattern(text, this.patterns.companyName),
            period: this.extractPattern(text, this.patterns.period),
            financial_data: {
                income_statement: {
                    revenue: this.extractNumberPattern(text, this.patterns.revenue),
                    cost_of_revenue: this.extractNumberPattern(text, this.patterns.costOfRevenue),
                    gross_profit: this.extractNumberPattern(text, this.patterns.grossProfit),
                    operating_expenses: this.extractNumberPattern(text, this.patterns.operatingExpenses),
                    operating_income: this.extractNumberPattern(text, this.patterns.operatingIncome),
                    net_income: this.extractNumberPattern(text, this.patterns.netIncome)
                },
                balance_sheet: {
                    total_assets: this.extractNumberPattern(text, this.patterns.totalAssets),
                    total_liabilities: this.extractNumberPattern(text, this.patterns.totalLiabilities),
                    total_equity: this.extractNumberPattern(text, this.patterns.totalEquity),
                    cash: this.extractNumberPattern(text, this.patterns.cash)
                },
                cash_flow: {
                    operating_cash_flow: this.extractNumberPattern(text, this.patterns.operatingCashFlow),
                    investing_cash_flow: this.extractNumberPattern(text, this.patterns.investingCashFlow),
                    financing_cash_flow: this.extractNumberPattern(text, this.patterns.financingCashFlow)
                }
            }
        };
        
        // Calculate derived metrics if possible
        this.calculateDerivedMetrics(data);
        
        return data;
    }

    /**
     * Extract pattern from text
     * @param {string} text - Text to search in
     * @param {RegExp} pattern - Regular expression pattern
     * @returns {string|null} - Matched string or null
     */
    extractPattern(text, pattern) {
        const match = text.match(pattern);
        return match ? match[1] : null;
    }

    /**
     * Extract and parse a number from text
     * @param {string} text - Text to search in
     * @param {RegExp} pattern - Regular expression pattern
     * @returns {number|null} - Parsed number or null
     */
    extractNumberPattern(text, pattern) {
        const match = text.match(pattern);
        
        if (!match) return null;
        
        // Remove commas and parse as float
        const numberString = match[1].replace(/,/g, '');
        return parseFloat(numberString);
    }

    /**
     * Calculate derived financial metrics from extracted data
     * @param {Object} data - Extracted financial data
     */
    calculateDerivedMetrics(data) {
        const income = data.financial_data.income_statement;
        const balance = data.financial_data.balance_sheet;
        
        // Calculate total equity if not found directly
        if (!balance.total_equity && balance.total_assets && balance.total_liabilities) {
            balance.total_equity = balance.total_assets - balance.total_liabilities;
        }
        
        // Calculate key ratios
        data.key_ratios = {};
        
        // Profitability ratios
        if (income.revenue && income.net_income) {
            data.key_ratios.profit_margin = income.net_income / income.revenue;
        }
        
        if (income.net_income && balance.total_assets) {
            data.key_ratios.return_on_assets = income.net_income / balance.total_assets;
        }
        
        if (income.net_income && balance.total_equity) {
            data.key_ratios.return_on_equity = income.net_income / balance.total_equity;
        }
        
        // Leverage ratio
        if (balance.total_liabilities && balance.total_equity) {
            data.key_ratios.debt_to_equity = balance.total_liabilities / balance.total_equity;
        }
        
        // Add additional derived metrics as needed
    }

    /**
     * Identify the type of financial document
     * @param {string} text - Document text content
     * @returns {string} - Document type
     */
    identifyDocumentType(text) {
        const lowerText = text.toLowerCase();
        
        // Count occurrences of key phrases
        const incomeCount = (lowerText.match(/income statement|statement of income|statement of operations|profit and loss/g) || []).length;
        const balanceCount = (lowerText.match(/balance sheet|statement of financial position/g) || []).length;
        const cashFlowCount = (lowerText.match(/cash flow|statement of cash flows/g) || []).length;
        const annualReportCount = (lowerText.match(/annual report|yearly report/g) || []).length;
        
        // Determine document type based on frequency
        if (annualReportCount > 0) {
            return 'annual_report';
        } else if (incomeCount > balanceCount && incomeCount > cashFlowCount) {
            return 'income_statement';
        } else if (balanceCount > incomeCount && balanceCount > cashFlowCount) {
            return 'balance_sheet';
        } else if (cashFlowCount > incomeCount && cashFlowCount > balanceCount) {
            return 'cash_flow_statement';
        } else if (incomeCount > 0 && balanceCount > 0 && cashFlowCount > 0) {
            return 'financial_statements';
        } else {
            return 'unknown';
        }
    }

    /**
     * Extract tables from PDF text
     * Note: This is a simplified approach. Real table extraction often requires
     * more sophisticated techniques using the PDF structure, positions, etc.
     * @param {string} text - PDF text content
     * @returns {Array} - Extracted tables as arrays of rows
     */
    extractTables(text) {
        const tables = [];
        const lines = text.split('\n');
        let currentTable = null;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // Skip empty lines
            if (!line) continue;
            
            // Check if this line might be part of a table
            // (simplified: look for multiple numbers in a line)
            const numberCount = (line.match(/[\d,.]+/g) || []).length;
            
            if (numberCount >= 3) {
                // This might be a table row
                if (!currentTable) {
                    currentTable = [];
                }
                
                // Split by whitespace and filter out empty entries
                const cells = line.split(/\s{2,}/).filter(cell => cell.trim());
                
                if (cells.length >= 2) {
                    currentTable.push(cells);
                }
            } else if (currentTable && currentTable.length > 1) {
                // End of a table
                tables.push(currentTable);
                currentTable = null;
            }
        }
        
        // Add the last table if there is one
        if (currentTable && currentTable.length > 1) {
            tables.push(currentTable);
        }
        
        return tables;
    }

    /**
     * Format extracted financial data for Excel import
     * @param {Object} data - Extracted financial data
     * @returns {Object} - Formatted data for Excel
     */
    formatForExcel(data) {
        const formattedData = {
            company: data.company || 'Unknown Company',
            period: data.period || 'Unknown Period',
            sheets: {}
        };
        
        // Income Statement
        const incomeStatementRows = [
            ['Income Statement'],
            ['', ''],
            ['Revenue', data.financial_data.income_statement.revenue],
            ['Cost of Revenue', data.financial_data.income_statement.cost_of_revenue],
            ['Gross Profit', data.financial_data.income_statement.gross_profit],
            ['Operating Expenses', data.financial_data.income_statement.operating_expenses],
            ['Operating Income', data.financial_data.income_statement.operating_income],
            ['Net Income', data.financial_data.income_statement.net_income]
        ];
        
        // Balance Sheet
        const balanceSheetRows = [
            ['Balance Sheet'],
            ['', ''],
            ['Cash', data.financial_data.balance_sheet.cash],
            ['Total Assets', data.financial_data.balance_sheet.total_assets],
            ['Total Liabilities', data.financial_data.balance_sheet.total_liabilities],
            ['Total Equity', data.financial_data.balance_sheet.total_equity]
        ];
        
        // Cash Flow
        const cashFlowRows = [
            ['Cash Flow Statement'],
            ['', ''],
            ['Operating Cash Flow', data.financial_data.cash_flow.operating_cash_flow],
            ['Investing Cash Flow', data.financial_data.cash_flow.investing_cash_flow],
            ['Financing Cash Flow', data.financial_data.cash_flow.financing_cash_flow],
            ['Net Change in Cash', 
                (data.financial_data.cash_flow.operating_cash_flow || 0) + 
                (data.financial_data.cash_flow.investing_cash_flow || 0) + 
                (data.financial_data.cash_flow.financing_cash_flow || 0)]
        ];
        
        // Key Ratios
        const keyRatiosRows = [
            ['Key Ratios'],
            ['', ''],
            ['Profit Margin', data.key_ratios.profit_margin],
            ['Return on Assets', data.key_ratios.return_on_assets],
            ['Return on Equity', data.key_ratios.return_on_equity],
            ['Debt to Equity', data.key_ratios.debt_to_equity]
        ];
        
        formattedData.sheets = {
            'Income Statement': incomeStatementRows,
            'Balance Sheet': balanceSheetRows,
            'Cash Flow': cashFlowRows,
            'Key Ratios': keyRatiosRows
        };
        
        return formattedData;
    }
}

module.exports = new PDFService();