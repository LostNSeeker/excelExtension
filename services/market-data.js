/**
 * Market Data Service
 * 
 * This service handles integration with financial market data APIs
 * and provides methods for retrieving and processing financial data.
 */

const axios = require('axios');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

class MarketDataService {
    constructor() {
        this.apiKey = process.env.MARKET_DATA_API_KEY;
        this.baseUrl = 'https://api.example.com'; // Replace with actual API URL
    }

    /**
     * Get historical stock price data for a given symbol
     * @param {string} symbol - Stock symbol (e.g., 'AAPL')
     * @param {string} startDate - Start date in YYYY-MM-DD format
     * @param {string} endDate - End date in YYYY-MM-DD format
     * @returns {Promise<Object>} - Historical price data
     */
    async getHistoricalPrices(symbol, startDate, endDate) {
        try {
            const response = await axios.get(`${this.baseUrl}/stocks/historical`, {
                params: {
                    symbol,
                    startDate,
                    endDate,
                    apiKey: this.apiKey
                }
            });

            return response.data;
        } catch (error) {
            console.error(`Error fetching historical prices for ${symbol}:`, error);
            throw new Error(`Failed to retrieve historical prices for ${symbol}`);
        }
    }

    /**
     * Get financial statement data for a company
     * @param {string} symbol - Stock symbol (e.g., 'AAPL')
     * @param {string} statement - Statement type ('income', 'balance', 'cash')
     * @param {string} period - Period ('annual', 'quarterly')
     * @returns {Promise<Object>} - Financial statement data
     */
    async getFinancialStatement(symbol, statement, period) {
        try {
            const response = await axios.get(`${this.baseUrl}/stocks/financials`, {
                params: {
                    symbol,
                    statement,
                    period,
                    apiKey: this.apiKey
                }
            });

            return response.data;
        } catch (error) {
            console.error(`Error fetching ${statement} statement for ${symbol}:`, error);
            throw new Error(`Failed to retrieve ${statement} statement for ${symbol}`);
        }
    }

    /**
     * Get company key metrics
     * @param {string} symbol - Stock symbol (e.g., 'AAPL')
     * @returns {Promise<Object>} - Company metrics
     */
    async getCompanyMetrics(symbol) {
        try {
            const response = await axios.get(`${this.baseUrl}/stocks/metrics`, {
                params: {
                    symbol,
                    apiKey: this.apiKey
                }
            });

            return response.data;
        } catch (error) {
            console.error(`Error fetching metrics for ${symbol}:`, error);
            throw new Error(`Failed to retrieve metrics for ${symbol}`);
        }
    }

    /**
     * Search for companies by keyword
     * @param {string} query - Search query
     * @returns {Promise<Array>} - List of matching companies
     */
    async searchCompanies(query) {
        try {
            const response = await axios.get(`${this.baseUrl}/search`, {
                params: {
                    query,
                    apiKey: this.apiKey
                }
            });

            return response.data;
        } catch (error) {
            console.error(`Error searching for companies with query "${query}":`, error);
            throw new Error(`Failed to search for companies with query "${query}"`);
        }
    }

    /**
     * Get current market indices data
     * @returns {Promise<Object>} - Market indices data
     */
    async getMarketIndices() {
        try {
            const response = await axios.get(`${this.baseUrl}/market/indices`, {
                params: {
                    apiKey: this.apiKey
                }
            });

            return response.data;
        } catch (error) {
            console.error('Error fetching market indices:', error);
            throw new Error('Failed to retrieve market indices');
        }
    }

    /**
     * Calculate financial ratios from raw financial data
     * @param {Object} financialData - Financial statement data
     * @returns {Object} - Calculated financial ratios
     */
    calculateFinancialRatios(financialData) {
        const ratios = {};

        // Profitability ratios
        if (financialData.income && financialData.balance) {
            // Return on Assets (ROA)
            ratios.roa = financialData.income.netIncome / financialData.balance.totalAssets;
            
            // Return on Equity (ROE)
            ratios.roe = financialData.income.netIncome / financialData.balance.totalEquity;
            
            // Gross Margin
            ratios.grossMargin = financialData.income.grossProfit / financialData.income.revenue;
            
            // Operating Margin
            ratios.operatingMargin = financialData.income.operatingIncome / financialData.income.revenue;
            
            // Net Profit Margin
            ratios.netProfitMargin = financialData.income.netIncome / financialData.income.revenue;
        }
        
        // Liquidity ratios
        if (financialData.balance) {
            // Current Ratio
            ratios.currentRatio = financialData.balance.currentAssets / financialData.balance.currentLiabilities;
            
            // Quick Ratio
            ratios.quickRatio = (financialData.balance.currentAssets - financialData.balance.inventory) / financialData.balance.currentLiabilities;
        }
        
        // Solvency ratios
        if (financialData.balance) {
            // Debt to Equity Ratio
            ratios.debtToEquity = financialData.balance.totalDebt / financialData.balance.totalEquity;
            
            // Debt to Assets Ratio
            ratios.debtToAssets = financialData.balance.totalDebt / financialData.balance.totalAssets;
        }
        
        // Efficiency ratios
        if (financialData.income && financialData.balance) {
            // Asset Turnover Ratio
            ratios.assetTurnover = financialData.income.revenue / financialData.balance.totalAssets;
            
            // Inventory Turnover Ratio
            if (financialData.income.costOfRevenue && financialData.balance.inventory) {
                ratios.inventoryTurnover = financialData.income.costOfRevenue / financialData.balance.inventory;
            }
        }

        return ratios;
    }

    /**
     * Format financial data for use in Excel
     * @param {Object} data - Raw financial data
     * @returns {Array} - Formatted data for Excel
     */
    formatDataForExcel(data) {
        // This method would transform API data into a format suitable for Excel
        // The exact implementation would depend on the specific data structure
        // and the desired Excel output format
        
        const formattedData = [];
        
        // Example implementation for stock price data
        if (data && data.prices) {
            data.prices.forEach(price => {
                formattedData.push([
                    new Date(price.date),
                    price.open,
                    price.high,
                    price.low,
                    price.close,
                    price.volume
                ]);
            });
        }
        
        return formattedData;
    }
}

module.exports = new MarketDataService();