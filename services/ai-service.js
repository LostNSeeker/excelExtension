/**
 * AI Service
 * 
 * This service handles integration with AI APIs for natural language processing,
 * providing methods for generating formulas, answering questions, and analyzing financial data.
 */

const { Configuration, OpenAIApi } = require('openai');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

class AIService {
    constructor() {
        // Initialize OpenAI configuration
        this.configuration = new Configuration({
            apiKey: process.env.OPENAI_API_KEY,
        });
        this.openai = new OpenAIApi(this.configuration);
        
        // Configuration options
        this.defaultModel = process.env.AI_MODEL || 'gpt-4';
        this.defaultTemperature = parseFloat(process.env.AI_TEMPERATURE || '0.7');
        this.defaultMaxTokens = parseInt(process.env.AI_MAX_TOKENS || '2000');
    }

    /**
     * Generate an Excel formula based on natural language description
     * @param {string} query - Natural language description of the formula
     * @returns {Promise<Object>} - Formula and explanation
     */
    async generateFormula(query) {
        try {
            const systemMessage = `You are a financial Excel formula expert. Provide Excel formulas and explanations for financial calculations. 
                Return your response in a format that clearly separates the formula and the explanation. 
                Start the formula with '=' and provide a detailed yet concise explanation.`;
            
            const completion = await this.openai.createChatCompletion({
                model: this.defaultModel,
                messages: [
                    { role: "system", content: systemMessage },
                    { role: "user", content: `Provide an Excel formula for: ${query}` }
                ],
                temperature: this.defaultTemperature,
                max_tokens: this.defaultMaxTokens
            });
            
            const response = completion.data.choices[0].message.content;
            
            // Parse the response to extract formula and explanation
            const formulaMatch = response.match(/=.*?(?=\n|$)/);
            const formula = formulaMatch ? formulaMatch[0] : "Could not generate formula";
            const explanation = response.replace(formula, '').trim();
            
            return {
                formula,
                explanation
            };
        } catch (error) {
            console.error('Error generating formula:', error);
            throw new Error('Failed to generate formula');
        }
    }

    /**
     * Get a response from the AI chatbot
     * @param {string} message - User message
     * @param {Array} history - Chat history 
     * @returns {Promise<Object>} - AI response
     */
    async getChatResponse(message, history = []) {
        try {
            const systemMessage = `You are a financial advisor and Excel expert helping with financial modeling and data analysis. 
                Provide practical, actionable advice focused on solving the user's financial modeling and Excel needs. 
                Keep responses concise yet informative, using financial terminology appropriately.`;
            
            // Convert chat history to OpenAI format
            const messages = [
                { role: "system", content: systemMessage },
                ...history.map(msg => ({
                    role: msg.role,
                    content: msg.content
                })),
                { role: "user", content: message }
            ];
            
            const completion = await this.openai.createChatCompletion({
                model: this.defaultModel,
                messages,
                temperature: this.defaultTemperature,
                max_tokens: this.defaultMaxTokens
            });
            
            const response = completion.data.choices[0].message.content;
            
            return {
                role: "assistant",
                content: response
            };
        } catch (error) {
            console.error('Error generating chat response:', error);
            throw new Error('Failed to generate chat response');
        }
    }

    /**
     * Analyze a formula for errors and provide corrections
     * @param {string} formula - Excel formula to analyze
     * @param {string} context - Additional context about the formula 
     * @returns {Promise<Object>} - Analysis results
     */
    async analyzeFormula(formula, context = '') {
        try {
            const systemMessage = `You are an Excel formula debugging expert. Analyze the following formula for errors and suggest corrections. 
                Provide a detailed explanation of any issues found and how to fix them.`;
            
            const userMessage = `Formula: ${formula}\n${context ? `Context: ${context}` : ''}`;
            
            const completion = await this.openai.createChatCompletion({
                model: this.defaultModel,
                messages: [
                    { role: "system", content: systemMessage },
                    { role: "user", content: userMessage }
                ],
                temperature: this.defaultTemperature,
                max_tokens: this.defaultMaxTokens
            });
            
            const response = completion.data.choices[0].message.content;
            
            // Parse the AI response to extract error analysis
            const hasError = /error|invalid|incorrect|issue|problem/i.test(response);
            const suggestedCorrection = response.match(/(?:suggested|corrected|fixed|proper)\s+formula[:\s]+([^]+?)(?:\n\n|$)/i);
            
            return {
                hasError,
                analysis: response,
                suggestedFormula: suggestedCorrection ? suggestedCorrection[1].trim() : null
            };
        } catch (error) {
            console.error('Error analyzing formula:', error);
            throw new Error('Failed to analyze formula');
        }
    }

    /**
     * Generate financial model assumptions based on industry and company description
     * @param {string} industry - Industry sector
     * @param {string} companyDescription - Description of the company
     * @returns {Promise<Object>} - Generated assumptions
     */
    async generateModelAssumptions(industry, companyDescription) {
        try {
            const systemMessage = `You are a financial modeling expert. Generate reasonable financial model assumptions for a company 
                based on its industry and description. Provide assumptions for revenue growth, margins, capex, working capital, and other 
                relevant metrics.`;
            
            const userMessage = `Industry: ${industry}\nCompany Description: ${companyDescription}\n\nProvide financial model assumptions for a 5-year projection.`;
            
            const completion = await this.openai.createChatCompletion({
                model: this.defaultModel,
                messages: [
                    { role: "system", content: systemMessage },
                    { role: "user", content: userMessage }
                ],
                temperature: this.defaultTemperature,
                max_tokens: this.defaultMaxTokens
            });
            
            const response = completion.data.choices[0].message.content;
            
            // Parse the AI response to extract structured assumptions
            const assumptions = this.parseAssumptions(response);
            
            return {
                rawResponse: response,
                structuredAssumptions: assumptions
            };
        } catch (error) {
            console.error('Error generating model assumptions:', error);
            throw new Error('Failed to generate model assumptions');
        }
    }

    /**
     * Parse assumptions from AI response text
     * @param {string} text - AI response text 
     * @returns {Object} - Structured assumptions
     */
    parseAssumptions(text) {
        const assumptions = {
            revenueGrowth: [],
            grossMargin: [],
            ebitdaMargin: [],
            taxRate: null,
            capexPercentOfRevenue: null,
            depreciationPercentOfRevenue: null,
            workingCapitalPercentOfRevenue: null,
            otherAssumptions: {}
        };
        
        // Extract revenue growth
        const revenueGrowthMatch = text.match(/revenue growth(?:\s+rate)?(?:\s+of)?\s*:\s*([\d.]+%[^]*?)(?:\n|$)/i);
        if (revenueGrowthMatch) {
            const growthText = revenueGrowthMatch[1];
            const percentages = growthText.match(/(\d+(?:\.\d+)?)\s*%/g);
            if (percentages) {
                assumptions.revenueGrowth = percentages.map(p => parseFloat(p) / 100);
            }
        }
        
        // Extract gross margin
        const grossMarginMatch = text.match(/gross margin(?:\s+of)?\s*:\s*([\d.]+%[^]*?)(?:\n|$)/i);
        if (grossMarginMatch) {
            const marginText = grossMarginMatch[1];
            const percentages = marginText.match(/(\d+(?:\.\d+)?)\s*%/g);
            if (percentages) {
                assumptions.grossMargin = percentages.map(p => parseFloat(p) / 100);
            }
        }
        
        // Extract EBITDA margin
        const ebitdaMarginMatch = text.match(/ebitda margin(?:\s+of)?\s*:\s*([\d.]+%[^]*?)(?:\n|$)/i);
        if (ebitdaMarginMatch) {
            const marginText = ebitdaMarginMatch[1];
            const percentages = marginText.match(/(\d+(?:\.\d+)?)\s*%/g);
            if (percentages) {
                assumptions.ebitdaMargin = percentages.map(p => parseFloat(p) / 100);
            }
        }
        
        // Extract tax rate
        const taxRateMatch = text.match(/tax rate(?:\s+of)?\s*:\s*(\d+(?:\.\d+)?)\s*%/i);
        if (taxRateMatch) {
            assumptions.taxRate = parseFloat(taxRateMatch[1]) / 100;
        }
        
        // Extract capex
        const capexMatch = text.match(/(?:capital expenditure|capex)(?:\s+as a percentage of revenue)?\s*:\s*(\d+(?:\.\d+)?)\s*%/i);
        if (capexMatch) {
            assumptions.capexPercentOfRevenue = parseFloat(capexMatch[1]) / 100;
        }
        
        // Extract depreciation
        const depreciationMatch = text.match(/depreciation(?:\s+as a percentage of revenue)?\s*:\s*(\d+(?:\.\d+)?)\s*%/i);
        if (depreciationMatch) {
            assumptions.depreciationPercentOfRevenue = parseFloat(depreciationMatch[1]) / 100;
        }
        
        // Extract working capital
        const wcMatch = text.match(/working capital(?:\s+as a percentage of revenue)?\s*:\s*(\d+(?:\.\d+)?)\s*%/i);
        if (wcMatch) {
            assumptions.workingCapitalPercentOfRevenue = parseFloat(wcMatch[1]) / 100;
        }
        
        return assumptions;
    }

    /**
     * Analyze financial ratio trends and provide insights
     * @param {Array} ratios - Array of financial ratio objects
     * @returns {Promise<Object>} - Analysis and insights
     */
    async analyzeFinancialRatios(ratios) {
        try {
            const systemMessage = `You are a financial analyst specializing in ratio analysis. Analyze the provided financial ratios 
                and provide insights on the company's performance, trends, strengths, weaknesses, and potential red flags.`;
            
            const userMessage = `Financial Ratios:\n${JSON.stringify(ratios, null, 2)}\n\nProvide a comprehensive analysis of these ratios.`;
            
            const completion = await this.openai.createChatCompletion({
                model: this.defaultModel,
                messages: [
                    { role: "system", content: systemMessage },
                    { role: "user", content: userMessage }
                ],
                temperature: this.defaultTemperature,
                max_tokens: this.defaultMaxTokens
            });
            
            return {
                analysis: completion.data.choices[0].message.content
            };
        } catch (error) {
            console.error('Error analyzing financial ratios:', error);
            throw new Error('Failed to analyze financial ratios');
        }
    }
}

module.exports = new AIService();