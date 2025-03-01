// server.js
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { Configuration, OpenAIApi } = require('openai');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const path = require('path');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configure OpenAI (replace with your preferred AI service)
const configuration = new Configuration({
    apiKey: process.env.OPENAI_API_KEY,
});
const openai = new OpenAIApi(configuration);

// Formula Assistant Endpoint
app.post('/api/formula', async (req, res) => {
    try {
        const { query } = req.body;
        
        if (!query) {
            return res.status(400).json({ error: 'Query is required' });
        }
        
        const completion = await openai.createChatCompletion({
            model: "gpt-4", // Use appropriate model
            messages: [
                { 
                    role: "system", 
                    content: "You are a financial Excel formula expert. Provide Excel formulas and explanations for financial calculations. " +
                            "Return your response in a format that clearly separates the formula and the explanation. " +
                            "Start the formula with '=' and provide a detailed yet concise explanation." 
                },
                { role: "user", content: `Provide an Excel formula for: ${query}` }
            ],
            temperature: 0.7,
        });
        
        const response = completion.data.choices[0].message.content;
        
        // Parse the response to extract formula and explanation
        const formulaMatch = response.match(/=.*?(?=\n|$)/);
        const formula = formulaMatch ? formulaMatch[0] : "Could not generate formula";
        const explanation = response.replace(formula, '').trim();
        
        res.json({
            formula,
            explanation
        });
    } catch (error) {
        console.error('Error generating formula:', error);
        res.status(500).json({ error: 'Failed to generate formula' });
    }
});

// Chat Assistant Endpoint
app.post('/api/chat', async (req, res) => {
    try {
        const { message, history } = req.body;
        
        if (!message) {
            return res.status(400).json({ error: 'Message is required' });
        }
        
        // Convert chat history to OpenAI format
        const messages = [
            { 
                role: "system", 
                content: "You are a financial advisor and Excel expert helping with financial modeling and data analysis. " +
                         "Provide practical, actionable advice focused on solving the user's financial modeling and Excel needs. " +
                         "Keep responses concise yet informative, using financial terminology appropriately." 
            },
            ...history.map(msg => ({
                role: msg.role,
                content: msg.content
            })),
            { role: "user", content: message }
        ];
        
        const completion = await openai.createChatCompletion({
            model: "gpt-4", // Use appropriate model
            messages,
            temperature: 0.7,
        });
        
        const response = completion.data.choices[0].message.content;
        
        res.json({
            role: "assistant",
            content: response
        });
    } catch (error) {
        console.error('Error generating chat response:', error);
        res.status(500).json({ error: 'Failed to generate chat response' });
    }
});

// PDF Extraction Endpoint
app.post('/api/extract-pdf', upload.single('pdf'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'PDF file is required' });
        }

        // Parse PDF
        const pdfData = await pdfParse(req.file.buffer);
        
        // In a real implementation, this would use NLP/ML to extract structured financial data
        // For MVP purposes, we're using a simple text parsing approach
        const text = pdfData.text;
        
        // Extract financial information using regex patterns
        // This is a simplified approach - a real implementation would use more sophisticated NLP
        const data = extractFinancialData(text);
        
        res.json({
            success: true,
            pageCount: pdfData.numpages,
            data
        });
    } catch (error) {
        console.error('Error extracting PDF data:', error);
        res.status(500).json({ error: 'Failed to extract PDF data' });
    }
});

// Simple financial data extraction from text
function extractFinancialData(text) {
    // This is a placeholder for the actual extraction logic
    // In a real implementation, this would use more sophisticated NLP techniques
    
    // Extract company name (simplified approach)
    const companyMatch = text.match(/([A-Z][a-z]+ [A-Z][a-z]+|[A-Z][A-Z]+) (Inc\.|Corp\.|Corporation|Company|Co\.)/);
    const company = companyMatch ? companyMatch[0] : "Unknown Company";
    
    // Extract year/period (simplified approach)
    const yearMatch = text.match(/(FY|Fiscal Year|Year End(ed|ing)?) (\d{4})/i);
    const period = yearMatch ? yearMatch[0] : "Unknown Period";
    
    // Extract financial figures (very simplified approach)
    const revenueMatch = text.match(/Revenue[:\s]*[$]?(\d{1,3}(,\d{3})*(\.\d+)?)/i);
    const netIncomeMatch = text.match(/Net Income[:\s]*[$]?(\d{1,3}(,\d{3})*(\.\d+)?)/i);
    const totalAssetsMatch = text.match(/Total Assets[:\s]*[$]?(\d{1,3}(,\d{3})*(\.\d+)?)/i);
    const totalLiabilitiesMatch = text.match(/Total Liabilities[:\s]*[$]?(\d{1,3}(,\d{3})*(\.\d+)?)/i);
    
    // Convert matched strings to numbers (removing commas and $ signs)
    const parseAmount = (match) => {
        if (!match) return null;
        return parseFloat(match[1].replace(/,/g, ''));
    };
    
    const revenue = parseAmount(revenueMatch);
    const netIncome = parseAmount(netIncomeMatch);
    const totalAssets = parseAmount(totalAssetsMatch);
    const totalLiabilities = parseAmount(totalLiabilitiesMatch);
    
    // Calculate some derived values
    const totalEquity = totalAssets && totalLiabilities ? totalAssets - totalLiabilities : null;
    const profitMargin = revenue && netIncome ? netIncome / revenue : null;
    
    return {
        company,
        period,
        financial_data: {
            income_statement: {
                revenue: revenue || 0,
                net_income: netIncome || 0
            },
            balance_sheet: {
                total_assets: totalAssets || 0,
                total_liabilities: totalLiabilities || 0,
                total_equity: totalEquity || 0
            }
        },
        key_ratios: {
            profit_margin: profitMargin || 0
        }
    };
}

// Financial Model Templates Endpoint
app.get('/api/model-templates/:type', (req, res) => {
    const { type } = req.params;
    
    // In a real implementation, this would fetch template data from a database
    // For the MVP, we'll return mock template data
    const templates = {
        dcf: {
            name: "Discounted Cash Flow Model",
            sections: [
                "Assumptions",
                "Income Statement Projections",
                "Cash Flow Projections",
                "Discount Rate Calculation",
                "Terminal Value",
                "Valuation Summary"
            ],
            // Template data would be more detailed in a real implementation
        },
        lbo: {
            name: "Leveraged Buyout Model",
            sections: [
                "Transaction Assumptions",
                "Purchase Price Calculation",
                "Debt and Financing Structure",
                "Income Statement Projections",
                "Debt Schedule",
                "Returns Analysis"
            ],
        },
        merger: {
            name: "Merger Model",
            sections: [
                "Acquirer Information",
                "Target Information",
                "Transaction Details",
                "Pro Forma Analysis",
                "Accretion/Dilution Analysis",
                "Synergy Analysis"
            ],
        },
        custom: {
            name: "Custom Financial Model",
            sections: [
                "Model Assumptions",
                "Revenue Projections",
                "Expense Projections",
                "Capital Structure",
                "Free Cash Flow Analysis",
                "Valuation"
            ],
        }
    };
    
    if (!templates[type]) {
        return res.status(404).json({ error: 'Template not found' });
    }
    
    res.json(templates[type]);
});

// Start the server
app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
