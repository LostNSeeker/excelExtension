// ai-service.js
// This service handles integration with the AI backend

class AIService {
    constructor(baseUrl = 'https://api.example.com') {
        this.baseUrl = baseUrl;
        this.headers = {
            'Content-Type': 'application/json'
        };
    }

    // Set authentication token if needed
    setAuthToken(token) {
        this.headers['Authorization'] = `Bearer ${token}`;
    }

    // Generic API request method
    async _request(endpoint, method = 'GET', data = null) {
        const url = `${this.baseUrl}${endpoint}`;
        const options = {
            method,
            headers: this.headers,
            credentials: 'include'
        };

        if (data && method !== 'GET') {
            options.body = JSON.stringify(data);
        }

        try {
            const response = await fetch(url, options);
            
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.error || `API request failed with status ${response.status}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error(`Error in ${method} request to ${endpoint}:`, error);
            throw error;
        }
    }

    // Formula Assistant API
    async getFormulaAssistance(query) {
        return this._request('/api/formula', 'POST', { query });
    }

    // Chat Assistant API
    async sendChatMessage(message, history = []) {
        return this._request('/api/chat', 'POST', { message, history });
    }

    // Financial Model Template API
    async getModelTemplate(modelType) {
        return this._request(`/api/model-templates/${modelType}`, 'GET');
    }

    // PDF Extraction API
    async extractPdfData(file) {
        // For file uploads, we need a different approach than the _request method
        const formData = new FormData();
        formData.append('pdf', file);

        try {
            const response = await fetch(`${this.baseUrl}/api/extract-pdf`, {
                method: 'POST',
                body: formData,
                credentials: 'include'
            });

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.error || `PDF extraction failed with status ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.error('Error extracting PDF data:', error);
            throw error;
        }
    }

    // Market Data API
    async getMarketData(symbols, startDate, endDate) {
        return this._request('/api/market-data', 'POST', { symbols, startDate, endDate });
    }

    // Forecasting API
    async generateForecast(historicalData, forecastPeriod, options = {}) {
        return this._request('/api/forecast', 'POST', { 
            historicalData, 
            forecastPeriod, 
            options 
        });
    }

    // Error Analysis API
    async analyzeFormulaErrors(formula, context) {
        return this._request('/api/analyze-formula', 'POST', { formula, context });
    }
}

// Export the service
window.AIService = AIService;
