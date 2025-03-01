/**
 * Forecasting Service
 * 
 * This service provides methods for generating financial forecasts
 * using various statistical and financial modeling techniques.
 */

const math = require('mathjs');

class ForecastingService {
    constructor() {
        // Configuration options
        this.defaultOptions = {
            method: 'linear', // linear, exponential, moving-average, arima
            periods: 5,       // number of periods to forecast
            confidence: 0.95, // confidence level for prediction intervals
            seasonality: 1,   // seasonality period (e.g., 4 for quarterly, 12 for monthly)
        };
    }

    /**
     * Generate a forecast based on historical data
     * @param {Array} historicalData - Array of historical data points
     * @param {Object} options - Forecasting options
     * @returns {Object} - Forecast results including predicted values and statistics
     */
    generateForecast(historicalData, options = {}) {
        // Merge default options with provided options
        const forecastOptions = { ...this.defaultOptions, ...options };
        
        // Validate data
        if (!Array.isArray(historicalData) || historicalData.length < 2) {
            throw new Error('Historical data must be an array with at least 2 data points');
        }
        
        // Select forecasting method based on options
        switch (forecastOptions.method) {
            case 'linear':
                return this.linearRegression(historicalData, forecastOptions);
            case 'exponential':
                return this.exponentialSmoothing(historicalData, forecastOptions);
            case 'moving-average':
                return this.movingAverage(historicalData, forecastOptions);
            case 'arima':
                return this.arima(historicalData, forecastOptions);
            default:
                throw new Error(`Forecasting method "${forecastOptions.method}" not supported`);
        }
    }

    /**
     * Generate a forecast using linear regression
     * @param {Array} historicalData - Array of historical data points
     * @param {Object} options - Forecasting options
     * @returns {Object} - Forecast results
     */
    linearRegression(historicalData, options) {
        // Create x values (time indices)
        const x = historicalData.map((_, index) => index + 1);
        const y = historicalData;
        
        // Calculate linear regression parameters
        const n = x.length;
        const sumX = x.reduce((sum, value) => sum + value, 0);
        const sumY = y.reduce((sum, value) => sum + value, 0);
        const sumXY = x.reduce((sum, value, index) => sum + value * y[index], 0);
        const sumX2 = x.reduce((sum, value) => sum + value * value, 0);
        
        const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        const intercept = (sumY - slope * sumX) / n;
        
        // Generate forecast values
        const forecast = [];
        for (let i = 1; i <= options.periods; i++) {
            const period = n + i;
            const forecastValue = intercept + slope * period;
            forecast.push(forecastValue);
        }
        
        // Calculate error metrics
        const fitted = x.map(xi => intercept + slope * xi);
        const residuals = y.map((yi, i) => yi - fitted[i]);
        const sse = residuals.reduce((sum, residual) => sum + residual * residual, 0);
        const mse = sse / n;
        const rmse = Math.sqrt(mse);
        
        // Calculate R-squared
        const meanY = sumY / n;
        const totalSumOfSquares = y.reduce((sum, yi) => sum + Math.pow(yi - meanY, 2), 0);
        const rSquared = 1 - (sse / totalSumOfSquares);
        
        // Calculate prediction intervals
        const tValue = 1.96; // Approximate t-value for 95% confidence
        const standardError = Math.sqrt(sse / (n - 2));
        
        const predictionIntervals = forecast.map((value, i) => {
            const period = n + i + 1;
            const predictionVariance = standardError * Math.sqrt(1 + (1 / n) + 
                (Math.pow(period - meanY, 2) / totalSumOfSquares));
            
            return {
                lower: value - tValue * predictionVariance,
                upper: value + tValue * predictionVariance
            };
        });
        
        return {
            method: 'linear',
            historicalData,
            forecast,
            predictionIntervals,
            statistics: {
                slope,
                intercept,
                rSquared,
                rmse
            }
        };
    }

    /**
     * Generate a forecast using exponential smoothing
     * @param {Array} historicalData - Array of historical data points
     * @param {Object} options - Forecasting options
     * @returns {Object} - Forecast results
     */
    exponentialSmoothing(historicalData, options) {
        // Simple Exponential Smoothing
        const alpha = options.alpha || 0.3; // Smoothing factor
        
        // Initialize smoothed values
        const smoothed = [historicalData[0]];
        
        // Calculate smoothed values
        for (let i = 1; i < historicalData.length; i++) {
            const newSmoothed = alpha * historicalData[i] + (1 - alpha) * smoothed[i - 1];
            smoothed.push(newSmoothed);
        }
        
        // Last smoothed value
        const lastSmoothed = smoothed[smoothed.length - 1];
        
        // Generate forecast values (all the same for simple exponential smoothing)
        const forecast = Array(options.periods).fill(lastSmoothed);
        
        // Calculate error metrics
        const errors = historicalData.map((actual, i) => 
            i === 0 ? 0 : actual - smoothed[i - 1]
        );
        
        const sse = errors.reduce((sum, error) => sum + error * error, 0);
        const mse = sse / (historicalData.length - 1);
        const rmse = Math.sqrt(mse);
        
        // Calculate prediction intervals
        const standardError = Math.sqrt(mse);
        const tValue = 1.96; // Approximate t-value for 95% confidence
        
        const predictionIntervals = forecast.map((value, i) => {
            const interval = tValue * standardError * Math.sqrt(1 + (i + 1));
            return {
                lower: value - interval,
                upper: value + interval
            };
        });
        
        return {
            method: 'exponential',
            historicalData,
            forecast,
            predictionIntervals,
            statistics: {
                alpha,
                rmse
            }
        };
    }

    /**
     * Generate a forecast using moving average
     * @param {Array} historicalData - Array of historical data points
     * @param {Object} options - Forecasting options
     * @returns {Object} - Forecast results
     */
    movingAverage(historicalData, options) {
        const windowSize = options.windowSize || 3;
        
        if (windowSize > historicalData.length) {
            throw new Error('Window size cannot be larger than the historical data length');
        }
        
        // Calculate moving averages
        const movingAverages = [];
        for (let i = windowSize - 1; i < historicalData.length; i++) {
            const window = historicalData.slice(i - windowSize + 1, i + 1);
            const average = window.reduce((sum, value) => sum + value, 0) / windowSize;
            movingAverages.push(average);
        }
        
        // Last moving average
        const lastMA = movingAverages[movingAverages.length - 1];
        
        // Generate forecast values (all the same for simple moving average)
        const forecast = Array(options.periods).fill(lastMA);
        
        // Calculate error metrics
        const errors = [];
        for (let i = windowSize - 1; i < historicalData.length; i++) {
            const actual = historicalData[i];
            const predicted = movingAverages[i - windowSize + 1];
            errors.push(actual - predicted);
        }
        
        const sse = errors.reduce((sum, error) => sum + error * error, 0);
        const mse = sse / errors.length;
        const rmse = Math.sqrt(mse);
        
        // Calculate prediction intervals
        const standardError = Math.sqrt(mse);
        const tValue = 1.96; // Approximate t-value for 95% confidence
        
        const predictionIntervals = forecast.map((value, i) => {
            const interval = tValue * standardError * Math.sqrt(1 + (i + 1) / windowSize);
            return {
                lower: value - interval,
                upper: value + interval
            };
        });
        
        return {
            method: 'moving-average',
            historicalData,
            forecast,
            predictionIntervals,
            statistics: {
                windowSize,
                rmse
            }
        };
    }

    /**
     * Generate a forecast using ARIMA (Auto-Regressive Integrated Moving Average)
     * Note: This is a simplified implementation for demonstration purposes
     * @param {Array} historicalData - Array of historical data points
     * @param {Object} options - Forecasting options
     * @returns {Object} - Forecast results
     */
    arima(historicalData, options) {
        // ARIMA parameters
        const p = options.p || 1; // AR order
        const d = options.d || 0; // Differencing order
        const q = options.q || 1; // MA order
        
        // For a real implementation, you would use a proper ARIMA library
        // This is a very simplified approximation
        
        // Apply differencing if d > 0
        let diffData = [...historicalData];
        for (let i = 0; i < d; i++) {
            diffData = this.diff(diffData);
        }
        
        // Use a simple AR model for demonstration
        // In a real implementation, you would estimate AR and MA parameters
        const arParams = new Array(p).fill(0.7 / p);
        const maParams = new Array(q).fill(0.3 / q);
        
        // Generate forecast for differenced data
        const diffForecast = [];
        for (let i = 0; i < options.periods; i++) {
            let forecast = 0;
            
            // AR component
            for (let j = 0; j < p; j++) {
                const index = diffData.length - j - 1;
                if (index >= 0) {
                    forecast += arParams[j] * diffData[index];
                }
            }
            
            // Add forecast value and update diffData for next iteration
            diffForecast.push(forecast);
            diffData.push(forecast);
        }
        
        // Undo differencing to get actual forecast
        let forecast = [...diffForecast];
        for (let i = 0; i < d; i++) {
            forecast = this.undiff(forecast, historicalData[historicalData.length - 1 - i]);
        }
        
        // For simplicity, use constant prediction intervals
        const stdDev = this.standardDeviation(historicalData);
        const tValue = 1.96; // Approximate t-value for 95% confidence
        
        const predictionIntervals = forecast.map((value, i) => {
            const interval = tValue * stdDev * Math.sqrt(i + 1);
            return {
                lower: value - interval,
                upper: value + interval
            };
        });
        
        return {
            method: 'arima',
            historicalData,
            forecast,
            predictionIntervals,
            statistics: {
                p,
                d,
                q
            }
        };
    }

    /**
     * Apply differencing to a time series
     * @param {Array} data - Time series data
     * @returns {Array} - Differenced series
     */
    diff(data) {
        const result = [];
        for (let i = 1; i < data.length; i++) {
            result.push(data[i] - data[i - 1]);
        }
        return result;
    }

    /**
     * Undo differencing
     * @param {Array} diffData - Differenced data
     * @param {number} lastValue - Last value of the original series
     * @returns {Array} - Undifferenced data
     */
    undiff(diffData, lastValue) {
        const result = [lastValue];
        for (let i = 0; i < diffData.length; i++) {
            result.push(result[i] + diffData[i]);
        }
        return result.slice(1);
    }

    /**
     * Calculate standard deviation of an array
     * @param {Array} data - Array of numbers
     * @returns {number} - Standard deviation
     */
    standardDeviation(data) {
        const mean = data.reduce((sum, value) => sum + value, 0) / data.length;
        const squaredDiffs = data.map(value => Math.pow(value - mean, 2));
        const variance = squaredDiffs.reduce((sum, value) => sum + value, 0) / data.length;
        return Math.sqrt(variance);
    }

    /**
     * Generate a Monte Carlo simulation forecast
     * @param {Array} historicalData - Historical data points
     * @param {Object} options - Simulation options
     * @returns {Object} - Simulation results
     */
    monteCarloSimulation(historicalData, options = {}) {
        const iterations = options.iterations || 1000;
        const periods = options.periods || 5;
        const confidenceLevel = options.confidence || 0.95;
        
        // Calculate mean and standard deviation of percent changes
        const percentChanges = [];
        for (let i = 1; i < historicalData.length; i++) {
            percentChanges.push((historicalData[i] / historicalData[i - 1]) - 1);
        }
        
        const mean = percentChanges.reduce((sum, value) => sum + value, 0) / percentChanges.length;
        const variance = percentChanges.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / percentChanges.length;
        const stdDev = Math.sqrt(variance);
        
        // Run simulations
        const simulations = [];
        const lastValue = historicalData[historicalData.length - 1];
        
        for (let i = 0; i < iterations; i++) {
            const simulation = [lastValue];
            for (let j = 0; j < periods; j++) {
                // Generate random percentage change using normal distribution
                const rand = this.boxMullerTransform();
                const percentChange = mean + stdDev * rand;
                
                // Apply percentage change to previous value
                const newValue = simulation[j] * (1 + percentChange);
                simulation.push(newValue);
            }
            simulations.push(simulation.slice(1)); // Remove initial value
        }
        
        // Calculate forecast and prediction intervals
        const forecast = [];
        const predictionIntervals = [];
        
        for (let i = 0; i < periods; i++) {
            const periodValues = simulations.map(sim => sim[i]);
            periodValues.sort((a, b) => a - b);
            
            // Mean forecast
            const mean = periodValues.reduce((sum, value) => sum + value, 0) / iterations;
            forecast.push(mean);
            
            // Prediction intervals
            const lowerIndex = Math.floor(iterations * (1 - confidenceLevel) / 2);
            const upperIndex = Math.floor(iterations * (1 - (1 - confidenceLevel) / 2));
            
            predictionIntervals.push({
                lower: periodValues[lowerIndex],
                upper: periodValues[upperIndex]
            });
        }
        
        return {
            method: 'monte-carlo',
            historicalData,
            forecast,
            predictionIntervals,
            statistics: {
                iterations,
                mean,
                stdDev
            }
        };
    }

    /**
     * Generate a normally distributed random number using Box-Muller transform
     * @returns {number} - Random number from standard normal distribution
     */
    boxMullerTransform() {
        const u1 = Math.random();
        const u2 = Math.random();
        
        const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
        return z0;
    }
}

module.exports = new ForecastingService();