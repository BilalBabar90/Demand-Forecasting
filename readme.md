# Demand Forecasting Tool

This project is a Demand Forecasting Tool developed for SIT DSC 2302 (2023) Lab03. It includes multiple time series forecasting methods such as ARIMA, Holt-Winter's Exponential Smoothing, Autoregressive (AR), and Simple Exponential Smoothing (SES). The tool reads historical data from an Excel file and generates forecasts for the next 12 months.

## Features

- **ARIMA Forecasting:** Autoregressive Integrated Moving Average model.
- **Holt-Winter's Exponential Smoothing:** Seasonal and trend forecasting.
- **Autoregressive (AR) Model:** Forecasting based on previous data points.
- **Simple Exponential Smoothing (SES):** Smoothing past data for trend estimation.
- **Performance Evaluation:** Calculates MAE, MSE, and RMSE to evaluate forecast accuracy.
- **Visualization:** Line plots to compare historical and forecasted data.

## Getting Started

### Prerequisites

Ensure the following Python packages are installed:

```bash
pip install numpy pandas matplotlib statsmodels python-dateutil
