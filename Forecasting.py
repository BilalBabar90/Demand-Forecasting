# -*- coding: utf-8 -*-
"""
SIT DSC 2302 (2023) Lab03 Demand Forecasting Tool (Forecasting Function Module)
"""

import numpy as np
import pandas as pd
from tkinter.messagebox import showerror, showwarning, showinfo
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.holtwinters import SimpleExpSmoothing
from statsmodels.tsa.ar_model import AR
from statsmodels.tsa.holtwinters import Holt
from statsmodels.tsa.holtwinters import ExponentialSmoothing

from math import sqrt


# Autoregressive Moving-Average (ARMA) forecasting ----------------------------
def forecast_ARIMA(data, order1, order2, order3):
    # data (Pandas Data Series):  historical data for the past 12 months
    # model parameters: order1, order2, order3
    # Return forecast for 12 months

    My_hist = data.to_numpy()
    My_model = ARIMA(My_hist, order=(order1, order2, order3))
    My_model_fit = My_model.fit()
    My_forecast = My_model_fit.forecast(steps=12)
    My_forecast = My_forecast.astype(int)

    return My_forecast


# Holt Winter's Exponential Smoothing (HWES) forecasting ----------------------
def forecast_HWExpSmoothing(data, my_alpha, my_deta, my_gamma):
    # data (Pandas Data Series)-- historical data for the past 12 months
    # my_alpha -- smoothing level from 0 to 1    
    # my_deta -- smoothing trend from 0 to 1
    # my_gamma -- smoothing seasonal from 0 to 1
    # return forecast for the next 12 months 

    My_hist = data.to_numpy()
    My_model = ExponentialSmoothing(My_hist, trend=my_deta, seasonal=my_gamma)
    My_model_fit = My_model.fit(smoothing_level=my_alpha)
    My_forecast = My_model_fit.forecast(steps=12)
    My_forecast = My_forecast.astype(int)

    return My_forecast


# Autoregressive (AR) forecasting ----------------------------------------------
def forecast_AR(data, p):
    """
    Autoregressive (AR) forecasting function.

    Parameters:
        data (Pandas Data Series): Historical data for past p months.
        p (int): Autoregressive order.

    Returns:
        forecast (numpy array): Forecast for the next 12 months.
    """
    # Convert historical data to a numpy array
    hist_data = data.to_numpy()

    # Create an AR model and fit it to the historical data
    model = AR(hist_data)
    model_fit = model.fit(maxlag=p)

    # Generate forecasts for the next 12 months
    forecast = model_fit.predict(start=len(hist_data), end=len(hist_data) + 11)

    return forecast.astype(int)


# Simple Exponential Smoothing (SES) forecasting ------------------------------
def forecast_SES(data, alpha):
    """
    Simple Exponential Smoothing (SES) forecasting function.

    Parameters:
        data (Pandas Data Series): Historical data for past 12 months.
        alpha (float): Smoothing parameter between 0 and 1.

    Returns:
        forecast (numpy array): Forecast for the next 12 months.
    """
    # Convert historical data to a numpy array
    hist_data = data.to_numpy()

    # Create an SES model and fit it to the historical data
    model = SimpleExpSmoothing(hist_data)
    model_fit = model.fit(smoothing_level=alpha)

    # Generate forecasts for the next 12 months
    forecast = model_fit.forecast(steps=12)

    return forecast.astype(int)


# Forecast performance checking
# MAE : Mean absolute Error
def MAE_value(expected_value, forecasted_value):
    mae = np.mean(np.abs(expected_value - forecasted_value))
    return mae

# MSE: Mean Squared Error
def MSE_value(expected_value, forecasted_value):
    mse = np.mean((expected_value - forecasted_value) ** 2)
    return mse

# MSE : Root Mean Squared Error
def RMSE_value(expected_value, forecasted_value):
    rmse = np.mean((expected_value - forecasted_value) ** 2)
    return rmse


def main():
    # read data from Excel file
    InputFileName = "DataSet1.xlsx"
    ExcelTab = "RawData"
    try:
        ExcelData = pd.ExcelFile(InputFileName)
    except:
        showinfo(title="File Error",
                 message=f"{InputFileName} opening error.")
    else:
        Data_df1 = pd.read_excel(ExcelData, ExcelTab)

    # LED study ----------------------
    # Forecasting variables
    p = 5
    d = 0
    q = 2
    alpha = 0.1
    beta = 0.6
    gamma = 0

    # Historical data
    My_data_list = Data_df1["LEDLight"].array
    My_month_list = Data_df1["Month"].array

    LED_Y2021 = pd.Series(data=My_data_list[0:12], index=My_month_list[0:12])
    LED_Y2022 = pd.Series(data=My_data_list[12:], index=My_month_list[12:])

    # Forecast data
    AR_forecast = forecast_ARIMA(LED_Y2022, p, 0, 0)
    MA_forecast = forecast_ARIMA(LED_Y2022, 0, 0, q)
    ARMA_forecast = forecast_ARIMA(LED_Y2022, p, 0, q)

    ExpSmoothing_forecast1 = forecast_HWExpSmoothing(LED_Y2022, alpha, 0, 0)
    ExpSmoothing_forecast2 = forecast_HWExpSmoothing(LED_Y2022, alpha, beta, 0)
    ExpSmoothing_forecast3 = forecast_HWExpSmoothing(LED_Y2022, alpha, beta, gamma)

    # plotting line chart
    plt.figure(figsize=(16, 7))

    My_13_month = LED_Y2022.index[11] + relativedelta(months=+1)
    My_forecast_index = pd.date_range(My_13_month, periods=12, freq='m')

    # plotting historical data and forecast
    plt.plot(LED_Y2022.index, LED_Y2022, color='black', label="Historical")
    plt.plot(My_forecast_index, AR_forecast, marker='o', label='AR')
    plt.plot(My_forecast_index, MA_forecast, marker='o', label='AM')
    plt.plot(My_forecast_index, ARMA_forecast, marker='o', label="ARAM")
    plt.plot(My_forecast_index, ExpSmoothing_forecast1, marker='D', label="SES")
    plt.plot(My_forecast_index, ExpSmoothing_forecast2, marker='D', label="Holt")

    plt.title("LED")
    My_x_index = pd.date_range(LED_Y2022.index[0], periods=24, freq="m")

    plt.xticks(My_x_index)
    plt.tick_params(labelrotation=45)
    plt.legend()
    # -------------------------------------------------------------------------
       
    # check forecast performance for LED Light---------------------------------
    expected_LED_2022 = LED_Y2022.array
    mae_ARMA_LED = MAE_value (expected_LED_2022, ARMA_forecast)
    print ('MAE for LED Light: %f' %mae_ARMA_LED)

    mse_ARMA_LED = MSE_value(expected_LED_2022, ARMA_forecast)
    print('MSE for LED Light: %f' % mse_ARMA_LED)
    
    rmse_ARMA_LED = RMSE_value (expected_LED_2022, ARMA_forecast)    
    print ('RMSE for LED Light: %f' %rmse_ARMA_LED)
    # -------------------------------------------------------------------------
    

if __name__ == "__main__":
    main()
