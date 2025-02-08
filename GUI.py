# -*- coding: utf-8 -*-
"""
"" SIT DSC 2302 (2023) lab03 Demand Forecasting Tool (GUI Module)
"""

# Step 0:
# Create the Main GUI Window --------------------------------------------------

import tkinter as tk
from tkinter import TclError, ttk, Text
from tkinter import *
from tkinter.messagebox import showerror, showwarning, showinfo
import pandas as pd
from pandas.tseries.offsets import DateOffset
import numpy as np

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

import Lab_3_Forecasting as lab_integration

# main Window
MainWindow = tk.Tk()
MainWindow.title ('Demand Forecasting')
MainWindow.geometry("910x450+50+50")
MainWindow.resizable(False,False)
# mune list
MainMenu =tk.Menu(MainWindow)
MainWindow.config(menu=MainMenu)
FileMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='File', menu=FileMenu)
FileMenu.add_command(label='Open from')
FileMenu.add_command(label='Save to')
FileMenu.add_separator()
FileMenu.add_command(label='Exit', command=MainWindow.destroy)
HelpMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='Help', menu=HelpMenu)
HelpMenu.add_command(label='About')

# Step 1: 
# show author name ------------------------------------------------------------
AuthorFrame = ttk.Frame(MainWindow)
AuthorFrame.grid(column = 0, row = 0)
ttk.Label(AuthorFrame,text = "Author : Developer name").grid( column = 0, row =0 , sticky=tk.W)
ttk.Label(AuthorFrame,text = "Group : Group name " ).grid( column = 0, row = 1, sticky=tk.W)


# Step 2: 
# upload data from Excel file -------------------------------------------------
# Store to data series

# Global variables 
MonthHist=pd.Series(24)
LEDHist = pd.Series(24)
MonitorHist = pd.Series(24)
LaptopHist = pd.Series(24)
MobileHist = pd.Series (24)

InputFileName = tk.StringVar()
# Update the path to your file
#InputFileName = "DataSet1.xlsx"

File_label = ttk.Label(MainWindow, text="Input File:")
# File_label.pack(fill='x', expand=True)
File_label.grid(column = 1, row=0, sticky=tk.E, padx=5, pady=5)

File_entry = ttk.Entry(MainWindow, width=45)
File_entry.grid(column = 2, row=0, columnspan=2,  sticky=tk.W, padx=5, pady=5)
File_entry.focus()


# Step 2.1: uploading button function  -------------- 
def Upload_data():
    
    global MonthHist
    global LEDHist 
    global MonitorHist 
    global LaptopHist 
    global MobileHist 
    
    """ upload data from excel file """
    InputFileName = File_entry.get()

    if InputFileName =="": 
        showinfo( title='Data uploading', message='Please type in a valid file name.' )
    else: 
        try: 
            ExcelFileInput = pd.ExcelFile(InputFileName)
        except:
            showinfo(title="File open", message = "Error is found for file openning.")
        else: 
            df1 = pd.read_excel(ExcelFileInput, "RawData")

            MonthHist = df1["Month"]
            LEDHist = df1["LEDLight"]
            MonitorHist =df1["Monitor"]
            LaptopHist = df1["Laptop"]
            MobileHist = df1["Mobile"]    
            # show message
            showinfo(title='Data uploading', message='The historical data has been uploaded successfully!')
   
Upload_button = ttk.Button(MainWindow, text="Upload Data", command = Upload_data)
Upload_button.grid(column = 1, row = 0)
#--------------------------------------------------------------

# Step 3 : select a product from combobox and display historical data ---------
# Step 3.1 select an item
Selected_item = ""
ItemSelection_cb = ttk.Combobox(MainWindow, textvariable=Selected_item)

ItemSelection_cb['values'] = ["LEDLight", "Monitor", "Laptop", "Mobile"]

ItemSelection_cb.set("")
# prevent typing a value
ItemSelection_cb['state'] = 'readonly'    
ItemSelection_cb.grid(column=0, row= 2, sticky = tk.E)    

# separator 
SeperatorFrame = ttk.Frame(MainWindow, width=800, height=50)
SeperatorFrame.grid (column = 0, row=1, columnspan =4, sticky=tk.EW)


# Step 3.2: disply an empty Treeview (for historical data later) ------
HistoricalDataLabel = ttk.Label(MainWindow, text = "Historical data for:" )
HistoricalDataLabel.grid(column = 0, row = 2, sticky=tk.W)

# first 12 months  --------------

HistColumns = ('SerialNo', 'HistMonth', 'HistSales' )
HistView = ttk.Treeview(MainWindow, columns=HistColumns, height=12, show='headings')

HistView.column("SerialNo", width=50 )
HistView.column("HistMonth", width=100 )
HistView.column("HistSales", width=100)

HistView.heading('SerialNo', text='No')
HistView.heading('HistMonth', text='Date')
HistView.heading('HistSales', text='Sales')

for i in range(12):
    HistView.insert('', tk.END, iid = i+1, text = "", values=(i+1, "" , ""))
   
HistView.grid(column = 0, row=3)

#  second 12 months  ---------------------
HistView2 = ttk.Treeview(MainWindow, columns=HistColumns, height=12, show='headings')
HistView2.column("SerialNo", width=50 )
HistView2.column("HistMonth", width=100 )
HistView2.column("HistSales", width=100)

HistView2.heading('SerialNo', text='No')
HistView2.heading('HistMonth', text='Date')
HistView2.heading('HistSales', text='Sales')

for i in range(12):
    HistView2.insert('', tk.END, iid = i+1, text = "", values=(12+i+1, "" , ""))
   
HistView2.grid(column = 1, row=3)

#-----------------------------------------------------------------

# Step 3.3 display historical data for the selected item
def Item_changed(event):
    """ handle the month changed event """
    
    ItemName=ItemSelection_cb.get()
    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        match ItemName:
            case 'LEDLight':
                for i in range(12):
                    HistView.item(i+1, text = LEDHist[i], values=(i+1, MonthHist[i].date(), LEDHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = LEDHist[12+i], values=(12+i+1, MonthHist[12+i].date(), LEDHist[12+i]))
            case 'Monitor': 
                for i in range(12):
                    HistView.item(i+1, text = MonitorHist[i], values=(i+1, MonthHist[i].date(), MonitorHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = MonitorHist[12+i], values=(12+i+1, MonthHist[12+i].date(), MonitorHist[12+i]))

            case 'Laptop' :
                for i in range(12):
                    HistView.item(i+1, text = LaptopHist[i], values=(i+1, MonthHist[i].date(), LaptopHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = LaptopHist[12+i], values=(12+i+1, MonthHist[12+i].date(), LaptopHist[12+i]))

            case 'Mobile' :
                for i in range(12):
                    HistView.item(i+1, text = MobileHist[i], values=(i+1, MonthHist[i].date(), MobileHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = MobileHist[12+i], values=(12+i+1, MonthHist[12+i].date(), MobileHist[12+i]))


# bind the selected value changes    
ItemSelection_cb.bind('<<ComboboxSelected>>', Item_changed)
# -----------------------------------------------------------------------------
    
# Step 4: display forecasting parameters --------------------------------------
p_string =tk.StringVar()
d_string = tk.StringVar()
q_string = tk.StringVar()
alpha_string = tk.StringVar()
beta_string = tk.StringVar()
gamma_string = tk.StringVar()

# set default paramters
p_string = "2"
d_string = "0"
q_string = "1"
alpha_string = "0.2"
beta_string = "0.3"
gamma_string = "0.0"

p_value = int()
d_value = int()
q_value = int()
alpha_value = float()
beta_value = float()
gamma_value = float()

ParameterFrame = ttk.Frame(MainWindow, width=253, height=70)
# ParameterFrame['padding'] = (2,2,2,2)
ParameterFrame['borderwidth'] = 5
ParameterFrame['relief'] ="groove"
ParameterFrame.grid_propagate(False)
ParameterFrame.grid(column = 0, row = 4)

ttk.Label(ParameterFrame, text = "Forecasting parameters:").grid( column = 0, row =0, columnspan= 4, sticky=tk.W)

ttk.Label(ParameterFrame, text = "p:").grid (column = 0, row =1, sticky=tk.E)
ttk.Label(ParameterFrame, text = "d:").grid (column = 2, row =1, sticky=tk.E)
ttk.Label(ParameterFrame, text = "q:").grid (column = 4, row =1, sticky=tk.E)

p_entry = ttk.Entry(ParameterFrame,   width = 4)
d_entry = ttk.Entry(ParameterFrame,   width = 4)
q_entry = ttk.Entry(ParameterFrame,   width = 4)
p_entry.grid(column=1,row =1)
d_entry.grid(column=3, row =1)
q_entry.grid(column=5, row =1)
p_entry.insert(0, p_string)
d_entry.insert(0, d_string)
q_entry.insert(0, q_string)

ttk.Label(ParameterFrame, text = "level:").grid (column = 0, row =2, sticky=tk.E)
ttk.Label(ParameterFrame, text = "trend:").grid (column = 2, row =2, sticky=tk.E)
ttk.Label(ParameterFrame, text = "season:").grid (column = 4, row =2, sticky=tk.E)
alpha_entry = ttk.Entry(ParameterFrame,  width = 4)
beta_entry = ttk.Entry (ParameterFrame,  width = 4)
gamma_entry = ttk.Entry (ParameterFrame, width = 4)
alpha_entry.grid(column=1,row =2)
beta_entry.grid(column=3, row =2)
gamma_entry.grid(column=5, row =2)
alpha_entry.insert(0, alpha_string)
beta_entry.insert(0, beta_string)
gamma_entry.insert(0, gamma_string)


# Step 5: display empty forecast result in treeview ---------------------------
ResultLabel = ttk.Label(MainWindow, text = " Demand Forecast " )
ResultLabel.grid(column = 3, row = 2)


ResultColumns = ('SerialNo', 'Month', 'Forecast' )
ResultView = ttk.Treeview(MainWindow, columns=ResultColumns, height=12, show='headings')

ResultView.column("SerialNo", width=50 )
ResultView.column("Month", width=100 )
ResultView.column("Forecast", width=100)

ResultView.heading('SerialNo', text='No')
ResultView.heading('Month', text='Month')
ResultView.heading('Forecast', text='Forecasts')

for i in range(12):
    ResultView.insert('', tk.END, iid = i+1, text = "", values=(i+1, "" , ""))
   
ResultView.grid(column = 3, row=3)    
#---------------------------------------------------------------

# Step 6: Shwo Algorithm Radiobutton for selection ----------------------------

ttk.Label(MainWindow,text = "Algorithms").grid( column = 2, row = 2 )


RadioFrame = ttk.Frame(MainWindow, width=150, height=120)
RadioFrame['padding'] = (10,10,10,10)
RadioFrame['borderwidth'] = 5
RadioFrame['relief'] = 'sunken'
#Stop the frame from propagating the widget to be shrink or fit
RadioFrame.grid_propagate(False)
RadioFrame.grid(column = 2, row = 3, sticky=tk.NW)

AlgoSelected = tk.IntVar()
AlgoButton1 = ttk.Radiobutton(RadioFrame, text='Algo 1: AR', variable=AlgoSelected, value=1)
AlgoButton1.grid(column=1, row=0, padx=2, pady=2, sticky=tk.W)


AlgoButton2 = ttk.Radiobutton(RadioFrame, text='Algo 2: ARMA', variable=AlgoSelected, value=2)
AlgoButton2.grid(column=1, row=1, padx=2, pady=2, sticky=tk.W)

AlgoButton3 = ttk.Radiobutton(RadioFrame, text='Algo 3: SES', variable=AlgoSelected, value=3)
AlgoButton3.grid(column=1, row=2, padx=2, pady=2, sticky=tk.W)

AlgoButton3 = ttk.Radiobutton(RadioFrame, text='Algo 4: HWES', variable=AlgoSelected, value=4)
AlgoButton3.grid(column=1, row=3, padx=2, pady=2, sticky=tk.W)

AlgoSelected.set(2)


# Step 7: Accuracy checkboxs for selections -----------------------------------

MAE_var = tk.StringVar()
MSE_var = tk.StringVar()
RMSE_var = tk.StringVar()
MASE_var = tk.StringVar()

AccuracyFrame = ttk.Frame(MainWindow, width=150, height=120)
AccuracyFrame['padding'] = (10,10,10,10)
AccuracyFrame['borderwidth'] = 5
AccuracyFrame['relief'] = 'sunken'
AccuracyFrame.grid_propagate(False)
AccuracyFrame.grid(column = 2, row = 3, columnspan = 2, sticky=tk.SW)

def calculate_MSE(actual, forecast):
    """Calculate Mean Squared Error (MSE)"""
    squared_errors = [(a - f) ** 2 for a, f in zip(actual, forecast)]
    mse = sum(squared_errors) / len(actual)
    return mse

# display MAE error 
def Display_MAE_error ():
    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    ar_forecast = lab_integration.forecast_AR(my_hist)  # Assuming your AR forecasting function
                    mae_ar = lab_integration.MAE_value(My_Y2022, ar_forecast)
                    MAEResult_label.config(text='%.2f' % mae_ar)

                case 2:
                    # ARMA algo is selected
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    mae_ARMA = lab_integration.MAE_value (My_Y2022, ARMA_forecast)
                    MAEResult_label.config(text = '%.2f'%mae_ARMA)
                case 3:
                    # SES algo is selected
                    ses_forecast = lab_integration.forecast_simple_exponential_smoothing(my_hist, alpha_value)  # Assuming your SES forecasting function
                    mae_ses = lab_integration.MAE_value(My_Y2022, ses_forecast)
                    MAEResult_label.config(text='%.2f' % mae_ses)

                case 4:
                    # Holt-Winters algo is selected
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    mae_HWES = lab_integration.MAE_value (My_Y2022, HWES_forecast)
                    MAEResult_label.config(text = '%.2f'%mae_HWES)                
                case _: 
                    print ("Error in algo or accuracy measure selection.")

#display MSE error
def Display_MSE_error():
    if len(LEDHist) < 24:
        showinfo(title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for the selected product
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]

        # 2. checking forecasting parameters
        try:
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:
            showinfo(title="Parameter error",
                     message="p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("AR algorithm is selected")
                    print("Forecasting with AR algorithm...")

                    # Replace this with your AR forecasting logic
                    ar_forecast = lab_integration.forecast_AR(my_hist)  # Assuming your AR forecasting function
                    print("AR Forecast:", ar_forecast)
                    mse_ar = lab_integration.MSE_value(My_Y2022, ar_forecast)
                    print("MSE for AR:", mse_ar)
                    MSEResult_label.config(text='%.2f' % mse_ar)

                case 2:
                    # ARMA algo is selected
                    print("ARMA")
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2022, p_value, 0, q_value)
                    print("Prediction for Y2022: ")
                    print(ARMA_forecast)
                    mse_ARMA = lab_integration.MSE_value(My_Y2022, ARMA_forecast)
                    print('MSE for selected product and algo: %f' % mse_ARMA)
                    MSEResult_label.config(text='%0.2f' % mse_ARMA)
                case 3:
                    # SES algo is selected
                    print("Simple Exponential Smoothing (SES) algorithm is selected")
                    # Forecast using SES algorithm
                    ses_forecast = lab_integration.forecast_simple_exponential_smoothing(my_hist, alpha_value)  # Assuming your SES forecasting function
                    # Calculate MSE
                    mse_ses = lab_integration.MSE_value(My_Y2022, ses_forecast)
                    # Update the MSE result label
                    MSEResult_label.config(text='%0.2f' % mse_ses)
                    print(f"SES Forecast: {ses_forecast}")
                    print(f"MSE for SES: {mse_ses}")

                case 4:
                    # Holt-Winters algo is selected
                    print("Holt-Winters")
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2022, alpha_value, beta_value, gamma_value)
                    print("Prediction for Y2022")
                    print(HWES_forecast)
                    mse_HWES = lab_integration.MSE_value(My_Y2022, HWES_forecast)
                    print('MSE for selected product and algo: %f' % mse_HWES)
                    MSEResult_label.config(text='%0.2f' % mse_HWES)
                case _:
                    print("Error in algo or accuracy measure selection.")

# display RMSE error 
def Display_RMSE_error ():
    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("AR algorithm is selected")
                    print("Forecasting with AR algorithm...")
                    
                    # Replace this with your AR forecasting logic
                    ar_forecast = lab_integration.forecast_AR(my_hist)  # Assuming your AR forecasting function
                    print("AR Forecast:", ar_forecast)
                    mae_ar = lab_integration.RMSE_value(My_Y2022, ar_forecast)
                    print("MAE for AR:", mae_ar)
                    RMSEResult_label.config(text='%.2f' % mae_ar)

                case 2:
                    # ARMA algo is selected
                    print ("ARMA")
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    print ("prediction for Y2022: ")
                    print (ARMA_forecast)
                    mae_ARMA = lab_integration.RMSE_value (My_Y2022, ARMA_forecast)
                    print ('MAE for selected produt and algo: %f' %mae_ARMA)
                    RMSEResult_label.config(text = '%.2f'%mae_ARMA)
                case 3:
                    # SES algo is selected
                    print("Simple Exponential Smoothing (SES) algorithm is selected")
                    # Forecast using SES algorithm
                    ses_forecast = lab_integration.forecast_simple_exponential_smoothing(my_hist, alpha_value)
                    # Calculate MAE
                    mae_ses = lab_integration.RMSE_value(My_Y2022, ses_forecast)
                    # Update the MAE result label
                    RMSEResult_label.config(text='%.2f' % mae_ses)
                    print(f"SES Forecast: {ses_forecast}")
                    print(f"MAE for SES: {mae_ses}")

                case 4:
                    # Holt-Winters algo is selected
                    print ("Hot Winters")
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    print ("prediction for Y2022")
                    print (HWES_forecast)
                    mae_HWES = lab_integration.RMSE_value (My_Y2022, HWES_forecast)
                    print ('MAE for selected product and algo: %f' %mae_HWES)
                    RMSEResult_label.config(text = '%.2f'%mae_HWES)                
                case _: 
                    print ("Error in algo or accuracy measure selection.")

def MAE_changed():
    if MAE_var.get() == "selected":
        Display_MAE_error()

    else: 
        if MAE_var.get() == "unselected":
            MAEResult_label.config(text = "")    
        else:
            showinfo(title="err", message=f" MAE var = {MAE_var.get()}, !")

def RMSE_changed():
    if RMSE_var.get() == "selected":
        Display_RMSE_error ()
    else:    
        if RMSE_var.get() == "unselected":
            RMSEResult_label.config(text = "")

    
def MSE_changed():
    if MSE_var.get() == "selected":
        Display_MSE_error()
    else:
        if MSE_var.get() == "unselected":
            MSEResult_label.config(text="")
        
def MASE_changed ():
    showinfo (title = "MASE", message = "Please incorporate a function for MASE forecast accuracy checking.")


MAE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MAE',
                command=MAE_changed,
                variable=MAE_var,
                onvalue="selected",
                offvalue="unselected")
MAE_checkbox.grid(column=0, row = 1, sticky=tk.W)

RMSE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='RMSE',
                command=RMSE_changed,
                variable=RMSE_var,
                onvalue='selected',
                offvalue='unselected')
RMSE_checkbox.grid(column=0, row = 2, sticky=tk.W)

MSE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MSE',
                command=MSE_changed,
                variable=MSE_var,
                onvalue='selected',
                offvalue='unselected')
MSE_checkbox.grid(column=0, row = 3, sticky=tk.W)

MASE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MASE',
                command=MASE_changed,
                variable=MASE_var,
                onvalue='selected',
                offvalue='unselected')
MASE_checkbox.grid(column=0, row = 4, sticky=tk.W)


MAEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding = 2)
MAEResult_label.config(text = "1020")
MAEResult_label.grid(column=1, row=1, sticky=tk.E)

RMSEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding=2)
RMSEResult_label.config(text = "3333")
RMSEResult_label.grid(column=1, row=2, sticky=tk.E)

MSEResult_label = ttk.Label(AccuracyFrame, relief="groove", width=5, padding=2)
MSEResult_label.config(text="0.00")
MSEResult_label.grid(column=1, row=3, sticky=tk.E)


# Step 8:  forecat generation and accuracy checking ---------------------------

def Forecasting_clicked ():
    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("Autoregression (AR) algorithm is selected")
                    # Replace this with your AR forecasting logic
                    ar_forecast = lab_integration.forecast_AR(my_hist)  # Assuming your AR forecasting function
                    # Update the forecast in the Treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i + 1, values=(i + 1, New_month[i].date(), ar_forecast[i]))
                    
                case 2:
                    # ARMA algo is selected
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2022, p_value, 0, q_value)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), ARMA_forecast[i]))
    
                case 3:
                    # SES algo is selected
                    print("Simple Exponential Smoothing (SES) algorithm is selected")
                    # Replace this with your SES forecasting logic
                    ses_forecast = lab_integration.forecast_simple_exponential_smoothing(my_hist, alpha_value)  # Assuming your SES forecasting function
                    # Update the forecast in the Treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i + 1, values=(i + 1, New_month[i].date(), ses_forecast[i]))

                case 4:
                    # Holt-Winters algo is selected
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2022, alpha_value, beta_value, gamma_value)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), HWES_forecast[i]))

                case _: 
                    print ("Error in algo or accuracy measure selection.")

    
    
ForecastGenerationbutton = tk.Button(MainWindow, text = 'Forecasting >', \
                                     fg ='black', command=Forecasting_clicked)
ForecastGenerationbutton.grid( column = 2, row = 3)


# Step 9: result and comments by student --------------------------------------

CommentFrame = ttk.Frame(MainWindow, width=656, height=70)
CommentFrame['padding'] = (2,2,2,2)
CommentFrame['borderwidth'] = 5
CommentFrame['relief'] ="groove"
#Stop the frame from propagating the widget to be shrink or fit
CommentFrame.grid_propagate(False)
CommentFrame.grid(column = 1, row = 4, columnspan=4)

ttk.Label(CommentFrame, text = "Result & comment by student:").grid( column = 0, row =0)

# ---------------------------
MainWindow.mainloop()
# ---------------------------