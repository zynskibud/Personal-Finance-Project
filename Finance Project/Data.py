import pandas as pd
import numpy as np
import yfinance as yf
from IPython.display import display

class Analysis:
    def __init__(self):
        
        DF_Data_Concat, Dict_DFS = self.loading()
        self.excel_writer(DF_Data_Concat, Dict_DFS)
        self.practical_analysis()

    def loading(self):
        NVIDIA = yf.Ticker("NVDA")
        NVIDIA_data = NVIDIA.history(period = "max", interval = "1d")
        NVIDIA_data

        amzn = yf.Ticker('AMZN')
        historical_data = amzn.history()
        historical_data.head()

        Apple = yf.Ticker("AAPL")
        Apple_data = Apple.history(period = "max", interval = "1d")
        Apple_data

        Microsoft = yf.Ticker("MSFT")
        Microsoft_data = Microsoft.history(period = "max", interval = "1d")
        Microsoft_data

        Tesla = yf.Ticker("TSLA")
        Tesla_data = Tesla.history(period = "max", interval = "1d")
        Tesla_data

        FNFund = yf.Ticker("FNORX")
        FNFund_data = FNFund.history(period = "max", interval = "1d")
        FNFund_data

        FSTCFund = yf.Ticker("FTUTX")
        FSTCFund_data = FSTCFund.history(period = "max", interval = "1d")
        FSTCFund_data

        ThirtTBill = yf.Ticker("^IRX")
        ThirTBill_data = ThirtTBill.history(period = "max", interval = "1d")

        TYield5years = yf.Ticker("^FVX")
        TYield5years_data = TYield5years.history(period = "max", interval = "1d")

        TYield10Years = yf.Ticker("^TNX")
        TYield10Years_data = TYield10Years.history(period = "max", interval = "1d")

        Copper = yf.Ticker("HG=F")
        Copper_data = Copper.history(period = "max", interval = "1d")   

        Oil = yf.Ticker("CL=F")
        Oil_data = Oil.history(period = "max", interval = "1d")
            
        Gold = yf.Ticker("GC=F")
        Gold_data = Gold.history(period = "max", interval = "1d")

        Steel = yf.Ticker("X")
        Steel_data = Steel.history(period = "max", interval = "1d")

        SNPFutures = yf.Ticker("ES=F")
        SNPFutures_data = SNPFutures.history(period = 'max', interval = '1d')
                
        USD_EURO = yf.Ticker("EURUSD=x")
        USD_EURO_data = USD_EURO.history(period = "max", interval="1d")

        Dollar_Yen = yf.Ticker("JPY=X")
        Dollar_Yen_Data = Dollar_Yen.history(period = "max", interval = "1d")

        SNP = yf.Ticker("INDEX")
        SNP_Data = SNP.history(period = "max", interval = "1d") 

        Dict_DFS = {"Stock: NVIDIA": NVIDIA_data, "Stock: Amazon": historical_data, "Stock: Apple": Apple_data, "Stock: Microsoft": Microsoft_data, "Stock: Tesla": Tesla_data, "Mutual Fund: Fidelity Nordic Fund": FNFund_data, "Mutual Fund: Fidelity Select Tele Communications Portfolio": FSTCFund_data, "Bond: 13 Week TBill": ThirTBill_data, "Bond: 5 Year T Bill Yield": TYield5years_data, "Bond: 10 Year T Bill Yield": TYield10Years_data, "Commodity: Copper": Copper_data, "Commodity: Oil": Oil_data, "Commodity: Gold": Gold_data, "Commodity: Steel": Steel_data, "Credit Data: SNP Futures": SNPFutures_data, "FX Pair: USD EUR": USD_EURO_data, "FX Pair: USD JPY": Dollar_Yen_Data, "Index Fund: S&P": SNP_Data}
        DF_Data_Concat = pd.concat(Dict_DFS.values(), axis = 1, keys=Dict_DFS.keys())
        display(DF_Data_Concat)

        return DF_Data_Concat, Dict_DFS
    
    def excel_writer(self, DF_Data_Concat, Dict_DFS):
        writer = pd.ExcelWriter(r"C:\Users\mattp\Finance Project\Logic\All Assets Final.xlsx")
        DF_Data_Concat.to_excel(writer, sheet_name="All Assets")

        for key in Dict_DFS:
            new_dataframe = DF_Data_Concat[key]
            new_dataframe.to_excel(writer, sheet_name=key)
        writer.close()
