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
        
        amzn = yf.Ticker('AMZN')
        historical_data = amzn.history()

        Apple = yf.Ticker("AAPL")
        Apple_data = Apple.history(period = "max", interval = "1d")

        Microsoft = yf.Ticker("MSFT")
        Microsoft_data = Microsoft.history(period = "max", interval = "1d")

        Tesla = yf.Ticker("TSLA")
        Tesla_data = Tesla.history(period = "max", interval = "1d")

        FNFund = yf.Ticker("FNORX")
        FNFund_data = FNFund.history(period = "max", interval = "1d")

        FSTCFund = yf.Ticker("FTUTX")
        FSTCFund_data = FSTCFund.history(period = "max", interval = "1d")

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
        df_Data_Concat = pd.concat(Dict_DFS.values(), axis = 1, keys=Dict_DFS.keys())
        display(df_Data_Concat)
        return df_Data_Concat, Dict_DFS
    
    def initial_excel_writer(self, df_Data_Concat, Dict_DFS):
        df_Data_Concat.index = df_Data_Concat.index.astype(str)
        writer = pd.ExcelWriter(r"C:\Users\mattp\Finance Project\Logic\All Assets Final.xlsx")
        df_Data_Concat.to_excel(writer, sheet_name="All Assets")

        for key in Dict_DFS:
            df_new_dataframe = df_Data_Concat[key]
            df_new_dataframe.to_excel(writer, sheet_name=key)
        writer.close()

    def practical_analysis(self, Dict_DFS):

        def reading_back():
            df_dict = pd.read_excel(r"C:\Users\mattp\Finance Project\Practical Analysis\All Assets Price Change.xlsx",  sheet_name = None)
            df_AllAssets = df_dict['All Assets']
            df_AllAssets = df_AllAssets.drop(["Unnamed: 0"], axis = 1)
            del df_dict['All Assets']
            df_AllAssets.columns = df_AllAssets.columns.to_series().replace('Unnamed:\s\d+',np.nan,regex=True).ffill().values
            return df_dict, df_AllAssets
        
        def firstpoint(df_common):
            starting_point = df_common[df_common.Open>0].iloc[0]
            date_starting_point = starting_point[0]
            df_common["Starting Point"] = date_starting_point
            return df_common
        
        def price_change(df_common):    
            df_common['Date'] = pd.to_datetime(df_common['Date'], errors='coerce')
            
            df_common["Day to Day"] = df_common["Close"].pct_change()
            df_common["Log Returns"] = np.log(df_common["Close"]) - np.log(df_common["Close"].shift(1))
            df_common['Month to Month'] = df_common.groupby(df_common['Date'].dt.to_period('M'))['Day to Day'].transform('mean')
            df_common['Year to Year'] = df_common.groupby(df_common['Date'].dt.to_period('Y'))['Day to Day'].transform('mean')
            
            df_common['Date'] = df_common['Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
            return df_common
        
        def vol(df_common):    
            df_common = df_common.fillna("0")
            df_common["Vol"] = df_common["Log Returns"].rolling(252).std()*(252**0.5)
            df_common = df_common.fillna(0)
            return df_common
        
        def sharpe(df_common):      
            df_common["Sharpe"] = (df_common["Log Returns"] - 0.0552)/df_common["Vol"]
            return df_common
        
        dict_Updated_dfs = {}
        for key in Dict_DFS:
            df_common = Dict_DFS[key]
            df_common = df_common.drop("Unnamed: 0", axis = 1)
            df_common = firstpoint(df_common)
            df_common = price_change(df_common)
            df_common = sharpe(df_common)
            df_common = df_common.fillna('0')
            df_common = sharpe(df_common)
            dict_Updated_dfs[key] = df_common

        return dict_Updated_dfs

    def full_excel_writer(self, dict_Updated_dfs):
        list_dfNames = ["Stock- NVIDIA", "Stock- Amazon", "Stock- Apple", "Stock- Microsoft", "Stock- Tesla", "Mutual Fund- FNF", "Mutual Fund-FSTP", "Bond- 13 Week TBill", "Bond- 5YTBY", "Bond- 10 YTBill Yield", "Commodity- Copper", "Commodity- Oil", "Commodity- Gold", "Commodity- Steel", "Credit Data- SNP Futures", "FX Pair- USD EUR", "FX Pair- USD JPY", "Index Fund- S&P"]
        df_Data_Concat = pd.concat(dict_Updated_dfs.values(), axis = 1, keys=list_dfNames[:])
        
        writer = pd.ExcelWriter(r"C:\Users\mattp\Finance Project\Practical Analysis\Full Finance Project.xlsx")

        for key in dict_Updated_dfs:
            new_dataframe = df_Data_Concat[key]
            new_dataframe.to_excel(writer, sheet_name=key)

        writer.close()

    


