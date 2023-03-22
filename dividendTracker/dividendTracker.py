#Imports
import xlwings as xw
import yfinance as yf
import pandas as pd
from datetime import date, datetime
import time, os, sys
import sqlite3


'''
Function for getting currency rate using currency tickers on Yahoo Finance
'''
def getAndUpdate_Currency_Rates():
    #HKD to SGD
    hkd2sgd = yf.Ticker("HKDSGD=X").fast_info.last_price
    #USD to SGD
    usd2sgd = yf.Ticker("SGD=X").fast_info.last_price
    
    #Update Rates in Ref sheet
    ref['B2'].value = hkd2sgd
    ref['B3'].value = usd2sgd

'''
Function for getting the earliest date of purchase for each stock from the excel sheet
Also to get the buy date with the total amount of shares at that point in time (For calculating dividend)

Returns a dict with (ticker as key) and (date as value),
and a df with columns ("Ticker", "Date", "Amount_of_shares")
'''
def get_stock_buy_date():
    #Column A contains the dates, Column C contains the tickers
    dates_and_tickers = buyTransac.range("A3:K3").options(expand='down').value
    date_dict = {}
    buydate_list = []
    
    #Save only ticker symbol and dates into a dict
    for row in dates_and_tickers:
        #Create nested list for df of amount_of_shares at given date
        buydate_list.append([row[2], row[0], row[10]])
        #As only the earliest buy date is needed, skip the current ticker if it already exist in the dict
        if row[2] in date_dict:
            pass
        else:
            date_dict[row[2]] = row[0].strftime("%Y-%m-%d") #yyyy-mm-dd string instead of datetime type
    
    checkDF = pd.DataFrame(buydate_list).reset_index(drop= True)
    checkDF.rename(columns = {0: "Ticker", 1: "Date", 2: "Amount_of_shares"}, inplace=True)
    checkDF.loc[:, "Date"] = checkDF["Date"].dt.tz_localize("Asia/Singapore")
    
    return date_dict, checkDF

'''
Function for pulling stock's dividend information from yFinance Module into a Pandas DF

For Perfect Shape (1830.HK), 
dividend information will be scraped from aastocks instead of yahoo finance, as it is not accurate

returns a pandas df
'''
def get_and_clean_info_from_source(ticker):
    try:     
        #Check ticker symbol, and returns a df with columns = ['date', 'dividends']
        if ticker == "1830.HK": #Perfect Shape Medical Stock Ticker
            dfs = pd.read_html('http://www.aastocks.com/en/stocks/analysis/company-fundamental/dividend-history?symbol=01830')

            uncleanedDf = dfs[25] #26th df in the webpage
            
            #Clean the df
            uncleanedDf.drop(columns = [0, 1, 2, 4, 6, 7], inplace = True)
            uncleanedDf.rename(columns = {3: 'Dividends', 5: 'Date'}, inplace = True)
            
            #Clean Dividends column
            divs = uncleanedDf.loc[:, ['Dividends']]
            divs = divs["Dividends"].str.split("HKD ").str[1]

            #Clean Date column
            dates = uncleanedDf.loc[:, ['Date']]
            dates = dates['Date'].str.replace("/", "-", regex = False)
                
            #Change dtypes, datetime has to be tz aware for comparison later on
            divs = divs.astype('float64')
            dates = pd.to_datetime(dates, format= '%Y-%m-%d', errors= 'coerce').dt.tz_localize('Asia/Singapore') #I live in SG, hence the SG timezone
            #Concat back both columns into a single df
            df = pd.concat([dates, divs], axis = 1) 
            #Drop NaN or NaT values
            df = df.dropna() 
        else:
            stock = yf.Ticker(ticker)
            dividend_data = stock.dividends
            df = pd.DataFrame(dividend_data).reset_index()
            
        return df
    except:
        pass
    
'''
Function for;
1. Initializing a table in the sqlite3 db, if it does not exist already
2. Get latest ex-date of dividends from the db
3. Get Dividend information from sources, df will be empty if no new dividend information
4. If there is new dividend information, clean the df and append it into sqlite3 db, else ignore and go next

'''
def get_Dividend_Information_into_sqldb(conn, buy_dates, checkDF):
    c = conn.cursor()
    
    #Initialize table if not created, else ignore
    c.execute('CREATE TABLE IF NOT EXISTS dividends (ticker text, date text, dividends float, amount_of_shares integer)')
    conn.commit()
    
    #Get latest date of dividend for each stock from sqlite3 db and save to a dict
    c.execute("SELECT ticker, MAX(date(date, 'localtime')) FROM dividends group by ticker")
    results = c.fetchall()
    list = []
    for result in results:
        list.append(result)
    tmp_dict = dict(list) 
    
    #Clean the date field, by keeping only yyyy-mm-dd, and save into another dict
    result_dict = {}
    for key, item in tmp_dict.items():
        item = item.split(" ", 1)[0]
        result_dict[key] = item
        
    #Get dividend information from tickers and yFinance
    tickers = mainSheet.range("B8").options(expand='down').value
    for ticker in tickers:
        try:
            df = get_and_clean_info_from_source(ticker)
                
            #Check whether db is empty or if ticker is not in db (Newly added stock)
            if result_dict and ticker in result_dict:
                dfCleaned = df[~(df['Date'] <= result_dict[ticker])].reset_index(drop= True) #Get date from existing db
            #If db is empty or stock is newly bought, use a default date to pull all dividends from that date onwards
            else:
                dfCleaned = df[~(df['Date'] <= buy_dates[ticker])].reset_index(drop= True) #Get date from excel

            #Check if dfClean is empty or not, skips appending df to sqlite3 db if so (No new dividend information)
            if not dfCleaned.empty:
                #Insert a column to keep track of stock's ticker symbol and amount of shares
                dfCleaned.insert(0, 'Ticker', ticker)
                dfCleaned.insert(3, 'Amount_of_shares', 0)
                
                #Change Dates into same datetime aware format as the dates in checkDF for comparison
                dfCleaned["Date"].dt.tz_convert("Asia/Singapore")
                
                #Get the amount_of_shares for each dividend entry at given point in time
                for divIndex, divSeries in dfCleaned.iterrows():
                    #Sort the date in checkDF so that it iterates over the latest buy date first, in desc order
                    for boughtIndex, boughtSeries in checkDF[checkDF["Ticker"] == ticker].sort_values(by=["Date"], ascending=False).iterrows():
                        if divSeries["Date"] > boughtSeries["Date"]:
                            dfCleaned.at[divIndex, "Amount_of_shares"] = boughtSeries["Amount_of_shares"]
                            break
                    
                #Transfer pandas df into sqlite db
                dfCleaned.to_sql('dividends', conn, if_exists='append', index=False)
                conn.commit()
        except Exception as error:
            #Error Handling
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

'''
Function for pulling stock price information from yfinance Module
Returns a list of dicts that contains stock information
    list of dicts will always be in the same order as shown in the Stock table in the main sheet
'''
def get_Stock_Information(conn):
    #Initialize an Empty list
    rows = []
    
    #Find sum of dividend for each stock and save into a dict
    c = conn.cursor()
    c.execute('SELECT ticker, sum(dividends * amount_of_shares) as total_div FROM dividends group by ticker')
    results = c.fetchall()
    dict_Of_Dividends = dict(results)
    
    #Get Stock ticker information from excel sheet and sqlite3 db
    tickers = mainSheet.range("B8").options(expand='down').value
    for ticker in tickers:
        try:
            #Initialize Stock ticker
            data = yf.Ticker(ticker)

            #Get Fast information of stock
            price = data.fast_info
            
            #Create dictionary for wanted information
            new_row = {
                "ticker": ticker,
                "current_Price": price.last_price,
                "last_close_Price": price.previous_close,
                "total_dividend_to_Date": dict_Of_Dividends[ticker]
            }
        
            #Append data (new row) to list
            rows.append(new_row)
        except Exception as e:
            #Error Handling
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            #Append Empty Row
            rows.append({})
    
    return rows #List of dicts


'''
Function for writing the cleaned data into the excel sheet
Only needs to iterate over each row in excel as the list of dicts will always be in the same order and size
Added a check just in case
'''
def write_value_to_excel(data):
    #TargetRange have to include tickers and columns that need to be filled
    last_col = mainSheet.range("B8").end('right').column
    last_row = mainSheet.range("B8").end('down').row
    targetRange = mainSheet.range((8,2), (last_row, last_col)).rows
    counter = 0
    
    #Iterate over each row in the Stock table residing in the main sheet
    for row in targetRange:
        #Check if Dict's Ticker is the same as the Excel sheet's ticker (Just in case)
        if data[counter] and data[counter]["ticker"] == row[0].value:
            row[2].value = data[counter]["current_Price"] #'Current Price' Column
            row[3].value = data[counter]["last_close_Price"] #'Last Close Price' Column
            row[11].value = (data[counter]["total_dividend_to_Date"] - float(row[10].value)) #'Dividends Collected' Colummn, row[10] is the Miscellaneous Fees column
        #If data is empty, replace nan with something (For delisted stocks like MNACT)
        elif not data[counter]:
            row[2].value = 0 #'Current Price' Column
            row[3].value = 0 #'Last Close Price' Column
            #For stocks that are delisted, like MNACT, dividends have to be manually keyed in :(
        
        counter += 1
        
    #Write date and time of last update to cell above button to keep track of refreshes
    mainSheet['K1'].value = datetime.now()
            
def main():
    #Initialize conenction to sqlite3 db
    conn = sqlite3.connect(r"C:\Users\Zhen Xuan\OneDrive\Desktop\CodingStuff\PersonalProjects\AutomatedDividendTracker\dividend_record.db")
    
    #Main Functions
    getAndUpdate_Currency_Rates()
    buy_dates, checkDF = get_stock_buy_date()
    get_Dividend_Information_into_sqldb(conn, buy_dates, checkDF)
    data = get_Stock_Information(conn)
    write_value_to_excel(data)
    
    #Close connection to sqlite3 db
    conn.close()
    
if __name__ == "__main__":
    xw.Book("dividendTracker.xlsm").set_mock_caller()
    main()


#Get Values from Excel
#These variables are GLOBAL
wb = xw.Book.caller()
#Sheets in the excel workbook
mainSheet = wb.sheets('Portfolio')
# graphs = wb.sheets('Graphs')
buyTransac = wb.sheets('Buy_Transactions')
#sellTransac = wb.sheets('Sell_Transactions')
ref = wb.sheets('Ref')

