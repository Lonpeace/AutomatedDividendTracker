import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import sqlite3

def portfolioByCostGraph():
    # Get costs from mainsheet  
    cost = mainSheet.range("A1:B4").options(convert=pd.DataFrame).value
    # Cast the column "Cost" as a float
    cost = cost.astype({'Cost': 'float64'})
    
    # Create autopct function as label
    def pctLabel(pct, series):
        total = series.sum()
        val = float(round(pct * total/100, 2))
        return f"${val:.2f}\n({pct:.2f}%)"
    
    fig, ax = plt.subplots()
    
    # Create pie chart
    ax.pie(cost["Cost"], autopct= lambda pct: pctLabel(pct, cost["Cost"]), pctdistance= 1.25)
    
    # Create legend
    ax.legend(cost.index,
              fontsize= "small",
              loc= "upper left",
              bbox_to_anchor= (1, 0, 1.5, 1))
    
    # Set chart title
    ax.set_title("Portfolio by Cost")
    
    # Add to excel, or update if already exists
    graphs.pictures.add(fig, name= "pbcGraph", update=True)
    
def divPerYearGraph(conn):
    hkd2sgd = ref['B2'].value
    usd2sgd = ref['B3'].value
    
    sqlQuery = f'''SELECT STRFTIME("%Y", date) AS divYear,
                CASE 
                    WHEN ticker LIKE "%.HK" THEN dividends * amount_of_shares * {hkd2sgd}
                    WHEN ticker = "H78.SI" THEN dividends * amount_of_shares * {usd2sgd}
                    ELSE dividends * amount_of_shares
                END as totalDiv
                FROM dividends'''
    
    df = pd.read_sql_query(sqlQuery, conn)
    plot_df = df.groupby(["divYear"])["totalDiv"].sum()
    
    fig, ax = plt.subplots()
    bars = ax.bar(plot_df.index, plot_df)
    
    # Label bars
    ax.bar_label(bars, fmt= "${:,.2f}")
    
    # Set chart title
    ax.set_title("Dividends collected per Year")
    
    graphs.pictures.add(fig, name= "dpyGraph", update=True)

def main():
    conn = sqlite3.connect(r"C:\Users\Zhen Xuan\OneDrive\Desktop\CodingStuff\PersonalProjects\AutomatedDividendTracker\dividend_record.db")
    portfolioByCostGraph()
    divPerYearGraph(conn)
    
    conn.close()

if __name__ == "__main__":
    xw.Book("dividendTracker.xlsm").set_mock_caller()
    main()
    
#Get Values from Excel
#These variables are GLOBAL
wb = xw.Book.caller()
mainSheet = wb.sheets('Portfolio')
graphs = wb.sheets('Graphs')
ref = wb.sheets('Ref')