import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

def portfolioByCostGraph():
    # Get costs from mainsheet  
    cost = mainSheet.range("A1:B4").options(convert=pd.DataFrame).value.reset_index()
    # Cast the column "Cost" as a float
    cost = cost.astype({'Cost': 'float64'})
    # Round the values in column "Cost" to 2 d.p
    cost["Cost"] = cost["Cost"].round(2)
    
    fig, ax = plt.subplots()
    ax.pie(cost["Cost"], labels= cost["Cost"])
    ax.legend(cost["index"],
              loc= "upper left",
              bbox_to_anchor= (1, 0, 1.5, 1))
    ax.set_title("Portfolio by Cost")
    
    graphs.pictures.add(fig, name= "test1", update=True)
    
def graph2():
    x = [3, 6, 9, 12, 15]
    y = [2, 4, 6, 8, 10]
    
    fig, ax = plt.subplots()
    ax.plot(x, y)
    
    graphs.pictures.add(fig, name= "test2", update=True)

def main():
    portfolioByCostGraph()
    graph2()

if __name__ == "__main__":
    xw.Book("dividendTracker.xlsm").set_mock_caller()
    main()
    
#Get Values from Excel
#These variables are GLOBAL
wb = xw.Book.caller()
mainSheet = wb.sheets('Portfolio')
graphs = wb.sheets('Graphs')