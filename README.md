# Automated Dividend Tracker

Hello, this is my attempt at creating a dividend tracker with MS Excel, Python and xlwings, tailored for a specific kind of dividend traders.

If you are a person who rarely buy or sell any new stocks, and is fustrated with the need to manually key in a dividend entry everytime it pops up, then this project might just be able to help you. 

For this application to run, you would need to have installed Python (3.11.1), xlwings, MS Excel, yfinance and pandas. (Versions shown in requirements.txt)

Click the "Click me to Refresh Sheet" button to get the latest stock prices, currency rates, and dividend information.

![image](https://user-images.githubusercontent.com/79985278/219069359-c0a0e9f7-9276-401a-852b-a2ed6e44a719.png)

## Notes
For the Portfolio sheet, or the main sheet as shown below, only the column "Miscellaneous Fees" has to be keyed in manually. These fees includes custodian fees and other fees that are deducted from your dividend income, and not fees that are incurred during the buying or selling of shares.

If there is a stock that has been delisted on Yahoo Finance, its dividends will have to be keyed in manually too. For example, the stock "Mapletree NAC Tr" shown below has been merged together with another stock to form "Mapletree Panasia Com Tr", causing it to be delisted from Yahoo Finance and losing all historical data of it's dividends.

Sadly, for the buy and sell transactions, you would still have to manually key in the values up to the "Fees" column, as shown in the tables below.

Buy transaction sheet
![image](https://user-images.githubusercontent.com/79985278/230271341-f5bf6b9c-adda-4c43-99dd-8004eb6dc4cb.png) 

Sell transaction sheet
![image](https://user-images.githubusercontent.com/79985278/230271395-bc238f60-9acc-433d-b0bb-98cd00f12ed7.png)

