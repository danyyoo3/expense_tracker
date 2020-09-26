#Expense Tracker

##Introduction
Python application to have a personal expense tracker
The program has options to create multiple wallets that are recorded in an Excel file
When the program is run, if no excel file called  "My_Wallets.xlsx", the program automatically creates the file and 
stores all information in the file.
The program uses openpyxl to handle data management through Excel. As well as basic line and bar charts from matplot
to show the progress of your spendings

##Things it can do
The program has an interface that allows you to select an already created wallet or create a new wallet.
When creating the wallet, you are able to specify the currency and the total amount.
Once the wallet is selected, you are able to add spendings and earnings that are recorded in the Excel file. 
The program is also able to show the information.
It has a currency converter option that uses the API from EuroBank. The link is found inside the code
Lastly, with matplot, it shows the trend of your spendings according to the different categories

The reason for using Excel file instead of any database, such as SQL, is to allow direct flexibility for the user
to either manually access the file and edit, or use the program to record his/her information

