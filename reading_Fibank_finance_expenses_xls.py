import pandas
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()

root.withdraw()

messagebox.showwarning("Important", "Please select only .xls Files")

file_path = filedialog.askopenfilename(filetypes=[('Excel', '*.xls')])



food_companies = ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET'] #companie names to check for Groceries
petrol_companies = ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', "SHELL", ] #companie names to check for Gas Expenses

data = pandas.read_excel(file_path, sheet_name='Sheet')

df = pandas.DataFrame(data)
food_exp = 0
gas_expenses = 0
withdraw_expenses = 0
other_expenses = 0

# df.loc[row][col]

print('OTHER EXPENSES:')

for i in range(9, len(df)):

    gas,food,draw = False,False,False

    draw_col = df.loc[i][5] #COLUMN from table to get payment method

    if type(draw_col) != float and type(df.loc[i][7]) != float: #if statement to check for withdraw from ATM
        if 'ATM' in draw_col:
            withdraw_expenses += float(df.loc[i][3])
            draw = True

    if type(df.loc[i][7]) != float: #if statement to check for Gas expenses using petrol list
        for element in petrol_companies:
            if element in df.loc[i][7]:
                gas_expenses += float(df.loc[i][3])
                gas = True

    if type(df.loc[i][7]) != float: #if statement to check for Grocerie expenses using Grocerie list
        for element in food_companies:
            if element in df.loc[i][7]:
                food_exp += float(df.loc[i][3])
                food = True

    if not gas and not food and not draw: #if Expense is not in above statements
        if str(df.loc[i][3]) != 'nan':
            other_expenses += float(df.loc[i][3])
            print(float(df.loc[i][3]), df.loc[i][7])

print()
print(f'Gas Expenses: {gas_expenses:.2f} BGN')
print(f'Food Expenses: {food_exp:.2f}')
print(f'Withdraw Expenses: {withdraw_expenses:.2f} BGN')
print(f'Useless Expenses: {other_expenses:.2f} BGN')

root.mainloop()