import tkinter as tk
from tkinter import filedialog, TOP, INSERT

import pandas

root = tk.Tk()
root.geometry('1280x720')


# This function will be used to open
# file in read mode and only Python files
# will be opened

def open_file():
    file = filedialog.askopenfile(filetypes=[('Excel', '*.xls')])
    if file is not None:
        print(file.name)
        return file.name


l = tk.Label(root, text="FINANCE EXPENSES APP", font=['ROBOTO', 25])

T = tk.Text(root, height=20, width=100)

btn = tk.Button(root, text='Open', command=lambda: open_file())

other_exp_label = tk.Label(root, text="OTHER EXPENSES")
other_exp_label2 = tk.Label(root, text="")

food_companies = ['KAUFLAND', 'BILLA', 'LIDL', 'BOLERO', 'ANET']  # companie names to check for Groceries
petrol_companies = ['BI OIL', 'DEGA', 'LUKOIL', 'EKO', "SHELL", ]  # companie names to check for Gas Expenses

data = pandas.read_excel(open_file())

df = pandas.DataFrame(data)
food_exp = 0
gas_expenses = 0
withdraw_expenses = 0
other_expenses = 0

# df.loc[row][col]

print('OTHER EXPENSES:')

for i in range(9, len(df)):

    gas, food, draw = False, False, False

    draw_col = df.loc[i][5]  # COLUMN from table to get payment method

    if type(draw_col) != float and type(df.loc[i][7]) != float:  # if statement to check for withdraw from ATM
        if 'ATM' in draw_col:
            withdraw_expenses += float(df.loc[i][3])
            draw = True

    if type(df.loc[i][7]) != float:  # if statement to check for Gas expenses using petrol list
        for element in petrol_companies:
            if element in df.loc[i][7]:
                gas_expenses += float(df.loc[i][3])
                gas = True

    if type(df.loc[i][7]) != float:  # if statement to check for Grocerie expenses using Grocerie list
        for element in food_companies:
            if element in df.loc[i][7]:
                food_exp += float(df.loc[i][3])
                food = True

    if not gas and not food and not draw:  # if Expense is not in above statements
        if str(df.loc[i][3]) != 'nan':
            other_expenses += float(df.loc[i][3])

            T.insert(INSERT, f'Cost: {df.loc[i][3]:.2f}  ||  ')
            T.insert(INSERT, f'NAME:{df.loc[i][7]}\n')

            print(float(df.loc[i][3]), df.loc[i][7])

gas = tk.Label(root, text=f'Gas Expenses: {gas_expenses:.2f} BGN', font=['ROBOTO', 15])
food = tk.Label(root, text=f'Food Expenses: {food_exp:.2f} BGN', font=['ROBOTO', 15])
withdraw = tk.Label(root, text=f'Withdraw Expenses: {withdraw_expenses:.2f} BGN', font=['ROBOTO', 15])
useless = tk.Label(root, text=f'Useless Expenses: {other_expenses:.2f} BGN', font=['ROBOTO', 15])

print()
print(f'Gas Expenses: {gas_expenses:.2f} BGN')
print(f'Food Expenses: {food_exp:.2f}')
print(f'Withdraw Expenses: {withdraw_expenses:.2f} BGN')
print(f'Useless Expenses: {other_expenses:.2f} BGN')

l.pack(side=TOP)

other_exp_label.pack()
T.pack()
gas.pack()
food.pack()
withdraw.pack()
useless.pack()

tk.mainloop()
