import requests
from bs4 import BeautifulSoup as bs
import re
import json
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import xlsxwriter


root = tk.Tk()
root.resizable(False, False) 
root.title("MacroTrends - GetData")

abbr_frame = tk.Frame(root)
abbr_frame.pack(pady=10)

abbr_label = tk.Label(abbr_frame, text="Enter company abbreviation: ")
abbr_label.pack(side=tk.LEFT)

abbr_entry = tk.Entry(abbr_frame, width=50)
abbr_entry.pack(side=tk.LEFT)

abbr_label = tk.Label(text="Example: AMZN, META, TSLA, JNJ etc.", fg='red')
abbr_label.pack(pady=10)

checkbox_frame = tk.Frame(root)
checkbox_frame.pack(pady=10)

income_statement_var = tk.IntVar()
income_statement_checkbox = tk.Checkbutton(checkbox_frame, text="Income statement", variable=income_statement_var)
income_statement_checkbox.pack(side=tk.LEFT)

balance_sheet_var = tk.IntVar()
balance_sheet_checkbox = tk.Checkbutton(checkbox_frame, text="Balance sheet", variable=balance_sheet_var)
balance_sheet_checkbox.pack(side=tk.LEFT)

cash_flow_statement_var = tk.IntVar()
cash_flow_statement_checkbox = tk.Checkbutton(checkbox_frame, text="Cash flow statement", variable=cash_flow_statement_var)
cash_flow_statement_checkbox.pack(side=tk.LEFT)

key_ratios_var = tk.IntVar()
key_ratios_checkbox = tk.Checkbutton(checkbox_frame, text="Key financial ratios", variable=key_ratios_var)
key_ratios_checkbox.pack(side=tk.LEFT)

frequency_frame = tk.Frame(root)
frequency_frame.pack(pady=10)

frequency_label = tk.Label(frequency_frame, text="Frequency:")
frequency_label.pack(side=tk.LEFT)

frequency_var = tk.StringVar()
frequency_var.set("annually")

annually_radio = tk.Radiobutton(frequency_frame, text="annually", variable=frequency_var, value="annually")
annually_radio.pack(side=tk.LEFT)

quarterly_radio = tk.Radiobutton(frequency_frame, text="quarterly", variable=frequency_var, value="quarterly")
quarterly_radio.pack(side=tk.LEFT)

def get_data():
    abbr = abbr_entry.get()
    
    if abbr != '':
        selected_options = []
        if income_statement_var.get():
            selected_options.append("income-statement")
        if balance_sheet_var.get():
            selected_options.append("balance-sheet")
        if cash_flow_statement_var.get():
            selected_options.append("cash-flow-statement")
        if key_ratios_var.get():
            selected_options.append("key-financial-ratios")
            
        if selected_options:
            abbr = abbr.upper()
            filename = abbr + '_' + frequency_var.get() + '.xlsx'

            if os.path.exists('Desktop/' + filename):
                messagebox.showinfo("File already exists", "The file - " + filename + " - already exists!")

            else:
                writer = pd.ExcelWriter(filename, engine='xlsxwriter')
                for option in selected_options:
                    try:
                        url = 'https://www.macrotrends.net/stocks/charts/' + abbr + '/a/' + option + '?freq=' + frequency_var.get()
                        r = requests.get(url)
                        p = re.compile(r' var originalData = (.*?);\r\n\r\n\r', re.DOTALL)
                        data = json.loads(p.findall(r.text)[0])

                        headers = list(data[0].keys())
                        headers.remove('popup_icon')
                        result = []

                        for row in data:
                            soup = bs(row['field_name'], features='xml')
                            field_name = soup.select_one('a, span').text
                            fields = list(row.values())[2:]
                            fields.insert(0, field_name)
                            result.append(fields)

                        df = pd.DataFrame(result, columns=headers)
                        df.to_excel(writer, sheet_name=option, index=False)

                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to retrieve data for abbreviation {abbr}. Please check the entered abbreviation and try again.")
                       
                        print(e)
                        
                        writer.close()
                        os.remove(filename)

                else:
                    writer.close()
                    messagebox.showinfo("Successfully", f"The file - {filename} - has been created!")
        
        else:
            messagebox.showinfo("No options selected", "Please select at least one option!")
    
    else:
        messagebox.showinfo("Error!", "Please insert abreviation!")

btn = tk.Button(root, text="ENTER", command=get_data)
btn.pack(pady=10)

root.mainloop()