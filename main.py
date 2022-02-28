from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PIL import ImageTk,Image
from os import path
import sqlite3
import os
import time
import shutil
import pandas as pd
import subprocess as sp
import database as db

# ROOT WINDOW 
root = Tk()

current_directory = os.getcwd()

# START UP PROGRAMS FUNCTIONS
def program_setup():
    database_file = path.exists('database/cigarettes.db')
    database_dir = path.exists('database')
    stock_take_dir = path.exists('stock take output')
    
    if not database_dir:
        os.mkdir('database')
        os.mkdir('database/backup')
    
    if not database_file:
        database_label = Label(root, text='Database not found! One was created for you')
        database_label.grid(row=8, column=0, columnspan=2, sticky=W, pady=(10, 0), padx=(5, 0))
        db.database_init()

    if not stock_take_dir:
        os.mkdir('stock take output')

# FUNCTIONS
def database_backup():
    dir_list = os.listdir('database/backup')
    directory = len(dir_list)

    try:
        delete_file = dir_list[0]
    except IndexError:
        time_stamp = time.strftime("%d%m%Y-%H%M%S")
        new_backup = shutil.copy('database/cigarettes.db', f'database/backup/cigarettes{time_stamp}.db')
    else:
        if directory == 100:
            os.remove(f'database/backup/{delete_file}')
            time_stamp = time.strftime("%d%m%Y-%H%M%S")
            new_backup = shutil.copy('database/cigarettes.db', f'database/backup/cigarettes{time_stamp}.db')
        else:
            time_stamp = time.strftime("%d%m%Y-%H%M%S")
            new_backup = shutil.copy('database/cigarettes.db', f'database/backup/cigarettes{time_stamp}.db')

def add_sales():
    file =  filedialog.askopenfilename(initialdir = current_directory, title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))    
    if file == "":
        update_sales_label.config(text="No Sales Added")
    else:
        # Backup Database
        database_backup()
        # Read in excel
        read_info = pd.read_excel(file, header=0)
        # Date to list
        data = read_info.to_numpy().tolist()
        # Put list to database
        conn = sqlite3.connect('database/cigarettes.db')
        c = conn.cursor()
        # Add sales to database
        for x in range(len(data)):
            query = ("""UPDATE cigarettes 
                        SET sales = sales + ? 
                        WHERE barcode = ?
                    """)
            c.execute(query, (data[x][2], data[x][0])) 
        conn.commit()
        conn.close()

        update_sales_label.config(text="Sales Added")

def add_purchases():
    file =  filedialog.askopenfilename(initialdir = current_directory, title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))    
    if file == "":
        update_purchases_label.config(text="No Purchases Added")
    else:
        # Backup Database
        database_backup()
        # Read in excel
        read_info = pd.read_excel(file, header=0)
        # Date to list
        data = read_info.to_numpy().tolist()
        # Put list to database
        conn = sqlite3.connect('database/cigarettes.db')
        c = conn.cursor()
        # Add sales to database
        for x in range(len(data)):
            query = ("""UPDATE cigarettes 
                        SET purchases = purchases + ? 
                        WHERE barcode = ?
                    """)
            c.execute(query, (data[x][2], data[x][0])) 
        conn.commit()
        conn.close()

        update_purchases_label.config(text="Purchases Added")

def stock_take():
    database_backup()
    conn = sqlite3.connect('database/cigarettes.db')
    c = conn.cursor()
    c.execute("SELECT * FROM cigarettes")
    r = c.fetchall()
    conn.commit()
    conn.close()
    
    # Create dictionary for excel
    results = {'Stock Code': [],
           'Stock Description': [],
           'Stock on hand Should Be': []}

    # loop through stock and add to dictionary and update database
    for data in r:
        results['Stock Code'].append(data[0])
        results['Stock Description'].append(data[1])
        cal = (data[2] + data[3]) - data[4]
        results['Stock on hand Should Be'].append(cal)

        # Update database
        conn = sqlite3.connect('database/cigarettes.db')
        c = conn.cursor()
        query = ("""UPDATE cigarettes 
                        SET 
                        stockOnHand = ?,
                        purchases = ?,
                        sales = ?
                        WHERE barcode = ?
                    """)
        c.execute(query, (cal, 0, 0, data[0])) 
        conn.commit()
        conn.close()

    # WRITE STOCK ON HAND TO EXCEL WORKBOOK
    # Create dataframe
    df = pd.DataFrame(results)
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('stock take output/cigarettes on hand.xlsx', engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    # Formate excel columns
    number_format = workbook.add_format({'num_format': '0'})
    worksheet.set_column('A:A', 18, number_format)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 22)
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    update_stocktake_cal_label.config(text='Stocktake and database update successfully')

def add_stock_item():
    add_stock = Toplevel()
    add_stock.attributes("-topmost", True)
    add_stock.title('Add Stock Item')
    add_stock.geometry('340x325')

    # FUNCTIONS
    def add():
        database_backup()
        conn = sqlite3.connect('database/cigarettes.db')
        c = conn.cursor()
        c.execute("SELECT barcode FROM cigarettes")
        database_barcodes = c.fetchall()
        conn.commit()
        conn.close()

        current_barcodes = []
        for x in database_barcodes:
            current_barcodes.append(x[0])

        barcode = int(barcode_input.get())
        item_name = stock_name_input.get().upper()

        if barcode in current_barcodes:
            update_add_label.config(text='Barcode Already Exists')
        elif item_name == '':
            update_add_label.config(text='Please Enter Item Name')
        else:
            conn = sqlite3.connect('database/cigarettes.db')
            c = conn.cursor()
            query = ("INSERT INTO cigarettes (barcode, itemName, stockOnHand, purchases, sales) VALUES (?, ?, ?, ?, ?)")
            c.execute(query, (barcode, item_name, 0, 0, 0)) 
            conn.commit()
            conn.close()
            barcode_input.delete(0,END)
            stock_name_input.delete(0,END)

            # Read in Master Stock Sheet
            df = pd.read_excel('master stock sheet/Cigarette Master Stock Sheet.xls', sheet_name='Sheet1') 
            # Add New Item to End of Sheet
            df.loc[len(df.index)] = [barcode, item_name] 

            # Start Writing Process
            writer = pd.ExcelWriter('master stock sheet/Cigarette Master Stock Sheet.xls', engine='xlsxwriter')
            # Convert the dataframe to an XlsxWriter Excel object.
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            # Get the xlsxwriter workbook and worksheet objects.
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            # Formate excel columns
            number_format = workbook.add_format({'num_format': '0'})
            worksheet.set_column('A:A', 18, number_format)
            worksheet.set_column('B:B', 35)
            writer.save()

            update_add_label.config(text='Successfully Added Stock Item')
    

    # Input Barcode
    barcode_label = Label(add_stock, text='Barcode')
    barcode_label.grid(row=0, column=0, padx=(10,0), pady=(10,0))

    barcode_input = Entry(add_stock)
    barcode_input.grid(row=1, column=0, padx=(10,0), pady=(10,0))

    # Input Stock Item Name
    stock_name_label = Label(add_stock, text='Stock Name')
    stock_name_label.grid(row=0, column=1, padx=(10,0), pady=(10,0))

    stock_name_input = Entry(add_stock)
    stock_name_input.grid(row=1, column=1, padx=(10,0), pady=(10,0))

    # BUTTONS
    add = Button(add_stock, text='Add', command=add)
    add.grid(row=2, column=0, columnspan=2, sticky=W+E, padx=(10,0), pady=(10,0))

    update_add_label = Label(add_stock, text='')
    update_add_label.grid(row=3, column=0, sticky=W, pady=(10, 0), padx=(5, 0))

def stock_levels():
    file =  filedialog.askopenfilename(initialdir = current_directory, title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))    
    if file == "":
        update_stocklevels_label.config(text="No Stock Level Set/Reset")
    else:
        # Backup Database
        database_backup()
        # Read in excel
        read_info = pd.read_excel(file, header=0)
        # Date to list
        data = read_info.to_numpy().tolist()
        # Put list to database
        conn = sqlite3.connect('database/cigarettes.db')
        c = conn.cursor()
        # Add sales to database
        for x in range(len(data)):
            query = ("""UPDATE cigarettes 
                        SET stockOnHand = ? 
                        WHERE barcode = ?
                    """)
            c.execute(query, (data[x][2], data[x][0])) 
        conn.commit()
        conn.close()

        update_stocklevels_label.config(text="Stock Level Set/Reset")

def read_me():
    programName = "notepad.exe"
    fileName = "read me/README.txt"
    sp.Popen([programName, fileName])

# RUN START UP PROGRAM FUNCTION
program_setup()

# BUTTONS
# Add Sales Button and Info Label
add_sales = Button(root, text='Add Sales', width=20, command=add_sales)
add_sales.grid(row=0, column=0, sticky=W, padx=(10,0), pady=(10,0))

update_sales_label = Label(root, text='')
update_sales_label.grid(row=0, column=1, sticky=W, pady=(10, 0), padx=(5, 0))

# Add Purchases and Info Label
add_purchases = Button(root, text='Add Purchases', width=20, command=add_purchases)
add_purchases.grid(row=1, column=0, sticky=W, padx=(10,0), pady=(10,0))

update_purchases_label = Label(root, text='')
update_purchases_label.grid(row=1, column=1, sticky=W, pady=(10, 0), padx=(5, 0))

# Stock take Calulation
stocktake_cal = Button(root, text='Cacluate Stock Take', width=20, command=stock_take)
stocktake_cal.grid(row=2, column=0, sticky=W, padx=(10,0), pady=(10,0))

update_stocktake_cal_label = Label(root, text='')
update_stocktake_cal_label.grid(row=2, column=1, pady=(10, 0), padx=(5, 0))

# Add Stock Levels and Info Label
add_stock_levels = Button(root, text='Set / Reset Stock Levels', width=20, command=stock_levels)
add_stock_levels.grid(row=3, column=0, sticky=W, padx=(10,0), pady=(10,0))

update_stocklevels_label = Label(root, text='')
update_stocklevels_label.grid(row=3, column=1, sticky=W, pady=(10, 0), padx=(5, 0))

# Add New Stock Item
add_stock_levels = Button(root, text='Add Stock Item', width=20, command=add_stock_item)
add_stock_levels.grid(row=4, column=0, sticky=W, padx=(10,0), pady=(10,0))

# Read Me Button
read_me = Button(root, text='Read Me', width=20, command=read_me)
read_me.grid(row=5, column=0, sticky=W, padx=(10,0), pady=(10,0))    

# Reset Database Button
reset = Button(root, text='Reset Cigarette Database', width=20, command=db.reset)
reset.grid(row=6, column=0, sticky=W, padx=(10,0), pady=(10,0))

# Restore database
restore = Button(root, text='Restore Database', width=20, command=db.restore)
restore.grid(row=7, column=0, sticky=W, padx=(10,0), pady=(10,0))

# Exit Button
exit_btn = Button(root, text='Exit', command=root.quit)
exit_btn.grid(row=9, column=0, columnspan=2, sticky=W+E, pady=(10, 0), padx=(10, 0))

# ROOT WINDOW CONFIG
root.title('Sasol De Bron - Cigarette Stock Count')
root.iconbitmap('icons/smoking.ico')
root.geometry('410x360')
root.columnconfigure(1, minsize=242)

# RUN PROGRAM
root.mainloop()