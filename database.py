from tkinter import *
from os import path
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import sqlite3
import time
import os
import shutil

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

def reset():
    response = messagebox.askyesno('WARNING MESSAGE', 'You are about to overwrite all stock levels. Would you like to proceed?', icon='warning')
    if response == 1:
        # Backup Database
        database_backup()
        # Master stock file
        file = 'master stock sheet/Cigarette Master Stock Sheet.xls'
        # Read in excel
        data = pd.read_excel(file, header=0)
        # Date to list
        info = data.to_numpy().tolist()

        # Create a database or connect to one
        conn = sqlite3.connect('database/cigarettes.db')
        # Create cursor 
        c = conn.cursor()
        # Query DB
        # Create table when run for first time
        c.execute("""CREATE TABLE IF NOT EXISTS cigarettes 
                    (barcode INTEGER, itemName TEXT, stockOnHand INTEGER, purchases INTEGER, sales INTEGER)
                """)
        # Delete all rows
        c.execute('DELETE FROM cigarettes')
        
        # Reset stock, perchases and level to zero
        for x in range(len(info)):
            query = ("INSERT INTO cigarettes (barcode, itemName, stockOnHand, purchases, sales) VALUES (?, ?, ?, ?, ?)")
            c.execute(query, (info[x][0], info[x][1], 0, 0, 0))   
    
        #  Commit Changes to db
        conn.commit()
        # Close Connection to db
        conn.close()

def database_init():
    # Master stock file
        file = 'master stock sheet/Cigarette Master Stock Sheet.xls'
        # Read in excel
        data = pd.read_excel(file, header=0)
        # Date to list
        info = data.to_numpy().tolist()

        # Create a database or connect to one
        conn = sqlite3.connect('database/cigarettes.db')
        # Create cursor 
        c = conn.cursor()
        # Query DB
        # Create table when run for first time
        c.execute("""CREATE TABLE IF NOT EXISTS cigarettes 
                    (barcode INTEGER, itemName TEXT, stockOnHand INTEGER, purchases INTEGER, sales INTEGER)
                """)
        # Delete all rows
        c.execute('DELETE FROM cigarettes')
        
        # Reset stock, perchases and level to zero
        for x in range(len(info)):
            query = ("INSERT INTO cigarettes (barcode, itemName, stockOnHand, purchases, sales) VALUES (?, ?, ?, ?, ?)")
            c.execute(query, (info[x][0], info[x][1], 0, 0, 0))   
    
        #  Commit Changes to db
        conn.commit()
        # Close Connection to db
        conn.close()  

def restore():
    try:
        file = filedialog.askopenfilename(initialdir = "database/backup",title = "Select Database to Restore")
        file_backup = os.path.split(file)[1]
        file_name = file_backup[0:10] + '.db'
        newPath = shutil.copy(f'database/backup/{file_backup}', f'database/{file_name}')
        messagebox.showinfo('Backup Successful', f'Back up {file_backup[10:24]} restored successfuly')
    except:
        messagebox.showinfo('Backup NOT Successful', f'Back up {file_backup[10:24]} NOT restored successfuly')
    
