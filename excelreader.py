import pandas as pd
import openpyxl as op
from pandas import ExcelWriter
from pandas import ExcelFile
from tkinter import *
from tkinter.filedialog import asksaveasfilename, askopenfilename
import subprocess
from tkinter import font
import tkinter as tk 

df = pd.read_excel('File.xlsx', sheet_name='Sheet1')
thislist = [df['Program 1'], df['Program 2'], df['Program 3']]
filee='cmnd.py'
list_of_Code = df['Name']

wb = op.load_workbook("C:\\Users\\asus\\Downloads\\New folder\\File.xlsx")
sh = wb.active

k=0
column_position=5
while k< len(thislist):
    row_position = 2
    program_list=thislist[k]
    for i in list_of_Code.index:
        command = program_list[i]
        with open (filee,'w') as f:
            f.write(command)

        commandd="C:\\Users\\asus\\Downloads\\New folder\\cmnd.py"
        process = subprocess.Popen(commandd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        output, error = process.communicate()
        c = sh.cell(row=row_position, column=column_position)
        c.value = output
        row_position+=1
        print(output)
    column_position +=1
    k=k+1

wb.save("C:\\Users\\asus\\Downloads\\New folder\\File.xlsx")

