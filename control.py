#Importación de librerías
import tkinter as tk
import tkinter 
import os
import openpyxl
import xlwings as xw
import tkinter.ttk as ttk
import win32com.client
from tkinter import *
from ttkthemes import ThemedTk
import sv_ttk
import customtkinter

# Se nombran las variables a ocupar
word_file_path = "min_template.docx"
minutas_file_path = "CONTROL/min_automation.xlsm"
pptx_file_path = "pptx_automation.xlsm"
contrato_cliente = "CONTROL/CONTRATO/CLIENTE 1/contrato_cliente.xlsm"


# Se realizan las funciones que se van a ocupar

def open_minuta_file():
    wb = xw.Book(minutas_file_path)
    xw.apps[0].visible = True
    wb.save()
    wb.close()

def open_pptx_file():
    wb = xw.Book(pptx_file_path)
    xw.apps[0].visible = True
    wb.save()
    wb.close()

def open_word_file():
    os.startfile(word_file_path)

def open_contrato_cliente():
    wb = xw.Book(contrato_cliente)
    xw.apps[0].visible = True
    wb.save()
    wb.close()



app = customtkinter.CTk()
app.title("CONTROL")
app.geometry("230x250")
app.grid_columnconfigure(0, weight=1)

button = customtkinter.CTkButton(app, text="MINUTAS", command=open_minuta_file)
button.grid(row=0, column=0, padx=20, pady=20)
button = customtkinter.CTkButton(app, text="PRESENTACIONES", command=open_minuta_file)
button.grid(row=1, column=0, padx=20, pady=20)
button = customtkinter.CTkButton(app, text="CONTRATOS", command=open_contrato_cliente)
button.grid(row=2, column=0, padx=20, pady=20)
app.mainloop()

