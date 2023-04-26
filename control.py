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


word_file_path = "min_template.docx"
minutas_file_path = "CONTROL/min_automation.xlsm"
pptx_file_path = "pptx_automation.xlsm"
contrato_forta = "CONTROL/CONTRATO/FORTALECEMOS/forta_contrato.xlsm"
contrato_mexti ="CONTROL/CONTRATO/MEXTI/mexti_contrato.xlsm"
contrato_constru= "CONTROL/CONTRATO/CONSTRUYENDO/constru_contrato.xlsm"

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

def open_constru_contrato():
    wb = xw.Book(contrato_constru)
    xw.apps[0].visible = True
    wb.save()
    wb.close()

def open_forta_contrato():
    wb = xw.Book(contrato_forta)
    xw.apps[0].visible = True
    wb.save()
    wb.close()

def open_mexti_contrato():
    wb = xw.Book(contrato_mexti)
    xw.apps[0].visible = True
    wb.save()
    wb.close()


#root = tkinter.Tk()
#root.title("PANEL DE CONTROL")

def button_callback():
    print("button pressed")

app = customtkinter.CTk()
app.title("my app")
app.geometry("400x150")

button = customtkinter.CTkButton(app, text="MINUTAS", command=open_minuta_file)
button.grid(row=0, column=0, padx=20, pady=20)
button = customtkinter.CTkButton(app, text="PRESENTACIONES", command=open_minuta_file)
button.grid(row=0, column=2, padx=20, pady=20)

app.mainloop()

