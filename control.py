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


word_file_path = "min_template.docx"
minutas_file_path = "MINUTAS/min_automation.xlsm"
pptx_file_path = "pptx_automation.xlsm"
contrato_forta = "MINUTAS/CONTRATO/FORTALECEMOS/forta_contrato.xlsm"
contrato_mexti ="MINUTAS/CONTRATO/MEXTI/mexti_contrato.xlsm"
contrato_constru= "MINUTAS/CONTRATO/CONSTRUYENDO/constru_contrato.xlsm"

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


root = tkinter.Tk()
root.title("PANEL DE CONTROL")



# Crear una imagen
img = tk.PhotoImage(file="MINUTAS/IMAGENES/BUSINESS.png")

# Crear un label para mostrar la imagen
img_label = tk.Label(root, image=img)
img_label.pack(fill='both', expand=True)



root.geometry("800x450+100+100")
menubar = tk.Menu(root)
menubar.config(font=("Arial", 20))
filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label="Plantilla Word", command=open_word_file)
filemenu.add_command(label="Contrato Construyendo", command=open_constru_contrato)
filemenu.add_command(label="Contrato Fortalecemos", command=open_forta_contrato)
filemenu.add_command(label="Contrato Mexti", command=open_mexti_contrato)
filemenu.add_command(label="Panel Minuta", command=open_minuta_file)
filemenu.add_command(label="Panel PPTX", command=open_pptx_file)
menubar.add_cascade(label="MENÚ", menu=filemenu)
root.config(menu=menubar)


title_bar = tk.Frame(root, bg='white', relief='raised')
title_bar.place(relx=0, rely=0, relwidth=1, relheight=0.1)
title_bar.bind('<B1-Motion>', lambda e: root.geometry(f'+{e.x_root}+{e.y_root}'))
# Crear un label para mostrar el texto
text_label = tk.Label(title_bar, text="BIENVENIDO, UTILIZA EL MENÚ PARA DESPLAZARTE ENTRE LAS OPCIONES", font=("Arial", 16))
text_label.pack()

sv_ttk.set_theme("dark")
root.mainloop()
