from pathlib import Path  # core library
import win32com.client as win32  # pip install pywin32
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl
import smtplib




def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", ".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


def main():
    
    # Conexión a Excel
    wb = xw.Book.caller()
    current_dir = Path(__file__).parent
    sht_clientes = wb.sheets["PANEL"]
    sht_panel = wb.sheets["PANEL"]
    folder_name = sht_clientes.range("C13").value

    # Path settings
    context = sht_panel.range("B2").options(dict, expand="table", numbers=int).value
    presta_name = context["NOMBRE_PRESTA"]
    if presta_name == "CONSTRUYENDO EQUIPOS HRM S.A. DE C.V.":
        template_path = current_dir / "min_template_constru.docx"
    elif presta_name == "FORTALECEMOS EMPRESAS HRM S.A. DE C.V..docx":
        template_path = current_dir / "min_template_forta.docx"
    else:
        template_path = current_dir / "min_template_mexti.docx"

    
   

    # Initialize template
    doc = DocxTemplate(str(template_path))



    # -- Render & Save Word Document
    output_name = f"C:\\Users\\Bryan Antonio Polito\\OneDrive - HRM\\PROYECTOS\\2023\\{folder_name}\\ARCHIVOS INTERNOS\\MINUTAS\\minuta_{context['NOMBRE_CLIENTE']}_{context['FECHA']}.docx"


    doc.render(context)
    doc.save(output_name)

    # Convertir a PDF
    convert_to_pdf(str(output_name))


   
    # Mensaje de listo
    show_msgbox = wb.macro("Module1.ShowMsgBox")
    show_msgbox("Listo!")


if __name__ == "__main__":
    xw.Book("min_automation.xlsm").set_mock_caller()
    main()

