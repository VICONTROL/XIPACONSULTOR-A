from pathlib import Path  # core library
import win32com.client as win32  # pip install pywin32
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl

# -- Documentation:
# python-docx-template: https://docxtpl.readthedocs.io/en/latest/




def convert_to_pdf(doc):
    """Convert given word document to pdf"""
    word = win32.DispatchEx("Word.Application")
    new_name = doc.replace(".docx", ".pdf")
    worddoc = word.Documents.Open(doc)
    worddoc.SaveAs(new_name, FileFormat=17)
    worddoc.Close()
    return None


def main():
    # Path settings
    current_dir = Path(__file__).parent
    template_path = current_dir / "Contrato_template.docx"

    # Conection to Excel
    wb = xw.Book.caller()
    sht_panel = wb.sheets["PANEL"]
    sht_sales = wb.sheets["Sales"]
    context = sht_panel.range("B2").options(dict, expand="table", numbers=int).value


    # Initialize template
    doc = DocxTemplate(str(template_path))



    # -- Render & Save Word Document
    output_name = current_dir / f'Contrato_{context["NOMBRE_CLIENTE"]}.docx'
    doc.render(context)
    doc.save(output_name)

    # -- Convert to PDF [OPTIONAL]
    convert_to_pdf(str(output_name))

    # -- Show Message Box [OPTIONAL]
    show_msgbox = wb.macro("Module1.ShowMsgBox")
    show_msgbox("Listo!")


if __name__ == "__main__":
    xw.Book("forta_contrato.xlsm").set_mock_caller()
    main()
