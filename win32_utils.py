import os
import win32com.client
import pythoncom

def docx_handle(docx_file, update_toc=True, pdf_file=None):
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    doc = word.Documents.Open(os.path.abspath(docx_file))
    if update_toc:
        doc.TablesOfContents(1).Update()
    if pdf_file:
        doc.SaveAs(pdf_file, FileFormat=17)
    doc.Close(SaveChanges=True)
    word.Quit()
    return True
