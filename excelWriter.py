import win32com.client
import os

def create_excel_file(text):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    workbook = excel.Workbooks.Add()
    sheet = workbook.Sheets(1)
    sheet.Cells(1, 1).Value = text

    file_path = r"C:\Users\runneradmin\Documents\excel_behave.xlsx"
    workbook.SaveAs(file_path)
    workbook.Close(False)
    excel.Quit()

    return file_path
📦