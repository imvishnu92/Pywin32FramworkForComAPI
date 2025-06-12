import time
import pythoncom
import win32com.client

def create_excel_file(text):

    pythoncom.CoInitialize()

    excel = win32com.client.Dispatch("Excel.Application")

    try:

        try:
            excel.Visible = False
        except AttributeError as ve:
            print("Warning: Could not set Visible property:", ve)

        # Create and write to workbook
        workbook = excel.Workbooks.Add()
        sheet = workbook.Sheets(1)
        sheet.Cells(1, 1).Value = text

        filepath = r"C:\Users\ASUS\Documents\from_behave.xlsx"
        workbook.SaveAs(filepath)
        workbook.Close(SaveChanges=False)
        excel.Quit()

        return filepath

    finally:
        pythoncom.CoUninitialize()
