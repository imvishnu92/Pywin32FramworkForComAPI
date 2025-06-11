from behave import *
import os
import win32com.client
import pythoncom

@given("Excel is available")
def step_given_excel_is_available(context):
    try:
        pythoncom.CoInitialize()  # Safe for threaded environments
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.Quit()
        context.excel_available = True
    except Exception as e:
        context.excel_available = False
        raise RuntimeError("Microsoft Excel is not installed or not registered as a COM server.") from e

