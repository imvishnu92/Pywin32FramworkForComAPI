from behave import *
import os
import win32com.client
import pythoncom

from excelWriter import create_excel_file


@given("Excel is available")
def step_given_excel_is_available(context):
    pass

@when('I write {text} to cell A1')
def step_write_text_to_excel(context, text):
     context.filepath = create_excel_file(text)

@then('The excel file should be saved')
def step_file_saved(context):
    assert os.path.exists(context.filepath)

