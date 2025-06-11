Feature: Excel Automation Using pywin32
  Scenario: Write data to Excel file
    Given Excel is available
    When I write "Hello from behave" to cell A1
    Then the excel file should be saved