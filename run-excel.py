import win32com.client
import sys

def run_excel_macro(application, file_path, macro):
    # Get Excel application
    xl = win32com.client.Dispatch(application)
    
    # Open the file. Full path must be provided
    wb = xl.Workbooks.Open(file_path)
    
    # Choose the sheet to work with and do some input
    ws = wb.Worksheets('Sheet1')
    ws.Range('E5').Value = 1337
    
    # Run a macro defined in the Excel file
    xl.Application.Run(macro)
    
    # Save and quit the application, to not cause a process mess
    wb.Save()
    xl.Application.Quit()

    
if __name__ == '__main__':
    run_excel_macro('Excel.Application', sys.argv[1], 'dummy.xlsm!macro1')
