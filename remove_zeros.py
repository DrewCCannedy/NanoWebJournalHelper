"""Removes sheets with a value of $0."""
import sys
import glob
import os.path
import win32com.client as win32
import decimal


def remove_zeros(date):
    """Pass a folder of Excel Files."""
    output_string = ""
    books = 0
    sheet_count = 0
    deleted_sheets = 0
    deleted_books = 0
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = True
    path = "C:\\Journals\\" + date
    print("Looking for excel folders in {}".format(path))
    for filename in glob.glob(path+'\\*'):
        wb = xl.Workbooks.Open(filename)
        books += 1
        sheets = wb.Sheets
        for sheet in sheets:
            sheet_count += 1
            if sheet.Cells(2, 6).Value == 0:
                if(len(sheets) == 1):
                    wb.Close(False)
                    os.remove(filename)
                    deleted_books += 1
                else:
                    sheet.Delete()
                    deleted_sheets += 1
                    wb.Save()
    xl.Application.Quit()
    output_string = """Went through:\n\t{} Workbooks\n\t{} Sheets\nDeleted:\n\t{} Workbooks
    \t{} Sheets\n""".format(books, sheet_count, deleted_books, deleted_sheets)
    if books == 0:
        output_string += "\nThis app did not go through any folders."
        output_string += "\nMake sure the Journals Folder is in 'Local Disk (C)'."
        output_string += "\nAlso make sure the folder only contains Excel files."
    return output_string

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("ERROR: No folder name entered.")
    elif len(sys.argv) > 2:
        print("ERROR: Too many arguments. Please only enter a folder name")
    else:
        print(remove_zeros(sys.argv[1]))
