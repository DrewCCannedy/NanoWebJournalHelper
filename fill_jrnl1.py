"""Fills out a template in excel for entering Journal entries."""
import sys
import glob
import os
import win32com.client as win32

class Journal_Line:
    def __init__(self, cost_center, ammount, description):
        self.cost_center = cost_center
        self.ammount = ammount
        self.description = description


def write_template(date):
    """Pass a folder of Excel Files."""
    journal_lines = []

    temp = date.split('-')
    month = number_to_date(int(temp[1]))
    year = temp[0]
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = True
    jrnl1path = 'C:\\Journals\\JRNL1\\Copy.xlsx'
    try:
        jrnl1 = xl.Workbooks.Open(jrnl1path)
    except:
        return "Error finding Copy file. Make sure it is in the JRNL1 folder\nAlso Make sure the Journals Folder is in 'Local Disk (C)'\n"
    total_ammount = 0
    current_line = 1
    path = 'C:\\Journals\\' + date
    print("Looking for folders at {}".format(path))
    for filename in glob.glob(path+'\\*'):
        wb = xl.Workbooks.Open(filename)
        sheets = wb.Sheets
        ammount = 0
        cc = ""
        cc_exists = False
        for sheet in sheets:
            for journal_line in journal_lines:
                if cc == journal_line.cost_center:
                    cc_exists = True
            temp = sheet.Cells(2,1).Value
            temp = temp.split(' ')
            name = temp[1]
            cc = sheet.Cells(1,6).Value
            ammount = sheet.Cells(2,6).Value
            description = "{}, {} {}".format(name, month, year)
            if not cc_exists:
                journal_lines.append(Journal_Line(cc, ammount, description))
            else: 
                for journal_line in journal_lines:
                    if journal_line.cost_center == cc:
                        journal_line.ammount += ammount
            cc_exists = False
    for journal_line in journal_lines:
        jrnl1.Sheets(1).Cells(current_line, 1).Value = 'DAL01'
        jrnl1.Sheets(1).Cells(current_line, 3).Value = 'ACTUALS'
        jrnl1.Sheets(1).Cells(current_line, 5).Value = 67017
        jrnl1.Sheets(1).Cells(current_line, 7).Value = journal_line.cost_center
        jrnl1.Sheets(1).Cells(current_line, 9).Value = journal_line.ammount
        jrnl1.Sheets(1).Cells(current_line, 11).Value = journal_line.description
        current_line += 1
        total_ammount += journal_line.ammount

    jrnl1.Sheets(1).Cells(current_line, 1).Value = 'DAL01'
    jrnl1.Sheets(1).Cells(current_line, 3).Value = 'ACTUALS'
    jrnl1.Sheets(1).Cells(current_line, 5).Value = 48101
    jrnl1.Sheets(1).Cells(current_line, 7).Value = 32639001
    jrnl1.Sheets(1).Cells(current_line, 9).Value = -total_ammount
    info = "{} {}".format(month, year)
    jrnl1.Sheets(1).Cells(current_line, 11).Value = info
    jrnl1.Save()
    xl.Application.Quit()
    return "Monthly Report Created!"
    

def number_to_date(num):
    if num == 1:
        return "January"
    elif num == 2:
        return "February"
    elif num == 3:
        return "March"
    elif num == 4:
        return "April"
    elif num == 5:
        return "May"
    elif num == 6:
        return "June"
    elif num == 7:
        return "July"
    elif num == 8:
        return "August"
    elif num == 9:
        return "September"
    elif num == 10:
        return "October"
    elif num == 11:
        return "November"
    elif num == 12:
        return "December"
    else:
        return ""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("ERROR: No folder name entered.")
    elif len(sys.argv) > 2:
        print("ERROR: Too many arguments.")
    else:
        write_template(sys.argv[1])
