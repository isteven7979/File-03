import xlwings as xw

# Connect to the work book Sample1.xlsx and Sample2.xlsx in the same directory
wb1 = xw.Book('Sample1.xlsx')

# Connect to the main workbook to write the result
main_wb = xw.Book('Main.xlsx')

# Connect to the sheet main of main workbook
main_sh = main_wb.sheets('Main')

# # Read all the Sheet name of workbook Sample1.xlsx 
# into the Cell A1 of Main.xlsx
# main_sh.range('A1').value = [sheet.name for sheet in wb1.sheets]

# # In case to read all the sheet name of a workbook many times, 
# it should be used by function
# def read_all_sheet_name(workbook):
#     return [sheet.name for sheet in workbook.sheets]

# main_sh.range('A2').value = read_all_sheet_name(wb1)

# Delete all the sheets of Sample1.xlsx except the Sheets ('Main')
# for sh in wb1.sheets:
#     if sh.name != 'Main':
#         sh.delete()

# In case to delete all the sheet name of a workbook many times, 
# it should be used by function
def delete_sheets_except(workbook, sheetisnotdelete):
    for sh in workbook.sheets:
        if sh.name != sheetisnotdelete:
            sh.delete()
           
delete_sheets_except(main_wb,"Main")

