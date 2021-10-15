from pathlib import Path  # Standard Python Module
import time  # Standard Python Module
import xlwings as xw  # pip install xlwings

SOURCE_DIR = 'Files'

excel_files = list(Path(SOURCE_DIR).glob('*.xlsx'))
combined_wb = xw.Book()
t = time.localtime()
timestamp = time.strftime('%Y-%m-%d_%H%M', t)

for excel_file in excel_files:
    wb = xw.Book(excel_file)
    for sheet in wb.sheets:
        sheet.copy(after=combined_wb.sheets[0])
    wb.close()

combined_wb.sheets[0].delete()
combined_wb.save(f'all_worksheets_{timestamp}.xlsx')
if len(combined_wb.app.books) == 1:
    combined_wb.app.quit()
else:
    combined_wb.close()
