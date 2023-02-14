import os
import glob
import csv
import time
from xlsxwriter.workbook import Workbook
import xlsxwriter as xw


input_path = r'U:\projets\9_Autres\csv2excel\input_csv'
output_path = r'U:\projets\9_Autres\csv2excel\output_xlsx'


for csvfile in glob.glob(os.path.join(input_path, '*.csv')):
    print('converting file: ', csvfile)
    workbook = Workbook(os.path.join(output_path, os.path.basename(csvfile)[:-4] + '.xlsx'))
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        header = None
        for r, row in enumerate(reader):
            if r == 0:
                header = row
                continue
            for c, col in enumerate(row):
                worksheet.write(r, c, col)

    # write header
    cell_range = xw.utility.xl_range(0, 0, r, c)
    header = [{'header': di} for di in header]
    worksheet.add_table(cell_range, {'header_row': True, 'columns': header})

    workbook.close()
    os.remove(csvfile)
    
print('done !')
time.sleep(5)