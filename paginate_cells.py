from helper_funcs import open_excel, get_xls_length, containsNumber, remove_ext, col_to_n
import os
import xlwt
import sys

xlspath = sys.path[0] + '\\'
workbookfile = input('Excel File: ' + ' \nin this dir: ' + str(''.join([xlspath])))

while not workbookfile or not os.path.exists(os.path.join(xlspath, workbookfile)):
    workbookfile = input('Excel File: ' + ' \nin this dir: ' + str(''.join([xlspath])))

excelpath = os.path.join(xlspath, workbookfile)
writebook = xlwt.Workbook()
writesheet = writebook.add_sheet(remove_ext(workbookfile).strip('\\'), cell_overwrite_ok=True)
writefile = os.path.join(xlspath, remove_ext(workbookfile) + "_new.xls")
cellranges_ = []

worksheet_number = input('Worksheet #:')

if workbookfile.lower().endswith('xls'):
    if containsNumber(worksheet_number):
        xlslen = get_xls_length(excelpath, sheet_num=int(worksheet_number)-1)
        cellranges_ = generate_sequence(input_r="", delim=',', first_n=1, last_n=xlslen)
        xlsfile_ = open_excel(to_read=excelpath, sheet_num=int(worksheet_number)-1, ranges=cellranges_)
    else:
        xlslen = get_xls_length(excelpath, sheet_num=0)
        cellranges_ = generate_sequence(input_r="", delim=',', first_n=1, last_n=xlslen)
        xlsfile_ = open_excel(to_read=excelpath, sheet_num=0, ranges=cellranges_)

nth_rows = int(input('Number of rows on each subpage:'))
nth_cols = input('Columns on each subpage(Ex. a,b,c,D,E,F):')

pages = 2
divm_rows_uneven = divmod(len(cellranges_), int(nth_rows))
numberofpages = divm_rows_uneven[0]
new_row = 0
col_iter = 0
tail_start_row = 0

for iter, n in enumerate(cellranges_[::nth_rows * pages]):
    for rownum in cellranges_[n:n + nth_rows]:
        col_iter = 0
        try:
            row_values = xlsfile_[rownum]
        except:
            break
        next_rown_after_ = rownum + nth_rows
        if next_rown_after_ < numberofpages * nth_rows:
            row_values_r = xlsfile_[cellranges_[next_rown_after_]]
            row_values.extend(row_values_r)
        else:
            row_values_r = []
        for p in range(pages):
            for col in nth_cols.split(','):
                colnum_ = col_to_n(col)
                if p < pages / 2:
                    writesheet.write(new_row, col_iter, str(row_values[colnum_]))
                else:
                    if row_values_r:
                        writesheet.write(new_row, col_iter, str(row_values[colnum_ + len(row_values_r)]))
                col_iter += 1
        new_row += 1

diff_ = len(cellranges_) - nth_rows - divm_rows_uneven[1]
start_row = new_row - nth_rows

if diff_ != nth_rows:
    start_at = numberofpages * nth_rows
    for rownum in cellranges_[start_at:]:
        col_iter = len(nth_cols.split(','))
        try:
            row_values = xlsfile_[rownum]
        except:
            break
        for col in nth_cols.split(','):
            colnum_ = col_to_n(col)
            writesheet.write(start_row, col_iter, str(row_values[colnum_]))
            col_iter += 1
        start_row += 1

writebook.save(writefile)
