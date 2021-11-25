from helper_funcs import *
import os
from os import path
import xlrd
import xlwt
import sys
xlspath = sys.path[0] + '\\'
workbookfile = ''
while not workbookfile or not path.exists(''.join([xlspath,workbookfile])):
    workbookfile = input('Excel File: ' + ' \nin this dir: ' + str(''.join([xlspath])))
excelpath=str(xlspath)+str(workbookfile)        
worksheet_number = input('Worksheet #:')
if containsNumber(worksheet_number):
    xlsfile_ = open_excel(to_read=excelpath,sheet_num=int(worksheet_number)-1)    
else:
    xlsfile_ = open_excel(to_read=excelpath)
writebook = xlwt.Workbook()
writesheet = writebook.add_sheet(str(remove_ext(workbookfile)).strip('\\'), cell_overwrite_ok=True)
writefile = xlspath + str(remove_ext(workbookfile)) + "_new.xls"
ranges_= []
nrowsinput = input('Select rows to paginate(leave blank for all):')
if workbookfile.lower().endswith('xls'):
    ranges_ = range_expand(input_r = str(nrowsinput), delim=',', first_n=0, last_n=len(xlsfile_))
nth_rows = int(input('Number of rows on each subpage:'))
nth_cols = input('Columns on each subpage(Ex. a,b,c,D,E,F):')
pages = 2
divm_rows_uneven = divmod(len(ranges_),int(nth_rows))
numberofpages = divm_rows_uneven[0]#the number of pages that can be divided evenly into sections of nth_row rows
new_row = 0
col_iter = 0
for iter, n in enumerate(ranges_[::nth_rows*pages]):
    if divm_rows_uneven[1] > 0:
        if iter > numberofpages-1:
            break
    for rownum in ranges_[n:n+nth_rows]:
        col_iter = 0
        row_values = [value for value in xlsfile_[rownum]]
        next_rown_after_=rownum+nth_rows
        if next_rown_after_ < numberofpages*nth_rows:
            row_values_r = [value for value in xlsfile_[ranges_[next_rown_after_]]] 
            row_values.extend(row_values_r) 
        else:
            row_values_r = []
        for p in range(pages):
            for col in nth_cols.split(','):
                colnum_ = int(col_to_n(str(col)))
                if p is not pages/2:
                    writesheet.write(new_row, col_iter, str(row_values[colnum_]))
                else:
                    if row_values_r:
                        writesheet.write(new_row, col_iter, str(row_values[colnum_+len(row_values_r)]))
                col_iter+=1
        new_row+=1
tail_start_row = new_row-nth_rows
tail_col_start = len(nth_cols.split(','))
if divm_rows_uneven[1]:
    if divm_rows_uneven[1] < nth_rows:
        for enum,n in enumerate(ranges_[nth_rows*numberofpages:]):
            col_tail = tail_col_start
            if enum <= nth_rows:         
                row_values = [value for value in xlsfile_[ranges_[n]]]
                for col in nth_cols.split(','):
                    colnum_ = int(col_to_n(str(col)))
                    writesheet.write(tail_start_row, col_tail, str(row_values[colnum_]))
                    col_tail+=1
            else:
                break
            tail_start_row+=1
writebook.save(writefile)
