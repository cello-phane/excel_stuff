#________________EXCEL HEADER LINE_____________
#[_]|==|==|==|==|==|[_]  [_]|==|==|==|==|==|[_]
#[_]|__|__|__|__|__|[_]  [_]|__|__|__|__|__|[_]
#[_]|__|__|__|__|__|[_]  [_]|__|__|__|__|__|[_]
#[_]|__|__|__|__|__|[_]  [_]|__|__|__|__|__|[_]
#[_]|__|__|__|__|__|[_]  [_]|__|__|__|__|__|[_]
#[_]|__|__|__|__|__|[_]  [_]|__|__|__|__|__|[_]
#[_]|__|__|__|__|__|[_]  ^^^^^^^^^^^^^^^^^^^^^^

#[_]|==|==|==|==|==|[_]  ####every#nth#row######
#[_]|__|__|__|__|__|[_]  /_//_//_//_//_//_//_//_/
#[_]|__|__|__|__|__|[_]  /_//_//_//_//_//_//_//_/
#[_]|__|__|__|__|__|[_]  <<Cells move up and right
#[_]|__|__|__|__|__|[_]  /_//_//_//_//_//_//_//_/
#[_]|__|__|__|__|__|[_]  /_//_//_//_//_//_//_//_/
#[_]|__|__|__|__|__|[_]  /_//_//_//_//_//_//_//_/
#[_]|==|==|==|==|==|[_]  <<Next subpage of cells
#[_]|__|__|__|__|__|[_]  continues up and right,etc.

#When worksheet is print previewed,
#every cell page should be evenly paginated
#by number of rows and columns.
#After completion, a new file is generated(file_name + "_new.xls")
from helper_funcs import *
import os
from os import path
import xlrd
import xlwt
import sys
xlspath = input('Path of dir containing the .xls file: ')
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
writesheet = writebook.add_sheet(str(remove_ext(workbookfile)), cell_overwrite_ok=True)
writefile = xlspath + str(remove_ext(workbookfile)) + "_new" + workbookfile[workbookfile.rindex("."):]
ranges_= []
nrowsinput = ''
nrowsinput = input('Select rows to paginate(leave blank for all):')
if workbookfile.endswith('xls'):
    ranges_ = range_expand(input_r = str(nrowsinput), delim=',', first_n=0, last_n=len(xlsfile_)-1)
nth_row = int(input('Rows on each "page":'))
every_n_col = int(input('Columns on each "page":'))
divm_rows_uneven = divmod(len(ranges_),nth_row)
numberofpages = divm_rows_uneven[0]#the number of pages that can be divided evenly into sections of nth_row rows
unadvance_row = nth_row#to track the head of the current 
for iter, n in enumerate(ranges_[::nth_row*2]):
    if n is not ranges_[n::nth_row] and divm_rows_uneven[0] and n is not len(ranges_)-divm_rows_uneven[0]:
        for rownum in ranges_[n:n+nth_row]:
            row_values = [value if value else '' for value in xlsfile_[ranges_[rownum]]]
            next_rown_after=rownum+nth_row
            if next_rown_after < numberofpages*nth_row:
                row_values_next = [value if value else '' for value in xlsfile_[ranges_[next_rown_after]]] 
                row_values.extend(row_values_next)
                for col in range(every_n_col*2):
                    if rownum > nth_row:
                        writesheet.write(rownum-(unadvance_row*iter), col, str(row_values[col]))#write to header line of page
                    else:
                        writesheet.write(rownum, col, str(row_values[col]))
            else:##Last two pages can be subtracted from leftover divmod[1](remainder of rows)
                 ##to line up nicely 
                 ##[nth_row(left page)] | last_left_off - nth_row (right page)
                if len(ranges_) > rownum:
                    remainder_row = len(ranges_) - (nth_row+divm_rows_uneven[1])
                    last_iter_ = (numberofpages/2) * nth_row
                    #last_iter is the tail end of evenly divided pages
                    #which is half of initial cells * nth_rows
                    for rownum in ranges_[remainder_row:remainder_row+nth_row]:
                        #left page(remainder_row until the next nth_rows)
                        row_values = [value if value else '' for value in xlsfile_[ranges_[rownum]]]
                        for col in range(every_n_col):
                            writesheet.write(last_iter_, col, str(row_values[col]))
                        last_iter_+=1
                    for rownum in ranges_[remainder_row+nth_row:len(ranges_)-1]:
                        #right page(remainder_row+nth_rows until the last row number)
                        row_values = [value if value else '' for value in xlsfile_[ranges_[rownum]]]
                        for col in range(every_n_col):
                            writesheet.write(last_iter_-nth_row, col+every_n_col, str(row_values[col]))
                        last_iter_+=1#keep track manually of next row
                break
writebook.save(writefile)