
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


#pip install xlrd
#pip install xlwt
#python paginate_cells.py
#input file name inside this directory or change xlspath
#input worksheet number and row and column sizes;all of which are supposed to be non-zero based

#When worksheet is print previewed,
#every cell page should be evenly paginated
#by number of rows and columns.
#After completion, a new file is generated(file_name + "_new.xls")
