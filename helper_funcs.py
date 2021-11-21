import csv
import xlrd
def open_excel(to_read, sheet_num=0):
    def read_lines(workbook):
        sheet = workbook.sheet_by_index(sheet_num)
        for row in range(sheet.nrows):
            yield [sheet.cell(row, col).value for col in range(sheet.ncols)]
    try:
        workbook = xlrd.open_workbook(to_read)
        return [line for line in read_lines(workbook)]
    except (IOError, ValueError):
        print("Couldn't read from file %s." % (to_read))
        raise
def containsNumber(value):
    for character in value:
        if character.isdigit():
            return True
    return False
def remove_ext(ifile):
   if '.' in str(ifile)[-5:]:
       return ifile[:ifile.rindex(".")]
#Example:
#    print(range_expand(input_r='start..2,3:8,9..end',first_n=1,last_n=10))
#    > [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
#    returns 0 based values for spreadsheet xlrd library
def range_expand(input_r="", delim=",", first_n=1, last_n=10):
    _ranges = []
    if input_r and not input_r.isnumeric():
        comma_sep_str=[]
        if delim in input_r:
            comma_sep_str=input_r.split(delim)
        else:
            comma_sep_str.append(input_r)
        delim_find=0
        for range_substr in comma_sep_str:
            if ':' in range_substr or '..' in range_substr:
                if ':' in range_substr:
                    delim_find = str(range_substr).index(':')
                elif '..' in range_substr:
                    range_substr = str(range_substr).replace('..',':')
                    delim_find = str(range_substr).index(':')
                from_n = str(range_substr)[:delim_find]
                to_n = str(range_substr)[delim_find+1:]
                if from_n.isnumeric():
                    from_n = int(from_n)-1
                elif 'start' in from_n:
                    from_n = first_n-1
                if to_n.isnumeric():
                    to_n = int(to_n)
                elif 'end' in to_n:
                    to_n = last_n
                if int(to_n) < int(from_n):
                    temp = int(to_n)
                    to_n = int(from_n)
                    from_n = temp
                for n in range(int(from_n),int(to_n)):
                    _ranges.append(n)
            elif range_substr.isnumeric():
                _ranges.append(int(range_substr)-1)
        return list(sorted(set(_ranges)))
    elif not input_r:
        for n in range(first_n-1,last_n):
            _ranges.append(n)
        return list(sorted(set(_ranges)))
    elif input_r.isnumeric():
        _ranges.append(int(input_r)-1)
        return _ranges
