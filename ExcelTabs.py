import os
import glob
import sys
import xlrd, xlwt
import time
# from sets import Set

allowedChars = set('0123456789')

def find_files_in_dir(dir, ext):
    os.chdir(dir)
    result = [i for i in glob.glob('*.{}'.format(ext))]
    return result

def process_file(inputfile):
    xl = xlrd.open_workbook(inputfile, on_demand = True, encoding_override="cp1251")
    directory = os.path.dirname(os.path.abspath(inputfile))

    for sheet in xl.sheets():
        # print(sheet.name)
        if not (set(sheet.name).issubset(allowedChars)):
            continue
        newExcelFile = xlwt.Workbook('cp1251')
        newSheet = newExcelFile.add_sheet(sheet.name)

        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                if row_index == 0:
                    continue
                newSheet.write(row_index, col_index, sheet.cell(row_index,col_index).value)

        newExcelFile.save(os.path.join(directory, sheet.name + '.xls'))

    xl.release_resources()
    del xl

if __name__ == "__main__":
    start_time = time.clock()
    path = ""
    if len(sys.argv) > 1:
        path = sys.argv[1]

    if path == "":
        print("Excel file must be specified as first parameter")
    else:
        if os.path.isfile(path):
            process_file(path)
            print(time.clock() - start_time, "seconds")
        elif os.path.isdir(path):
            print("Excel file must be specified as first parameter")
