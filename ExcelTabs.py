import os
import glob
import sys
import xlrd, xlwt
# import time
allowedChars = set('0123456789')

# alignment
al = xlwt.Alignment()
al.horz = xlwt.Alignment.HORZ_CENTER
al.vert = xlwt.Alignment.VERT_CENTER
al.wrap = True

# fonts
fnt = xlwt.Font()
fnt.bold = True
fnt.height = 8*20

fnt_bold = xlwt.Font()
fnt_bold.bold = True

# borders
borders = xlwt.Borders()
borders.top = 6 #xlwt.Borders.double
borders.bottom = 1#xlwt.Borders.thin

# styles
borders_style = xlwt.XFStyle()
borders_style.borders = borders
borders_style.font = fnt
borders_style.alignment = al

fnt_style = xlwt.XFStyle()
fnt_style.font = fnt_bold

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
                if row_index < 5 and col_index == 4:
                    newSheet.write(row_index, col_index, sheet.cell(row_index,col_index).value, fnt_style)
                elif row_index == 5:
                    newSheet.write(row_index, col_index, sheet.cell(row_index,col_index).value, borders_style)
                else:
                    newSheet.write(row_index, col_index, sheet.cell(row_index,col_index).value)
        newSheet.row(5).set_style(borders_style)
        newSheet.row(5).height_mismatch = 1
        newSheet.row(5).height = 650
        newSheet.col(0).width = 256*2
        newSheet.col(1).width = 256*10
        newSheet.col(2).width = 256*18
        newSheet.col(3).width = 256*18
        newSheet.col(6).width = 256*11
        newExcelFile.save(os.path.join(directory, sheet.name + '.xls'))

    xl.release_resources()
    del xl

if __name__ == "__main__":
    # start_time = time.clock()
    path = ""
    if len(sys.argv) > 1:
        path = sys.argv[1]

    if path == "":
        print("Excel file must be specified as first parameter")
    else:
        if os.path.isfile(path):
            process_file(path)
            # print(time.clock() - start_time, "seconds")
        elif os.path.isdir(path):
            print("Excel file must be specified as first parameter")
