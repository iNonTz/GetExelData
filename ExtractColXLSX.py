# Program extracting first column
import xlrd
import os


def getYoutube(count, arr):
    for file_name in arr:
        if file_name == '.history' or '.py' in file_name:
            print('SKIP : ' + file_name)
        else:
            try:
                wb = xlrd.open_workbook(file_name)
                sheet = wb.sheet_by_index(3) ## Define Index sheet of file (First sheet is 0)
                # sheet.cell_value(0, 3)
                with open('Twitter.txt', 'a', encoding='utf-8') as file:
                    row_filename = str(file_name) + '\n'
                    file.write(row_filename)
                    for i in range(sheet.nrows):
                        row_info = str(sheet.cell_value(i, 3)) + '\n' ## Define Index cell of sheet to get (First col is 0)
                        if row_info == 'Video URL\n':
                            print('SKIP Row : ' + str(row_info))
                        else:
                            file.write(row_info)
                count+=1
                print('Done with : ' + str(file_name))
            except:
                print('Cannot parse row at file : ' + str(file_name))
    print('Success : ' + str(count) + ' files')

def getTwitter(count, arr):
    for file_name in arr:
        if file_name == '.history' or '.py' in file_name:
            print('SKIP : ' + file_name)
        else:
            try:
                wb = xlrd.open_workbook(file_name)
                sheet = wb.sheet_by_name("twitter") ## Define Index sheet of file (First sheet is 0)
                # sheet.cell_value(0, 3)
                with open('Twitter.txt', 'a', encoding='utf-8') as file:
                    row_filename = '\n' + str(file_name) + '\n'
                    file.write(row_filename)
                    for i in range(sheet.nrows):
                        if i <= 0:
                            i += 1
                        if type(sheet.cell_value(i, 13)) == float:
                            data = str(sheet.cell_value(i, 0)) + ' ' + str(sheet.cell_value(i, 13)) + '\n'
                            file.write(data) 
                        else:
                            data = str(sheet.cell_value(i, 0)) + ' ' + str(sheet.cell_value(i, 11)) + '\n'
                            file.write(data)
                count+=1
                print('Done with : ' + str(file_name))
            except:
                print('Cannot parse row at file : ' + str(file_name))
    print('Success : ' + str(count) + ' files')

if __name__ == "__main__":
    count = 0
    arr = os.listdir()
    print('Total Source : ' + str(len(arr) -1) + ' files')
    # getYoutube(count, arr)
    getTwitter(count, arr)