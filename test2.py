import xlrd
import xlwt

file = 'test.xls'
wb = xlrd.open_workbook(filename=file)  # 打开文件
num = [[0] * 24 for _ in range(16)]


def read_excel():
    sheet1 = wb.sheet_by_name('Results')  # 通过名字获取表格
    for i in range(16):
        for j in range(24):
            print(sheet1.cell_value(i * 24 + j + 1, 1), sheet1.cell_value(i * 24 + j + 1, 2))
            num[i][j] = sheet1.cell_value(i * 24 + j + 1, 2)


def write_excel():
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('sheet1', cell_overwrite_ok=True)
    for i in range(16):
        for j in range(24):
            print(j, i, num[i][j])
            sheet1.write(j + 2, i + 2, num[i][j])
    for i in range(16):
        sheet1.write(1, i + 2, chr(ord('A') + i))
    for j in range(24):
        sheet1.write(j + 2, 1, j + 1)

    f.save("help.xls")


if __name__ == '__main__':
    read_excel()
    write_excel()
