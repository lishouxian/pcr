import xlrd
import xlwt

file = '0715.xls'
wb = xlrd.open_workbook(filename=file)  # 打开文件
num = [[0] * 16 for _ in range(24)]
style1 = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
style2 = xlwt.easyxf('pattern: pattern solid, fore_colour green;')

a, b = 35, 10


def read_excel():
    sheet1 = wb.sheet_by_name('Results')  # 通过名字获取表格
    for i in range(16):
        for j in range(24):
            print(sheet1.cell_value(i * 24 + j + 1 + a, 1 + b), sheet1.cell_value(i * 24 + j + 1 + a, 2 + b))
            num[j][i] = sheet1.cell_value(i * 24 + j + 1 + a, 2 + b)


def write_excel():
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('sheet1', cell_overwrite_ok=True)
    for i in range(24):
        for j in range(16):
            sheet1.write(j + 2, i + 2, num[i][j])
    for j in range(16):
        sheet1.write(j + 2, 1, chr(ord('A') + j), style2)

    for j in range(24):
        sheet1.write(1, j + 2, j + 1, style2)

    for i in range(24):
        for j in range(0, 15, 3):
            if isinstance(num[i][j], float) and isinstance(num[i][j + 1], float) and isinstance(num[i][j + 2], float):
                a = max(num[i][j], num[i][j + 1], num[i][j + 2])
                c = min(num[i][j], num[i][j + 1], num[i][j + 2])
                b = num[i][j] + num[i][j + 1] + num[i][j + 2] - a - c
                if a - c <= 0.5:
                    for k in range(3):
                        sheet1.write(j + 2 + k, i + 2, num[i][j + k])
                else:
                    for k in range(3):
                        if num[i][j + k] == a:
                            if a - b > 0.5 or a - b > b - c:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k], style1)
                            else:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k])
                        if num[i][j + k] == c:
                            if b - c > 0.5 or a - b < b - c:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k], style1)
                            else:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k])
                        if num[i][j + k] == b:
                            if a - b > 0.5 and b - c > 0.5:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k], style1)
                            else:
                                sheet1.write(j + 2 + k, i + 2, num[i][j + k])
            else:
                for k in range(3):
                    sheet1.write(j + 2 + k, i + 2, num[i][j + k], style1)
    f.save("help0715.xls")


if __name__ == '__main__':
    read_excel()
    write_excel()
