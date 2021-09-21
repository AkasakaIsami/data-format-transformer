import xlwt
from configparser import ConfigParser

cp = ConfigParser()
cp.read("config.ini")

sys_name = cp.get('server', 'sys_name')

if __name__ == '__main__':
    a = ["SCD-IN", "SCD-SE"]
    book = xlwt.Workbook(encoding='utf-8')  # 创建Workbook，相当于创建Excel
    # 创建sheet，Sheet1为表的名字，cell_overwrite_ok为是否覆盖单元格
    sheet1 = book.add_sheet('Sheet1', cell_overwrite_ok=True)

    # 调整格式
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01
    style.alignment = alignment

    # 此为合并单元格，如果不需要自行删除
    # 起始行，终止行，起始列，终止列
    sheet1.write_merge(0, 1, 0, 0, "sys / sys-level  Metric", style)
    sheet1.write_merge(0, 0, 1, 2, "SCD")


    for i in range(0, len(a)):
        sheet1.write(1, i+1, a[i])

    sheet1.write(2, 0, sys_name)

    f = open("SCD.txt", encoding='utf-8')
    line = f.readline()
    sheet1.write(2, 1, line.split('[LEVEL3]')[1])

    line = f.readline()
    sheet1.write(2, 2, line.split('[LEVEL4]')[1])

    book.save("result.xls")
