# -*- coding:utf-8 -*-

import xlwt
import re

#直接把数据全部贴进transfer.txt，只要一条数据不换行就行
if __name__ == '__main__':
    book = xlwt.Workbook(encoding='utf-8')  # 创建Workbook，相当于创建Excel
    # 创建sheet，Sheet1为表的名字，cell_overwrite_ok为是否覆盖单元格
    sheet1 = book.add_sheet('Sheet1', cell_overwrite_ok=True)
    sheet1.write(0,1,"读写比的平均值")
    sheet1.write(0,2,"最大的字段读写比")
    f = open("transfer.txt", encoding='gbk')
    # 使用readline()读文件
    num = 1;
    while True:
        line = f.readline()
        if line:
            if line.find('平均值') != -1:
                str=line.split("接口读取")
                str1=str[0]
                str2=str[1].split("接口更新")[0]
                sheet1.write(num,0,str1+str2)
                res=re.findall(r"\d+\.?\d*", line)
                sheet1.write(num,1,res[0])
                num = num + 1;
            elif line.find('最大的字段读写比')!=-1:
                str=line.split("两接口中的")
                sheet1.write(num,0,str[0])
                res = re.findall(r"\d+\.?\d*", line)
                sheet1.write(num,2,res[0])
                num = num + 1;
        else:
            break

    f.close()

    book.save("test.xls")