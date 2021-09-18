import re
import xlwt

if __name__ == '__main__':
    book = xlwt.Workbook(encoding='utf-8')  # 创建Workbook，相当于创建Excel
    # 创建sheet，Sheet1为表的名字，cell_overwrite_ok为是否覆盖单元格
    sheet1 = book.add_sheet('Sheet1', cell_overwrite_ok=True)

    f = open("transfer2.txt", encoding='utf-8')

    result=[];
    labels=[];
    a=[];

    #拿到每一个数据对
    for line in f.readlines():
        result.append(line.split(','))

    print(result)

    #将每一个数据对拆分为key和value
    for data in result[0]:
        a.append(re.findall(r"[A-Za-z]+\-?[A-Za-z]*",data));
        a.append(re.findall(r"\d+\.?\d*",data))

    #得到所有的标签
    for i in range(0,30,2):
        if a[i][0] not in labels:
            labels.append(a[i][0])
        if a[i][1] not in labels:
            labels.append(a[i][1])

    #有一个标签没出现过，补上去
    if "mail-service" not in labels:
        labels.append("mail-service")

    #将标签填上去
    for i in range(1, 8):
        sheet1.write(i, 0, labels[i - 1])
    for i in range(1, 8):
        sheet1.write(0, i, labels[i - 1])

    sheet1.write(0,0,"Service / Service")

    #填充上半部分
    for i in range(0, len(a), 2):
        print(a[i])
        row=labels.index(a[i][0]);
        col = labels.index(a[i][1]);
        if row>col:
            tmp=row;
            row=col;
            col=tmp;
        print('(',row,',',col,')')
        sheet1.write(row+1,col+1,a[i+1])


    #填充下半部分
    a=[];
    # 将每一个数据对拆分为key和value
    for data in result[1]:
        a.append(re.findall(r"[A-Za-z]+\-?[A-Za-z]*", data));
        a.append(re.findall(r"\d+\.?\d*", data))

    for i in range(0, len(a), 2):
        print(a[i])
        row=labels.index(a[i][0]);
        col = labels.index(a[i][1]);
        if row<col:
            tmp=row;
            row=col;
            col=tmp;
        print('(',row,',',col,')')
        sheet1.write(row+1,col+1,a[i+1])

    book.save("test.xls")

    f.close();