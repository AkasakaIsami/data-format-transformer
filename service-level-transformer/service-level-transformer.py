import xlwt
from configparser import ConfigParser

cp = ConfigParser()
cp.read("config.ini")

service_count = cp.getint('server','service_count')

def _get_services(service_count):
    services = []
    f = open("transfer.txt", encoding='utf-8')
    for i in range(0, service_count):
        line = f.readline()
        if line:
            service = line.split('"')
            services.append(service[1])
        else:
            break
    return services

def _get_metrics(service_count):
    DAH_FNs = []
    DAH_FCs = []
    PIH_FNs = []
    PIH_FCs = []
    SIH_FNs = []
    SIH_FCs = []
    f = open("transfer.txt", encoding='utf-8')
    for i in range(0, 6*service_count):
        line = f.readline()
        if line and i < service_count:
            temp1 = line.split(': ')
            temp2 = temp1[1].split(',')
            DAH_FNs.append(temp2[0])
        elif line and i < 2*service_count:
            temp1 = line.split(': ')
            temp2 = temp1[1].split('\n')
            DAH_FCs.append(temp2[0])
        elif line and i < 3*service_count:
            temp1 = line.split('\t')
            temp2 = temp1[1].split('\n')
            PIH_FNs.append(temp2[0])
        elif line and i < 4*service_count:
            temp1 = line.split('\t')
            temp2 = temp1[1].split('\n')
            PIH_FCs.append(temp2[0])
        elif line and i < 5*service_count:
            temp1 = line.split('\t')
            temp2 = temp1[1].split('\n')
            SIH_FNs.append(temp2[0])
        elif line and i < 6*service_count:
            temp1 = line.split('\t')
            temp2 = temp1[1].split('\n')
            SIH_FCs.append(temp2[0])
        else:
            break
    return [DAH_FNs, DAH_FCs, PIH_FNs, PIH_FCs, SIH_FNs, SIH_FCs]

if __name__ == '__main__':
    a = ["DAH-FN", "DAH-FC", "PIH-FN", "PIH-FC", "SIH-FN", "SIH-FC"]
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
    sheet1.write_merge(0, 1, 0, 0, "Service / Servcie-level  Metric", style)
    sheet1.write_merge(0, 0, 1, 2, "DAH")
    sheet1.write_merge(0, 0, 3, 4, "PIH")
    sheet1.write_merge(0, 0, 5, 6, "SIH")
    for i in range(0, len(a)):
        sheet1.write(1, i+1, a[i])

    services = _get_services(service_count=service_count)
    for i in range(0, len(services)):
        sheet1.write(i+2, 0, services[i])

    metrics = _get_metrics(service_count=service_count)

    DAH_FNs = metrics[0]
    DAH_FCs = metrics[1]
    PIH_FNs = metrics[2]
    PIH_FCs = metrics[3]
    SIH_FNs = metrics[4]
    SIH_FCs = metrics[5]

    for i in range(0, len(services)):
        sheet1.write(i+2, 1, DAH_FNs[i])
        sheet1.write(i+2, 2, DAH_FCs[i])
        sheet1.write(i+2, 3, PIH_FNs[i])
        sheet1.write(i+2, 4, PIH_FCs[i])
        sheet1.write(i+2, 5, SIH_FNs[i])
        sheet1.write(i+2, 6, SIH_FCs[i])

    book.save("result.xls")
