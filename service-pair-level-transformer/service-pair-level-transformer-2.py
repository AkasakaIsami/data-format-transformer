import xlwt
from configparser import ConfigParser

cp = ConfigParser()
cp.read("config.ini")

service_count = cp.getint('server', 'service_count')
service_pair_count = int(service_count * (service_count - 1) / 2)


def _get_services_pairs(service_pair_count):
    services_pairs = []
    f = open("transfer2.txt", encoding='utf-8')
    for i in range(0, service_pair_count):
        line = f.readline()
        if line:
            services_pair = line.split('\t')
            services_pairs.append(services_pair[0])
        else:
            break
    return services_pairs


def _get_metrics(service_pair_count):
    PIP_FNs = []
    PIP_FCs = []
    SIP_FNs = []
    SIP_FCs = []

    f = open("transfer2.txt", encoding='utf-8')
    for i in range(0, 4*service_pair_count):
        line = f.readline()
        if line and i < service_pair_count:
            temp1 = line.split('\t')
            PIP_FNs.append(float(temp1[1]))
        elif line and i < 2*service_pair_count:
            temp1 = line.split('\t')
            PIP_FCs.append(float(temp1[1]))
        elif line and i < 3*service_pair_count:
            temp1 = line.split('\t')
            SIP_FNs.append(float(temp1[1]))
        elif line and i < 4*service_pair_count:
            temp1 = line.split('\t')
            SIP_FCs.append(float(temp1[1]))
        else:
            break

    return [PIP_FNs, PIP_FCs, SIP_FNs, SIP_FCs]

if __name__ == '__main__':
    a = ["PIP-FN", "PIP-FC",
         "SIP-FN", "SIP-FC"]
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
    sheet1.write_merge(
        0, 1, 0, 0, "Service pair / Servcie-pair-level  Metric", style)
    sheet1.write_merge(0, 0, 1, 2, "PIP")
    sheet1.write_merge(0, 0, 3, 4, "SIP")
    for i in range(0, len(a)):
        sheet1.write(1, i+1, a[i])

    service_pairs = _get_services_pairs(service_pair_count=service_pair_count)

    i = 0
    for i in range(0, len(service_pairs)):
        sheet1.write(i+2, 0, service_pairs[i])

    metrics = _get_metrics(service_pair_count=service_pair_count)


    PIP_FNs = metrics[0]
    PIP_FCs = metrics[1]
    SIP_FNs = metrics[2]
    SIP_FCs = metrics[3]


    for i in range(0, len(service_pairs)):
        sheet1.write(i+2, 1, PIP_FNs[i])
        sheet1.write(i+2, 2, PIP_FCs[i])
        sheet1.write(i+2, 3, SIP_FNs[i])
        sheet1.write(i+2, 4, SIP_FCs[i])

    book.save("result2.xls")
