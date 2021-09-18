import xlwt
from configparser import ConfigParser

cp = ConfigParser()
cp.read("config.ini")

service_count = cp.getint('server', 'service_count')
service_pair_count = int(service_count * (service_count - 1) / 2)


def _get_services_pairs(service_pair_count):
    services_pairs = []
    f = open("transfer1.txt", encoding='utf-8')
    line = f.readline()
    temp1 = line.split(', ')
    for pair in temp1:
        temp2 = pair.split('=')
        services_pairs.append(temp2[0])
    if len(services_pairs) != service_pair_count:
        print(
            f"get service pairs wrong: {len(services_pairs)} {service_pair_count}")
        return None
    temp3 = services_pairs[0].split('{')
    services_pairs[0] = temp3[1]

    return services_pairs


def _get_metrics():
    DAP_FNs = []
    DAP_FCs = []
    f = open("transfer1.txt", encoding='utf-8')
    line = f.readline()
    temp1 = line.split(', ')
    for pair in temp1:
        temp2 = pair.split('=')
        DAP_FNs.append(temp2[1])
    temp3 = DAP_FNs[len(DAP_FNs)-1].split('}')
    DAP_FNs[len(DAP_FNs)-1] = temp3[0]

    line = f.readline()
    temp1 = line.split(', ')
    for pair in temp1:
        temp2 = pair.split('=')
        DAP_FCs.append(temp2[1])
    temp3 = DAP_FCs[len(DAP_FCs)-1].split('}')
    DAP_FCs[len(DAP_FCs)-1] = temp3[0]

 
    return [DAP_FNs, DAP_FCs]


if __name__ == '__main__':
    a = ["DAP-FN", "DAP-FC"]
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
    sheet1.write_merge(0, 0, 1, 2, "DAP")

    for i in range(0, len(a)):
        sheet1.write(1, i+1, a[i])

    service_pairs = _get_services_pairs(service_pair_count=service_pair_count)

    i = 0
    for i in range(0,len(service_pairs)):
        sheet1.write(i+2, 0, service_pairs[i])

    metrics = _get_metrics()

    DAP_FNs = metrics[0]
    DAP_FCs = metrics[1]


    for i in range(0, len(service_pairs)):
        sheet1.write(i+2, 1, DAP_FNs[i])
        sheet1.write(i+2, 2, DAP_FCs[i])

    book.save("result1.xls")
