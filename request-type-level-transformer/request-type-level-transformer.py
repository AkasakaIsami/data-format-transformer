import xlwt
from configparser import ConfigParser

cp = ConfigParser()
cp.read("config.ini")

request_count = cp.getint('server', 'request_count')


def _get_requests(request_count):
    requests = []
    pairs = []
    f = open("PCC.txt", encoding='utf-8')
    for i in range(0, request_count):
        line = f.readline()
        if line:
            pair = line.split('\t')
            pairs.append(pair[0])
            requests.append(pair[0].split('.')[1])
        else:
            break
    return [requests, pairs]


def _get_PCC(request_count):
    PCC = []
    f = open("PCC.txt", encoding='utf-8')
    for i in range(0, request_count):
        line = f.readline()
        if line and i < request_count:
            temp1 = line.split('\t')
            PCC.append(temp1[1])
        else:
            break
    return PCC


def _get_LIC(requests, request_count):
    LIC = [-1]*request_count
    f = open("LIC.txt", encoding='utf-8')
    for i in range(0, request_count):
        line = f.readline()
        if line and i < request_count:
            temp = line.split('\t')
            op = temp[0]
            value = temp[1].split('、')[0].split('[LENGTH]')[1]
            LIC[requests.index(op)] = value
        else:
            break
    return LIC


def _get_RIN(requests, request_count):
    RIN = [-1]*request_count
    f = open("RIN.txt", encoding='utf-8')
    for i in range(0, request_count):
        line = f.readline()
        if line and i < request_count:
            temp = line.split('\t')
            op = temp[0]
            value = temp[1].split('[COUNT]')[1]
            RIN[requests.index(op)] = value
        else:
            break
    return RIN


def _get_RCD(requests, request_count):
    RCD_INs = [-1]*request_count
    RCD_SEs = [-1]*request_count
    f = open("RCD.txt", encoding='utf-8')
    for i in range(0, request_count):
        line = f.readline()
        if line and i < request_count:
            temp = line.split('\t')
            op = temp[0]
            data = temp[1]

            RCD_IN = set()
            RCD_SE = set()
            paths = data.split("[PATHID]")
            for i in range(1, len(paths)):
                path = paths[i]
                temp = path.split("[LEVEL1]")[1]
                temp1 = temp.split('\n')[0]
                print(temp)

                RCD_IN.add(temp1.split("[LEVEL2]")[0])
                RCD_SE.add(temp1.split("[LEVEL2]")[1])
                # RCD_IN = RCD_IN + temp.split("[LEVEL2]")[0]
                # RCD_SE = RCD_SE + temp.split("[LEVEL2]")[1]

            RCD_INs[requests.index(op)] = RCD_IN
            RCD_SEs[requests.index(op)] = RCD_SE
        else:
            break
    return [RCD_INs, RCD_SEs]


if __name__ == '__main__':
    a = ["RCD-IN", "RCD-SE"]
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
    sheet1.write_merge(0, 1, 0, 0, "op / op-level  Metric", style)
    sheet1.write_merge(0, 1, 1, 1, "service")
    sheet1.write_merge(0, 1, 2, 2, "PCC")
    sheet1.write_merge(0, 0, 3, 4, "RCD")
    sheet1.write_merge(0, 1, 5, 5, "RIN")
    sheet1.write_merge(0, 1, 6, 6, "LIC")

    for i in range(0, len(a)):
        sheet1.write(1, i+3, a[i])

    result = _get_requests(request_count=request_count)
    reqs = result[0]
    pairs = result[1]

    for i in range(0, len(reqs)):
        temp = pairs[i].split('.')
        sheet1.write(i+2, 0, temp[1])
        sheet1.write(i+2, 1, temp[0])

    PCC = _get_PCC(request_count=request_count)
    LIC = _get_LIC(requests=reqs, request_count=request_count)
    RIN = _get_RIN(requests=reqs, request_count=request_count)
    RCD = _get_RCD(requests=reqs, request_count=request_count)

    RCD_IN = RCD[0]
    RCD_SE = RCD[1]
    # PIH_FNs = metrics[2]
    # PIH_FCs = metrics[3]
    # SIH_FNs = metrics[4]
    # SIH_FCs = metrics[5]

    for i in range(0, len(reqs)):
        sheet1.write(i+2, 2, PCC[i])
        sheet1.write(i+2, 6, LIC[i])
        sheet1.write(i+2, 5, RIN[i])
        sheet1.write(i+2, 3, ''.join(RCD_IN[i]))
        sheet1.write(i+2, 4, ''.join(RCD_SE[i]))

    book.save("result.xls")
