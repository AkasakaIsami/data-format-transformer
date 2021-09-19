# Service-pair-level metric transformer

HOW TO USE:

1.2.py是cqq写的，我根据他的脚本重写了一份



大致使用方法和服务级别指标是一样的，先在配置文件里写服务数量，再将原始数据写进文件、用脚本来执行格式转换，保存在excel文件中。由于输出格式的特殊性（每个服务对表达形式都不一样，做匹配工作量太大……）所以这里四个指标分三个文件处理和存储：

| 指标      | 文件                                                         |
| --------- | ------------------------------------------------------------ |
| DAP       | transfer1.txt & service-pair-level-transformer-1.py &  result1.txt |
| PIP & SIP | transfer2.txt &  service-pair-level-transformer-2.py & result2.txt |
| TCC       | transfer3.txt &  service-pair-level-transformer-3.py & result3.txt |

1. DAP

   FN和FC都只有两行数据，把这两行数据输入transfer1.txt即可

   

2. PIP & SIP

   输入4个指标的所有数据，设服务数为n，transfer2.txt应有 nC2 * 4 条数据（比如train-ticket有32个服务，这里应该有1984行数据）

   

3. TCC

   输入2个指标的所有数据，设服务数为n，transfer3.txt应有 nC2 * 2 条数据

