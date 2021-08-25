# -*- coding:utf-8
import pandas as pd
from pandas import DataFrame
import locale

ORIGINAL_FILE = "./电子对账单8月-原始.csv"  # 原始文件
OUTPUT_FILE = "电子对账单7月~8月.xlsx"  # 生成文件


def main():
    """
    :return:
    """
    print("*******begin******")
    # 1. 打开原始文件，读取其中表格部分
    table = pd.read_csv(ORIGINAL_FILE, encoding="gbk", header=4)

    # 2. 处理表格数据
    lt = list()
    locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
    for i in range(table.shape[0]):
        lt.append([table.loc[i]["日期"], table.loc[i]["交易类型"],
                   0, table.loc[i]["摘要"],
                   -round(locale.atof(table.loc[i]["借方发生额"]), 2) if table.loc[i]["贷方发生额"] == '0.00' else
                   round(locale.atof(table.loc[i]["贷方发生额"]), 2),
                   table.loc[i]["余额"], table.loc[i]["对方户名"],
                   table.loc[i]["对方账号"], ""])

    # 3. 存储文件
    df = pd.DataFrame()
    df.to_excel(OUTPUT_FILE)

    data = pd.read_excel(OUTPUT_FILE)
    data["日期"] = None
    data["交易类型"] = None
    data["票据号"] = None
    data["摘要"] = None
    data["借/贷方发生额"] = None
    data["余额"] = None
    data["对方户名"] = None
    data["对方账户"] = None
    data["七零一备注"] = None

    for i in range(table.shape[0]):
        data.loc[i] = lt[i]
    # df = DataFrame(data)
    # df.to_excel("./test.xlsx", sheet_name="Sheet", startrow=4, index=False, header=True)

    # 4.插入头说明
    with open(ORIGINAL_FILE, encoding="gbk") as f:
        text_pre = f.readlines()[:4]
    print("text_pre[0]", text_pre[0].split(",")[:-1])

    df = DataFrame(data)
    writer = pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter")
    df.to_excel(writer, startrow=4, index=False)
    worksheet = writer.sheets["Sheet1"]

    worksheet.write(0, 0, text_pre[0].split(",")[:-1][0])
    worksheet.write(0, 1, text_pre[0].split(",")[:-1][1].split('"')[1])
    worksheet.write(1, 0, text_pre[1].split(",")[:-1][0])
    worksheet.write(1, 1, text_pre[1].split(",")[:-1][1].split('"')[1])
    worksheet.write(2, 0, text_pre[2].split(",")[:-1][0])
    worksheet.write(2, 1, text_pre[2].split(",")[:-1][1].split('"')[1])
    worksheet.write(3, 0, text_pre[3].split(",")[:-1][0])
    worksheet.write(3, 1, text_pre[3].split(",")[:-1][1].split('"')[1])
    # writer.save()

    # 5.设置格式
    worksheet.set_column("A:A", 10)
    worksheet.set_column("B:B", 10)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 22)
    worksheet.set_column("E:E", 15)
    worksheet.set_column("F:F", 10)
    worksheet.set_column("G:G", 19)
    worksheet.set_column("H:H", 19)
    worksheet.set_column("I:I", 11)
    writer.save()
    # writer.close()

    print("操作完成")
    print("*******end******")


if __name__ == "__main__":
    """ python3 changecsv2xlsx.py
    
    参考资料:
        1. pandas 保存数据到excel,csv https://blog.csdn.net/weixin_35757704/article/details/88180581
        2. python中如何使用pandas创建excel文件 https://jingyan.baidu.com/article/ca41422f79039c1eaf99ed73.html
        3. excel前插入若干行 https://stackoverflow.com/questions/66459258/inserting-a-row-above-pandas-column-headers-to-save-a-title-name-in-the-first-ce
        4. 如何在 Python 中将字符串转换为浮点或整数 https://www.delftstack.com/zh/howto/python/how-to-convert-string-to-float-or-int/
        5. https://blog.csdn.net/weixin_44606231/article/details/105552164
        6. 对dataframe用to_excel输出表格的格式修改 https://blog.csdn.net/qq_38115310/article/details/98031934
    """

    main()
