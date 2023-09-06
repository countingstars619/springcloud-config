import os
from tkinter import filedialog
import tkinter as tk
import pandas as pd


# NVH组         标准法规组      车身组       尺寸工程组   冲压工艺组 底盘组 电气组 动力组 工厂物流组 焊装工艺组
# 焊装组         零部件质量组    软件组       饰件组
# 同步工程组      涂装工艺组      网联组       增程组      整车测试组
# 整车控制策略组   自动驾驶组     总装工艺组    总装组       座舱组   24
def create_rectification_issues_table(file_path):
    colunms_to_read = ['责任组/部门', '问题状态']
    data = pd.read_excel(file_path, usecols=colunms_to_read, engine='openpyxl')

    # 执行透视表操作,创建新的数据表格,这个表格以指定的行索引和列索引,聚合函数生成新的数据表格
    # data 原始数据表格
    # values='责任组/部门' 指定要在透视表中使用的数值列，用于填充透视表的数据单元格
    # index='责任组/部门' 指定哪一列的唯一值作为透视表的行索引
    # columns='问题状态' 指定哪一列的唯一值作为透视表的列索引
    # aggfunc='size', 统计行和列中值相等的个数
    # 以0填充缺失值
    pivot_table = pd.pivot_table(data, values='责任组/部门', index='责任组/部门', columns='问题状态', aggfunc='size',
                                 fill_value=0)
    # 添加新的行：工艺=冲焊涂总+尺寸+同步
    row1 = pivot_table.iloc[[1, 9, 10, 14, 15, 16, 3, 5]].sum()
    pivot_table = pivot_table.append(pd.Series(row1, name="工艺"))
    # 添加新的行：制造=冲焊涂总+物流
    row2 = pivot_table.iloc[[1, 9, 10, 14, 15, 16, 3, 6]].sum()
    pivot_table = pivot_table.append(pd.Series(row2, name="制造"))

    # row3 = pivot_table.iloc[[1,2]].sum()
    # pivot_table = pivot_table.append(pd.Series(row3, name="row3"))

    # 创建新的列“问题数”，该列保存每一行值的总和。axis=1表示操作沿着列的方向进行，会对每一行的数据进行操作。
    pivot_table['问题数'] = pivot_table.sum(axis=1)

    # 重新排列列的顺序，将新列移动到第一列的位置
    column_order = ['问题数'] + pivot_table.columns.tolist()[:-1]
    pivot_table = pivot_table[column_order]

    # 添加新的行：合计
    row3 = pivot_table.sum(axis=0)
    pivot_table = pivot_table.append(pd.Series(row3, name="合计"))

    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    pivot_table.to_excel(f"{desktop_path}/问题整改数_.xlsx")



# 生成发现问题数量excel表
def create_find_issues_table(file_path: object) -> object:
    colunms_to_read = ['问题状态']
    data = pd.read_excel(file_path, usecols=colunms_to_read, engine='openpyxl')
    # 返回每个不同行的频率
    counts = data.value_counts()
    # counts.append({'问题状态':'合计', '数量':data['问题状态'].sum})
    # 将seres转dataframe
    counts_df = pd.DataFrame({'数量': counts})
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    counts_df.to_excel(f"{desktop_path}/发现问题数_.xlsx")


# file_path = "E:\\部门轮岗\\学习资料\\测试项目管理部学习资料\\data_copy\\2023-08-31_17-28-37.xlsx"
# create_find_issues_table(file_path)
root = tk.Tk()
root.withdraw()
# print("bbbbb")
file_path = filedialog.askopenfilename()
root.update()
# print("aaaa")
if file_path:
    print("你选择了文件: %s" % file_path)
else:
    print("你取消了文件选择。")
create_find_issues_table(file_path)
create_rectification_issues_table(file_path)
