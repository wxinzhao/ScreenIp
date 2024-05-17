# @Time : 2024-05-16 14:00
# @Author : WXZ
# @File : ScreenIpPlus.py
# @Software: PyCharm
#固定IP和引用另一个表格对比
import pandas as pd
from tkinter import Tk, filedialog
import os
from datetime import datetime

def select_file():
    root = Tk()
    root.withdraw()  # 隐藏Tk窗口
    file_path = filedialog.askopenfilename()  # 打开文件选择器
    return file_path

def find_duplicate_ips():
    try:
        # 定义固定的IP地址列表
        fixed_ips = ['192.168.0.1', '192.168.0.2', '192.168.0.3']  # 修改为你的固定IP列表

        print("请选择一个Excel文件：")
        file1 = select_file()
        if not file1:
            print("未选择文件，退出程序。")
            return

        print(f"成功选择文件: {file1}")

        # 获取当前日期和时间
        now = datetime.now()
        # 构造输出文件名，包含年月日时分秒
        output_file = now.strftime("output_%Y-%m-%d_%H-%M-%S.xlsx")

        # 获取上次输出的文件列表
        previous_output_files = [file for file in os.listdir() if file.startswith("output")]

        # 删除上次输出的文件
        for file in previous_output_files:
            os.remove(file)
            print(f"已删除上一次的输出文件 {file}。")

        # 读取选定的Excel文件
        df1 = pd.read_excel(file1, header=None)  # 没有表头，直接读取第一行作为数据

        # 假设IP地址在第一列
        ips1 = df1[0]

        # 将固定IP列表转换为Series
        fixed_ips_series = pd.Series(fixed_ips)

        # 统计每个IP的出现次数
        ip_counts = pd.concat([fixed_ips_series, ips1]).value_counts()
        print("IP统计完成。")

        # 找到重复的IP
        duplicate_ips = ip_counts[ip_counts > 1]
        print(f"找到 {len(duplicate_ips)} 个重复的IP。")

        if not duplicate_ips.empty:
            # 创建一个新的DataFrame来保存结果
            df_duplicates = pd.DataFrame({'IP': duplicate_ips.index, 'Count': duplicate_ips.values})
            df_duplicates.to_excel(output_file, sheet_name='Duplicate IPs', index=False)
            print(f"结果已写入 {output_file}")
        else:
            print("所有IP地址均不重复，不生成输出文件。")

    except Exception as e:
        print(f"运行过程中出现错误: {e}")

# 调用函数
find_duplicate_ips()
