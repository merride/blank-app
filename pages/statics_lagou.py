from dotenv import load_dotenv
import os
import oss2
import re
import openpyxl
import json
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd  # 添加 pandas 导入

# 加载 .env 文件
load_dotenv()

# 读取环境变量
access_key_id = os.getenv('ACCESS_KEY_ID')
access_key_secret = os.getenv('ACCESS_KEY_SECRET')

# 初始化 OSS 客户端
auth = oss2.Auth(access_key_id, access_key_secret)
bucket = oss2.Bucket(auth, 'https://oss-cn-shanghai.aliyuncs.com', 'checkbill')  # 修改为正确的 endpoint

# 获取 hongxin 目录下的文件列表
file_list = []
for obj in oss2.ObjectIterator(bucket, prefix='lagou/'):
    if not obj.key.endswith('/'):  # 排除目录
        file_list.append(obj.key)

# 处理每个文件
result = []
for file_key in file_list:
    # 下载文件到本地
    local_file_path = os.path.join("temp", os.path.basename(file_key))
    os.makedirs("temp", exist_ok=True)  # 确保 temp 目录存在
    bucket.get_object_to_file(file_key, local_file_path)

    # 使用正则表达式提取年和月
    match = re.search(r'(\d{4})年(\d{1,2})月', file_key)
    if match:
        year = match.group(1)
        month = match.group(2)

        # 使用 openpyxl 打开文件
        wb = openpyxl.load_workbook(local_file_path)
        # 选择“月账单”工作表
        ws = wb["月账单"]

        # 使用 pandas 读取 Excel 文件，指定引擎为 openpyxl
        try:
            # 方法1：使用 pandas 读取，但指定 engine='openpyxl'
            df = pd.read_excel(local_file_path, sheet_name="月账单", engine='openpyxl')
            f17 = df.iloc[15, -1]
            # st.write(f"DF原始值: {f17}")
            # 如果读取到的是 NaN，则使用 openpyxl 直接读取
            if pd.isna(f17):
                # 将sheet"1472"G34单元格的值赋值给 单元格D15
                try:
                    # 检查工作簿中是否存在"1472"工作表
                    if "1472" in wb.sheetnames:
                        # 获取"1472"工作表
                        sheet_1472 = wb["1472"]
                        # 获取G34单元格的值
                        g34_value = sheet_1472["G34"].value
                        # 将值赋给当前工作表的D15单元格
                        ws["D15"] = g34_value
                        st.write(f"已将工作表'1472'的G34单元格值 {g34_value} 赋值给D15单元格")
                    else:
                        st.warning("工作簿中不存在'1472'工作表")
                except Exception as e:
                    st.error(f"赋值过程中出错: {e}")
                 
                # 方法2：直接使用 openpyxl 读取单元格的值
                f17 = ws['F17'].value
                st.write(f"原始值: {f17}")
                # 如果仍然是公式，则尝试获取计算后的值
                if isinstance(f17, str) and f17.startswith('='):
                    # 获取公式计算后的值
                    f17_value = ws['F17'].value
                    st.write(f"原始公式: {f17}")
                    st.write(f"尝试直接获取值: {f17_value}")
                    
                    # 如果无法获取计算后的值，则尝试手动计算
                    # 这里假设 F17 是 F10:F16 的和
                    if "SUM(F10:F16)" in f17:
                        sum_value = 0
                        for row in range(10, 17):
                            cell_value = ws[f'F{row}'].value
                            if isinstance(cell_value, (int, float)):
                                sum_value += cell_value
                        f17 = sum_value
                        st.write(f"手动计算的和: {f17}")
        except Exception as e:
            st.error(f"读取 Excel 文件时出错: {e}")
            continue

        st.write(f"最终 F17 值: {f17}")

        # 处理带有 "=" 符号的数据
        def clean_value(value):
            if isinstance(value, str):
                if value.startswith('='):
                    # 如果是公式，返回0或其他默认值
                    return 0
                try:
                    return float(value)
                except ValueError:
                    return 0  # 如果无法转换为浮点数，返回0
            return float(value) if value is not None else 0  # 确保返回数值

        f17 = clean_value(f17)

        # 保存结果
        result.append({
            "年": year,
            "月": month,
            "F17": f17,
        })

    # 删除临时文件
    os.remove(local_file_path)

# 将结果保存为 JSON 文件
result_json_path = os.path.join("result.json")
with open(result_json_path, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=4)

# 绘制折线图
result_sorted = sorted(result, key=lambda x: (int(x['年']), int(x['月']))) 
years_months = [f"{item['年']}-{item['月']}" for item in result_sorted]
f17_values = [item['F17'] for item in result_sorted]

plt.figure(figsize=(10, 6))
plt.plot(years_months, f17_values, label='小号收入')

plt.xlabel('年月')
plt.ylabel('通话量')
plt.title('通话量随时间变化')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
chart_path = os.path.join("call_volume_chart.png")
plt.savefig(chart_path)

# 使用 Streamlit 显示结果
st.title("拉勾隐私号呼叫统计")

# 显示折线图
st.subheader("通话量随时间变化")
st.image(chart_path)

# 显示详细数字
st.subheader("详细数字")
with st.expander("点击展开查看详细数字", expanded=False):
    st.json(result)

st.write("处理完成，结果已保存到 result.json，图表已保存到 call_volume_chart.png")

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题