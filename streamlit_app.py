import streamlit as st

import re
import io

st.title("🎈 本页处理异网话单和51原始话单...")

from openpyxl import load_workbook
import math


def process_excel(file):
    # 提起文件的年月信息

    # 加载 Excel 文件
    wb = load_workbook(file)

    # 选择第一个工作表
    ws = wb.active

    # 在 L 列和 M 列添加标题
    ws["L1"] = "计算通话"
    ws["M1"] = "通话分钟"

    # 遍历每一行数据，进行计算并填充到 L 列和 M 列
    l_sum = 0
    m_sum = 0
    row_count = 2
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # for row in ws.iter_rows(min_row=2, values_only=True):
        k_value = row[10]  # 第 10 列对应 K 列，索引从 0 开始  --呼叫时长（秒）

        # 计算 L 列的值
        l_value = 1 if k_value != 0 else 0
        ws.cell(row=idx, column=12, value=l_value)
        # row[12].value = l_value  # 列号 12 对应 L 列，索引从 1 开始
        l_sum = l_sum + l_value

        # 计算 M 列的值
        m_value = math.ceil(k_value / 60)
        ws.cell(row=idx, column=13, value=m_value)  # 列号 13 对应 M 列，索引从 1 开始
        m_sum = m_sum + m_value

        row_count += 1

    # 保存修改后的 Excel 文件
    ws.cell(row=row_count, column=12, value=l_sum)
    ws.cell(row=row_count, column=13, value=m_sum)

    # 获取文件名
    filename = file.name
    # 使用正则表达式提取年月信息
    match = re.search(r'(\d{4})(\d{2})', filename)
    if match:
        year = str(match.group(1))
        month = str(match.group(2))
        st.write(f"Year: {year}, Month: {month}")
        output_filename = f"前程无忧异网话单{year}{month}.xlsx"
        # 将Excel文件保存到内存中
        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_data = excel_file.getvalue()
        return output_filename, excel_data, l_sum, m_sum
    else:
        st.write("无法从文件名中提取年月信息")
        return None, None, None, None


def process_correct(file, l_sum, m_sum):
    # 打开文件
    wb = load_workbook(file)
    ws = wb.active
    ws["C22"] = l_sum
    ws["d22"] = m_sum

    filename = file.name
    match = re.search(r'(\d{1,2})月', filename)
    if match:
        month = str(int(match.group(1))).zfill(2)
        new_filename = f"【前程】小号{month}月对账单-东讯昆辰（原始）.xlsx"
    else:
        new_filename = "修正后的原始话单.xlsx"

    # 将Excel文件保存到内存中
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_data = excel_file.getvalue()
    return new_filename, excel_data


def process_bill(file):
    def process_last_month(file):
        print("Processing last month:", file)
        # 在这里添加处理上月账单的逻辑
        pass

    def process_summary(file):
        print("Processing summary:", file)
        # 在这里添加处理汇总账单的逻辑
        pass


# 主界面
def main():
    st.title("Excel 处理工具")

    uploaded_file = st.file_uploader("上传上月异网话单", type=["xlsx"])

    uploaded_51org_file = st.file_uploader("上传上月51原始话单", type=["xlsx"])

    uploaded_month_before_last_file = st.file_uploader("上传前一个月51_云号话单", type=["xlsx"])

    if uploaded_month_before_last_file is not None:
        # 处理上传的 Excel 文件
        filename, data, l_sum, m_sum = process_excel(uploaded_file)

        st.success("文件处理完成！")
        # 下载生成的文件

    if uploaded_file is not None:
        uploaded_filename = uploaded_file.name
        # 处理上传的 Excel 文件
        filename, data, l_sum, m_sum = process_excel(uploaded_file)

        st.success("异网话单文件处理完成！")
        if filename and data:
            st.download_button(
                label="下载异网话单文件",
                data=data,
                file_name=filename,
                mime="application/vnd.ms-excel"
            )
        else:
            st.write("文件名生成失败")

    if uploaded_51org_file is not None:
        # 修正原始话单：
        excel_result = process_excel(uploaded_file)
        if excel_result:
            filename, data, l_sum, m_sum = excel_result
            correct_result = process_correct(uploaded_51org_file, l_sum, m_sum)
            if correct_result:
                new_filename, new_data = correct_result

                st.success("修正话单文件处理完成！")

                if new_filename and new_data:
                    st.download_button(
                        label="下载修正后的原始话单",
                        data=new_data,
                        file_name=new_filename,
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.write("文件名生成失败")
            else:
                st.write("修正话单文件处理失败")
        else:
            st.write("异网话单文件处理失败")

    if st.button("处理账单"):
        process_bill(uploaded_51org_file)


if __name__ == "__main__":
    main()
