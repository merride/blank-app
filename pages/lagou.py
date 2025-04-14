import streamlit as st
from openpyxl import load_workbook
import math

st.title("🎈 本页处理拉勾隐私号话单...")

# Upload button for last month's billing data
uploaded_file = st.file_uploader("上传上月原始话单", type=["xlsx"])

# Upload button for the previous month's billing data with seal
uploaded_file2 = st.file_uploader("上传上上月盖章话单", type=["xlsx"])


if uploaded_file is not None:
    # 获取文件名
    uploaded_filename = uploaded_file.name
    st.write(uploaded_filename)

    import re
    # 使用正则表达式提取年份和月份
    match = re.search(r'(\d{4})年(\d{1,2})月', uploaded_filename)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        st.write(f"Year: {year}, Month: {month}")  # 输出年份和月份
    else:
        st.write("无法从文件名中提取年份和月份")
    
    # Load the workbook
    workbook = load_workbook(filename=uploaded_file)

    def process_data():
        if uploaded_file2 is None:
            st.write("请上传上上月盖章话单")
            return

        # Load the workbooks
        wb1 = load_workbook(filename=uploaded_file)  # 上月原始话单
        wb2 = load_workbook(filename=uploaded_file2)  # 上上月盖章话单

        # 单元格赋值
        sheet1 = wb1["月账单"]
        sheet2 = wb2["月账单"]

        #标题和月份
        sheet2["A1"] = sheet1["A1"].value
        sheet2["B2"] = sheet1["B2"].value

        # C10:D11 --> C10:D11
        for i in range(10, 12):
            for j, col in enumerate(['C', 'D']):
                sheet2[f'{col}{i}'] = sheet1[f'{col}{i}'].value

        # C15:D18 --> C13:D16
        for i in range(15, 19):
            for j, col in enumerate(['C', 'D']):
                sheet2[f'{col}{i - 2}'] = sheet1[f'{col}{i}'].value
                st.write(f'sheet2{col}{i - 2}')
        
        # 获取1472工作表中F34和G34的值，并直接赋值给C15和D15
        if "1472" in wb1.sheetnames:
            sheet_1472 = wb1["1472"]
            # 获取F34和G34的值
            f34_value = sheet_1472["F34"].value
            g34_value = sheet_1472["G34"].value
            
            # 直接赋值为数字，而不是公式
            sheet2["C15"] = f34_value
            sheet2["D15"] = g34_value
            
            st.write(f"已将1472工作表的F34值({f34_value})赋给C15，G34值({g34_value})赋给D15")
        else:
            st.warning("未找到1472工作表，无法获取F34和G34的值")
        
        # Sheet 拷贝
        sheets_to_copy = ["1421", "1472"]

        # 删除工作表
        for sheet_name in sheets_to_copy:
            if sheet_name in wb2.sheetnames:
                del wb2[sheet_name]

        for sheet_name in sheets_to_copy:
            if sheet_name in wb1.sheetnames:
                source_sheet = wb1[sheet_name]
                # 创建一个新的工作表
                target_sheet = wb2.create_sheet(title=sheet_name)

                # 复制单元格内容
                for row in source_sheet.iter_rows():
                    for cell in row:
                        target_sheet[cell.coordinate].value = cell.value

                # 删除 B 列
                target_sheet.delete_cols(2)  # 删除第二列，即 B 列
            else:
                st.write(f"Sheet {sheet_name} not found in the source workbook.")

        
        
        # 保存文件
        output_filename = f"【拉勾】小号{year}年{month}月对账单_上海大颂.xlsx"
        wb2.save(output_filename)

    if st.button("处理"):
        process_data()
        st.success("数据已成功处理！")

        # Save the processed data to a new workbook
        output_workbook = load_workbook(filename=uploaded_file)
        # Add your code here to write the processed_data into the output_workbook

        # Download button for the processed billing data
        output_filename = f"【拉勾】小号{year}年{month}月对账单_上海大颂.xlsx"
        with open(output_filename, "rb") as file:
            btn = st.download_button(
                label="下载",
                data=file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("请下载上月正式话单！")