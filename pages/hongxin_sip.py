import streamlit as st
import openpyxl
import pandas as pd
import re
from dotenv import load_dotenv
import oss2
import os  # 添加 os 模块导入

st.title("🎈 宏信SIP语音对账单处理...")

# Upload button for last month's account data
account_last_month = st.file_uploader("上传上上月账单", type=["xlsx"])

# Upload button for the original billing data
orgbill = st.file_uploader("上传上月原始话单", type=["xls"])


if account_last_month and orgbill:
    def process_data():
        # 1. 读取原始话单数据
        try:
            # 使用pandas读取xls文件
            df_orgbill = pd.read_excel(orgbill)
        except Exception as e:
            st.error(f"读取原始话单失败: {e}")
            return

        # 2. 统计原始话单中Trunk列的数量
        tk05 = 0
        tk0910 = 0
        tk17 = 0
        for trunk_value in df_orgbill["队列"]:  # H列
            if trunk_value == "TRUNK005":
                tk05 += 1
            elif trunk_value == "TRUNK009" or trunk_value == "TRUNK010":
                tk0910 += 1
            elif trunk_value == "TRUNK017":
                tk17 += 1

        # 3. 读取上上月账单数据
        try:
            # 使用pandas读取上上月账单数据
            df_account_last_month = pd.read_excel(account_last_month)
            st.write("上上月账单数据:")
            st.write(df_account_last_month)
            try:
                acount_tk05 = str(df_account_last_month.iloc[8, 6])
                st.write(f"acount_tk05: {acount_tk05}")
                acount_tk0910 = str(df_account_last_month.iloc[9, 6])
                st.write(f"acount_tk0910: {acount_tk0910}")
                acount_tk17 = str(df_account_last_month.iloc[10, 6])
                st.write(f"acount_tk17: {acount_tk17}")
            except Exception as e:
                st.error(f"读取上上月账单数据失败: {e}")
                return
        except Exception as e:
            st.error(f"pandas读取上上月账单失败: {e}")
            return

        # 4. 读取上上月账单数据
        try:
            # 使用openpyxl读取xlsx文件
            workbook_account = openpyxl.load_workbook(account_last_month)
            sheet_account = workbook_account.active  # 默认读取第一个sheet
        except Exception as e:
            st.error(f"openpyxl读取上上月账单失败: {e}")
            return

        # 5. 修改上月账单的年份和月份
        try:
            b2_value = sheet_account["B2"].value
            st.write(b2_value)
            match = re.search(r"(\d{4})年(\d{1,2})月对账表", b2_value)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                next_month = (month % 12) + 1
                next_year = year + 1 if next_month == 1 else year  # 如果N=1，年份加一
                new_b2_value = f"{next_year}年{next_month}月对账表"
                sheet_account["B2"] = new_b2_value
            else:
                st.warning("无法从B2单元格中提取年份和月份")
        except Exception as e:
            st.error(f"修改上月账单的年份和月份失败: {e}")
            return

        # 6. 修改A1单元格
        try:
            next_month = (month % 12) + 1  # 重新计算next_month，确保与步骤4一致
            sheet_account["A1"] = f"{next_month}月SIP账单（宏信）"
        except Exception as e:
            st.error(f"修改A1单元格失败: {e}")
            return

        # 7. 修改A9、A10、A11单元格
        try:
            sheet_account["A10"] = f"{next_month}月份"
            sheet_account["A11"] = f"{next_month}月份"
            sheet_account["A12"] = f"{next_month}月份"
        except Exception as e:
            st.error(f"修改A10、A11、A12单元格失败: {e}")
            return

        # 8. 修改G10、G11、G12单元格的公式
        try:
            sheet_account["G10"] = f"={acount_tk05}-F10"
            sheet_account["G11"] = f"={acount_tk0910}-F11"
            sheet_account["G12"] = f"={acount_tk17}-F12"
        except Exception as e:
            st.error(f"修改G10、G11、G12单元格的公式失败: {e}")
            return

        # 9. 修改D10、D11、D12单元格的值
        try:
            sheet_account["D10"] = tk05
            sheet_account["D11"] = tk0910
            sheet_account["D12"] = tk17
        except Exception as e:
            st.error(f"修改D10、D11、D12单元格的值失败: {e}")
            return

        # 10. 保存文件
        try:
            next_month = (month % 12) + 1  # 重新计算next_month，确保与步骤4和5一致
            next_year = year + 1 if next_month == 1 else year
            output_filename = f"{next_year}年{next_month}月SIP语音对账单_上海大颂.xlsx"
            workbook_account.save(output_filename)
            st.success("数据已成功处理！")

            # 保存文件名到 session_state
            st.session_state.output_filename = output_filename
            
            # 只显示下载按钮
            with open(output_filename, "rb") as file:
                st.download_button(
                    label="下载",
                    data=file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success("请下载本月正式话单！")

        except Exception as e:
            st.error(f"保存文件失败: {e}")

    if st.button("处理"):
        process_data()

# 如果已经处理过数据并生成了文件，显示保存到阿里OSS的按钮
if 'output_filename' in st.session_state:
    output_filename = st.session_state.output_filename
    # 显示保存到阿里OSS的按钮
    if st.button("保存到阿里oss", key="save_to_oss_button"):
        try:
            # 读取环境变量
            access_key_id = os.getenv('ACCESS_KEY_ID')
            access_key_secret = os.getenv('ACCESS_KEY_SECRET')

            # 初始化 OSS 客户端
            auth = oss2.Auth(access_key_id, access_key_secret)
            bucket = oss2.Bucket(auth, 'https://oss-cn-shanghai.aliyuncs.com', 'checkbill')

            # 上传文件到 OSS
            oss_key = f"hongxin/{output_filename}"
            with open(output_filename, "rb") as file:
                bucket.put_object(oss_key, file)
            st.success("保存成功")
        except Exception as e:
            st.error(f"保存到阿里 OSS 失败: {e}")

# 添加安装依赖项的说明
st.markdown("请确保已安装以下库：")
st.markdown("- `streamlit`")
st.markdown("- `openpyxl`")
st.markdown("- `pandas`")