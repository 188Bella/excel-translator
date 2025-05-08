import streamlit as st
import openpyxl
from googletrans import Translator
from io import BytesIO

st.title("Excel 批量中英文对照翻译工具（无需API Key）")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    translator = Translator()
    st.write("正在翻译，请稍候...")

    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                try:
                    detected = translator.detect(cell.value)
                    dest_lang = 'en' if detected.lang == 'zh-cn' else 'zh-cn'
                    result = translator.translate(cell.value, dest=dest_lang)
                    cell.value = f"{cell.value}\n{result.text}"
                except Exception as e:
                    cell.value = f"{cell.value}\n[翻译失败]"
    # 保存到内存
    output = BytesIO()
    wb.save(output)
    st.success("翻译完成！")
    st.download_button(
        label="下载翻译后的Excel文件",
        data=output.getvalue(),
        file_name="output_with_translation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )