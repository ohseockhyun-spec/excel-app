import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO

st.title("엑셀 자동 정리 (LOT 가로 / 항목 세로)")

uploaded_files = st.file_uploader(
    "엑셀 또는 ZIP 파일 업로드",
    type=["xlsx", "zip"],
    accept_multiple_files=True
)

def extract_excel_from_zip(file):
    files = []
    with zipfile.ZipFile(file) as z:
        for name in z.namelist():
            if name.endswith(".xlsx"):
                files.append(BytesIO(z.read(name)))
    return files

def process_excel(file, filename):
    df = pd.read_excel(file, header=None)

    try:
        result = {
            "LOT": filename,
            "점착력(F8)": df.iloc[7,5],
            "점착력(G8)": df.iloc[7,6],
            "투습도": df.iloc[17,12],
            "흡수도": df.iloc[18,12],
            "투습 표준편차": df.iloc[19,12],
            "흡수 표준편차": df.iloc[20,12],
        }
        return result
    except:
        return None

if uploaded_files:
    data = []

    for file in uploaded_files:
        if file.name.endswith(".zip"):
            extracted = extract_excel_from_zip(file)
            for f in extracted:
                res = process_excel(f, "ZIP_FILE")
                if res:
                    data.append(res)
        else:
            res = process_excel(file, file.name)
            if res:
                data.append(res)

    df = pd.DataFrame(data)

    # 🔥 소수 첫째자리 반올림
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(1)

    # 🔥 LOT 가로 / 항목 세로 변환
    final_df = df.set_index("LOT").T

    st.subheader("결과")
    st.dataframe(final_df)

    # 다운로드
    output = BytesIO()
    final_df.to_excel(output)
    st.download_button(
        "엑셀 다운로드",
        data=output.getvalue(),
        file_name="정리결과.xlsx"
    )