import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import os

st.title("엑셀 자동 정리 (LOT 가로 / 항목 세로)")

uploaded_files = st.file_uploader(
    "엑셀 또는 ZIP 파일 업로드 (.xlsx, .xlsm, .zip)",
    type=["xlsx", "xlsm", "zip"],
    accept_multiple_files=True
)

def clean_lot_name(filename):
    base = os.path.basename(filename)
    base = os.path.splitext(base)[0]
    base = base.split("_")[0]
    return base

def extract_excel_from_zip(file):
    files = []

    with zipfile.ZipFile(file) as z:
        for name in z.namelist():
            lower_name = name.lower()

            if lower_name.endswith((".xlsx", ".xlsm")) and not os.path.basename(name).startswith("~$"):
                file_bytes = BytesIO(z.read(name))
                file_bytes.name = os.path.basename(name)
                files.append((file_bytes, os.path.basename(name)))

    return files

def process_excel(file, filename):
    try:
        df = pd.read_excel(file, header=None, engine="openpyxl")

        result = {
            "LOT": clean_lot_name(filename),
            "점착력(F8)": df.iloc[7, 5],
            "점착력(G8)": df.iloc[7, 6],
            "투습도": df.iloc[17, 12],
            "흡수도": df.iloc[18, 12],
            "투습 표준편차": df.iloc[19, 12],
            "흡수 표준편차": df.iloc[20, 12],
        }

        return result, None

    except Exception as e:
        return None, str(e)

if uploaded_files:
    data = []
    errors = []

    for file in uploaded_files:
        lower_name = file.name.lower()

        if lower_name.endswith(".zip"):
            extracted_files = extract_excel_from_zip(file)

            for excel_file, excel_filename in extracted_files:
                res, err = process_excel(excel_file, excel_filename)

                if res:
                    data.append(res)
                else:
                    errors.append({
                        "파일명": excel_filename,
                        "오류내용": err
                    })

        else:
            res, err = process_excel(file, file.name)

            if res:
                data.append(res)
            else:
                errors.append({
                    "파일명": file.name,
                    "오류내용": err
                })

    if data:
        df = pd.DataFrame(data)

        for col in df.columns[1:]:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(1)

        final_df = df.set_index("LOT").T
        final_df.index.name = "항목"

        st.subheader("결과")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="결과")
        output.seek(0)

        st.download_button(
            "엑셀 다운로드",
            data=output.getvalue(),
            file_name="정리결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if errors:
        st.subheader("오류 파일")
        st.dataframe(pd.DataFrame(errors), use_container_width=True)

else:
    st.info("엑셀 또는 ZIP 파일을 업로드하세요.")
