# ============================================================
# 100OrgData 資料整理程式（Streamlit 版）
# 功能：上傳 Excel → 判斷完整/不完整 → 分 Sheet → 下載 Excel
# ============================================================

import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st


# --------------------------------------------------
# 頁面設定
# --------------------------------------------------
st.set_page_config(
    page_title="100OrgData 資料整理工具",
    page_icon="📊",
    layout="wide",
)

st.title("📊 100OrgData 資料整理工具")
st.write("上傳 Excel 檔案，系統會自動檢查資料完整性，並將「完整資料」與「不完整資料」分成兩個工作表。")


# --------------------------------------------------
# 1. 上傳檔案
# --------------------------------------------------
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案",
    type=["xlsx", "xls"],
)


# --------------------------------------------------
# 2. 資料檢查規則
# --------------------------------------------------
def is_null_like(value):
    """判斷是否為空值、NULL、N/A、空白等"""
    if pd.isna(value):
        return True

    val = str(value).strip().lower()
    return val in [
        "", "null", "none", "n/a", "#na",
        "nil", "-", "nan", "none_provided",
    ]


def is_invalid_date(value):
    """判斷日期是否無效或格式不一致"""
    if is_null_like(value):
        return True

    val = str(value).strip()

    if re.search(r"00-00|02-31|1990-00-00|invalid", val, re.IGNORECASE):
        return True

    formats = [
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%Y.%m.%d",
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    ]

    for fmt in formats:
        try:
            datetime.strptime(val, fmt)
            return False
        except ValueError:
            continue

    try:
        pd.to_datetime(val, errors="raise")
        return False
    except Exception:
        return True


def is_invalid_phone(value):
    """判斷電話格式是否異常"""
    if is_null_like(value):
        return True

    val = str(value).strip()
    cleaned = re.sub(r"[\s\-\(\)\+]", "", val)

    return not re.match(r"^\d{8,15}$", cleaned)


def is_abnormal_age(value):
    """年齡異常"""
    if is_null_like(value):
        return True

    try:
        age = float(str(value).strip())
        return not (0 <= age <= 120)
    except ValueError:
        return True


def is_abnormal_credit_score(value):
    """Credit Score 異常"""
    if is_null_like(value):
        return True

    try:
        score = float(str(value).strip())
        return not (300 <= score <= 850)
    except ValueError:
        return True


def is_abnormal_salary(value):
    """薪資異常"""
    if is_null_like(value):
        return True

    try:
        salary = float(
            str(value)
            .replace(",", "")
            .replace("$", "")
            .replace("HK$", "")
            .strip()
        )
        return not (0 < salary < 500000)
    except ValueError:
        return True


def has_mixed_case_or_messy(value):
    """檢查是否含多餘空白"""
    if is_null_like(value):
        return False

    val = str(value)

    if val != val.strip():
        return True

    return "  " in val


def is_invalid_email(value):
    """Email 格式異常"""
    if is_null_like(value):
        return True

    val = str(value).strip().lower()

    if not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", val):
        return True

    return (
        "domain.org" in val
        or "none_provided" in val
        or "@@" in val
    )


# --------------------------------------------------
# 3. 執行完整性檢查
# --------------------------------------------------
def check_data_quality(df):
    incomplete_reasons = []

    key_fields = [
        "First_Name", "Last_Name", "Email_Address", "Phone_Number",
        "Age", "DOB", "Salary_Monthly", "Credit_Score",
    ]

    date_fields = [
        "DOB", "Hire_Date", "Registration_Date",
        "Last_Login_Timestamp", "Last_Purchase_Date",
    ]

    name_fields = ["First_Name", "Last_Name", "Full_Name"]

    for _, row in df.iterrows():
        reasons = []

        for field in key_fields:
            if field in df.columns and is_null_like(row[field]):
                reasons.append(f"{field}為空")

        for field in date_fields:
            if field in df.columns and is_invalid_date(row[field]):
                reasons.append(f"{field}日期無效")

        if "Phone_Number" in df.columns and is_invalid_phone(row["Phone_Number"]):
            reasons.append("電話格式異常")

        if "Age" in df.columns and is_abnormal_age(row["Age"]):
            reasons.append("年齡異常")

        if "Credit_Score" in df.columns and is_abnormal_credit_score(row["Credit_Score"]):
            reasons.append("Credit Score異常")

        if "Salary_Monthly" in df.columns and is_abnormal_salary(row["Salary_Monthly"]):
            reasons.append("薪資異常")

        if "Email_Address" in df.columns and is_invalid_email(row["Email_Address"]):
            reasons.append("Email格式異常")

        for field in name_fields:
            if field in df.columns and has_mixed_case_or_messy(row[field]):
                reasons.append(f"{field}格式混亂")
                break

        incomplete_reasons.append("; ".join(reasons) if reasons else "")

    df = df.copy()
    df["不完整原因"] = incomplete_reasons
    df["資料完整性"] = df["不完整原因"].apply(
        lambda x: "不完整" if x else "完整"
    )

    return df


# --------------------------------------------------
# 4. 產生 Excel 下載檔
# --------------------------------------------------
def create_excel_download(df_complete, df_incomplete):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_complete.to_excel(writer, sheet_name="完整資料", index=False)
        df_incomplete.to_excel(writer, sheet_name="不完整資料", index=False)

    output.seek(0)
    return output.getvalue()


# --------------------------------------------------
# 5. 主流程
# --------------------------------------------------
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"無法讀取 Excel 檔案：{e}")
        st.stop()

    st.success(f"成功上傳檔案：{uploaded_file.name}")
    st.info(f"原始資料共 {len(df):,} 筆，{df.shape[1]} 個欄位。")

    with st.spinner("正在檢查資料完整性..."):
        df_checked = check_data_quality(df)

    df_complete = df_checked[df_checked["資料完整性"] == "完整"].copy()
    df_incomplete = df_checked[df_checked["資料完整性"] == "不完整"].copy()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("原始資料", f"{len(df_checked):,} 筆")

    with col2:
        st.metric("完整資料", f"{len(df_complete):,} 筆")

    with col3:
        st.metric("不完整資料", f"{len(df_incomplete):,} 筆")

    tab_complete, tab_incomplete, tab_all = st.tabs(
        ["✅ 完整資料", "⚠️ 不完整資料", "📋 全部資料"]
    )

    with tab_complete:
        st.dataframe(df_complete, use_container_width=True)

    with tab_incomplete:
        st.dataframe(df_incomplete, use_container_width=True)

    with tab_all:
        st.dataframe(df_checked, use_container_width=True)

    excel_data = create_excel_download(df_complete, df_incomplete)

    st.download_button(
        label="📥 下載整理後 Excel 檔案",
        data=excel_data,
        file_name="整理後資料_完整與不完整.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("請先上傳 `.xlsx` 或 `.xls` 檔案。")