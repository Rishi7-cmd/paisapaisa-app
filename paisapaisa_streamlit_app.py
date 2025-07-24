
import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Paisa Paisa App", page_icon="💸", layout="wide")
st.markdown(
    "<h1 style='text-align: center; color: #FFD700;'>💡 Paisa Paisa Transaction Flow Analyzer</h1>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("📂 Upload the input Excel file", type=["xlsx"])

def match_column(possible_matches, df_columns):
    for match in possible_matches:
        for col in df_columns:
            if match.lower() in col.lower():
                return col
    return None

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Flexible column name detection
    sender_col = match_column(["Sender account", "Account No./ (Wallet /PG/PA) Id"], df.columns)
    receiver_col = match_column(["Receiver account", "Account No"], df.columns)
    txn_amt_col = match_column(["Transaction Amount", "Amount"], df.columns)

    if not sender_col or not receiver_col or not txn_amt_col:
        st.error("❌ Required columns not found in your file.")
        st.write("Required (flexible match):")
        st.code("Sender account → e.g., 'Account No./ (Wallet /PG/PA) Id'\nReceiver account → e.g., 'Account No'\nTransaction Amount → e.g., 'Transaction Amount'", language="text")
        st.stop()

    # Clean and filter
    df[txn_amt_col] = pd.to_numeric(
        df[txn_amt_col]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip(),
        errors="coerce"
    )
    df = df[df[txn_amt_col] > 50000]

    if df.empty:
        st.warning("⚠️ No transactions above ₹50,000 found after filtering.")
        st.stop()

    victim = df[sender_col].value_counts().idxmax()
    layer1_df = df[df[sender_col] == victim]
    layer1_receivers = layer1_df[receiver_col].unique()

    output_data = []
    for l1 in layer1_receivers:
        l1_trans = df[df[sender_col] == l1]
        layer2_receivers = l1_trans[receiver_col].unique()
        l2_data = []

        for l2 in layer2_receivers:
            wd = df[(df[sender_col] == l2) & (df[receiver_col].isna())]
            if not wd.empty:
                amount = wd[txn_amt_col].sum()
                highlight = "⚠️" if amount > 100000 else ""
                l2_data.append(f"{l2}\n💰 {amount}{highlight}")

        if l2_data:
            output_data.append([f"{victim}", f"{l1}", "\n\n".join(l2_data)])

    # Excel output
    wb = Workbook()
    ws = wb.active
    ws.title = "Layered Flow"

    ws.append(["Victim", "Layer 1", "Layer 2 + Withdrawals"])
    for row in output_data:
        ws.append(row)

    # Highlight large withdrawals
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for row in ws.iter_rows(min_row=2, max_col=3):
        if "⚠️" in str(row[2].value):
            for cell in row:
                cell.fill = yellow_fill

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    out_filename = uploaded_file.name.replace(".xlsx", "_output.xlsx")

    st.success("✅ Analysis complete! Download your output Excel file below:")
    st.download_button("📥 Download Output Excel", output, file_name=out_filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
