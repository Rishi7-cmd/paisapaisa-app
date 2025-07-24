
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

st.set_page_config(page_title="Paisa Paisa Flowchart", page_icon="📊", layout="wide")
st.markdown("<h1 style='text-align: center; color: #FFD700;'>📊 Paisa Paisa Transaction Flowchart</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload the input Excel file", type=["xlsx"])

def match_column(possible_matches, df_columns):
    for match in possible_matches:
        for col in df_columns:
            if match.lower() in col.lower():
                return col
    return None

def format_account(row, amount_col):
    lines = []
    if 'Bank' in row:
        lines.append(f"Bank: {row['Bank']}")
    if 'A/c No' in row:
        lines.append(f"A/c No: {row['A/c No']}")
    if 'IFSC' in row:
        lines.append(f"IFSC: {row['IFSC']}")
    amount = row.get(amount_col, None)
    if amount and pd.notna(amount):
        lines.append(f"Amount: ₹{amount:,.0f}")
    return "\n".join(lines)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    sender_col = match_column(["Sender account", "Account No./ (Wallet /PG/PA) Id"], df.columns)
    receiver_col = match_column(["Receiver account", "Account No"], df.columns)
    txn_amt_col = match_column(["Transaction Amount", "Amount"], df.columns)
    bank_col = match_column(["Bank/FIs", "Bank"], df.columns)
    ifsc_col = match_column(["IFSC Code", "Ifsc Code"], df.columns)

    if not sender_col or not receiver_col or not txn_amt_col:
        st.error("❌ Required columns not found.")
        st.stop()

    df[txn_amt_col] = pd.to_numeric(
        df[txn_amt_col].astype(str).str.replace(",", "").str.replace("₹", "").str.strip(),
        errors="coerce"
    )
    df = df[df[txn_amt_col] > 50000]

    if bank_col: df["Bank"] = df[bank_col]
    if ifsc_col: df["IFSC"] = df[ifsc_col]
    df["A/c No"] = df[receiver_col]

    victim = df[sender_col].value_counts().idxmax()
    layer1_df = df[df[sender_col] == victim]
    layer1_receivers = layer1_df[receiver_col].unique()

    wb = Workbook()
    ws = wb.active
    ws.title = "Flowchart"
    ws["A1"] = f"Victim: {victim}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=50)

    col_pos = 1
    for l1 in layer1_receivers:
        row_pos = 2
        l1_row = df[df[receiver_col] == l1].iloc[0]
        l1_text = format_account(l1_row, txn_amt_col)
        ws.cell(row=row_pos, column=col_pos, value=l1_text)
        ws.cell(row=row_pos + 1, column=col_pos, value="↓")

        l2_df = df[df[sender_col] == l1]
        for _, l2_row in l2_df.iterrows():
            acct_text = format_account(l2_row, txn_amt_col)
            ws.cell(row=row_pos + 2, column=col_pos, value=acct_text)
            ws.cell(row=row_pos + 3, column=col_pos, value="↓")
            row_pos += 2

        col_pos += 2

    for col in ws.columns:
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = uploaded_file.name.replace(".xlsx", "_flowchart.xlsx")

    st.success("✅ Flowchart created. Download below:")
    st.download_button("📥 Download Excel Flowchart", data=output, file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
