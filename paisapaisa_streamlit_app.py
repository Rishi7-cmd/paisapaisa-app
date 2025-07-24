
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

st.set_page_config(page_title="Paisa Paisa Visual Flow", page_icon="📊", layout="wide")
st.markdown("<h1 style='text-align: center; color: #FFD700;'>📊 Paisa Paisa L1 → L2 → Withdrawals Flowchart</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload Excel with Transactions", type=["xlsx"])

def match_column(possible, cols):
    for opt in possible:
        for col in cols:
            if opt.lower() in col.lower():
                return col
    return None

def format_account(row, bank_col, acct_col, ifsc_col, amount, label="Sent"):
    lines = []
    if bank_col: lines.append(f"Bank: {row.get(bank_col, '')}")
    if acct_col: lines.append(f"A/c No: {row.get(acct_col, '')}")
    if ifsc_col: lines.append(f"IFSC: {row.get(ifsc_col, '')}")
    lines.append(f"Amount {label}: ₹{int(amount):,}")
    return "\n".join(lines)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    sender_col = match_column(["Sender", "Account No./ (Wallet /PG/PA) Id"], df.columns)
    receiver_col = match_column(["Receiver", "Account No"], df.columns)
    amount_col = match_column(["Transaction Amount", "Amount"], df.columns)
    bank_col = match_column(["Bank/FIs", "Bank"], df.columns)
    ifsc_col = match_column(["IFSC Code", "Ifsc Code"], df.columns)

    if not sender_col or not receiver_col or not amount_col:
        st.error("Missing required columns.")
        st.stop()

    df[amount_col] = pd.to_numeric(df[amount_col].astype(str).str.replace(",", "").str.replace("₹", ""), errors="coerce")
    df = df[df[amount_col] > 50000]

    victim = df[sender_col].value_counts().idxmax()

    wb = Workbook()
    ws = wb.active
    ws.title = "Flowchart"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=50)
    ws.cell(row=1, column=1).value = f"Victim: {victim}"

    col = 1
    layer1_df = df[df[sender_col] == victim]
    for l1_acct in layer1_df[receiver_col].unique():
        row = 3
        l1_row = layer1_df[layer1_df[receiver_col] == l1_acct].iloc[0]
        l1_text = format_account(l1_row, bank_col, receiver_col, ifsc_col, l1_row[amount_col], "Sent")
        ws.cell(row=row, column=col).value = l1_text
        ws.cell(row=row+1, column=col).value = "↓"

        layer2_df = df[df[sender_col] == l1_acct]
        for _, l2_row in layer2_df.iterrows():
            acct_text = format_account(l2_row, bank_col, receiver_col, ifsc_col, l2_row[amount_col], "Received")
            ws.cell(row=row+2, column=col).value = acct_text
            ws.cell(row=row+3, column=col).value = "↓"

            # withdrawal logic
            withdrawals = df[(df[sender_col] == l2_row[receiver_col]) & (df[receiver_col].isna())]
            for _, wd in withdrawals.iterrows():
                amt = wd[amount_col]
                acct = wd[sender_col]
                ws.cell(row=row+4, column=col).value = f"💸 Withdrawal Made\nFrom: Layer 2\nA/c No: {acct}\nAmount: ₹{int(amt):,}"
                row += 2

            row += 4

        col += 2

    for col_cells in ws.columns:
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    fname = uploaded_file.name.replace(".xlsx", "_flowchart.xlsx")
    st.success("✅ Flowchart generated.")
    st.download_button("📥 Download Flowchart Excel", data=output, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
