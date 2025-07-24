
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="Paisa Paisa Flowchart", page_icon="📊", layout="wide")
st.markdown("<h1 style='text-align: center; color: #FFD700;'>📊 Unique Flowchart: Victim → L1 → L2 → Withdrawal</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

def match_column(possibles, columns):
    for option in possibles:
        for col in columns:
            if option.lower() in col.lower():
                return col
    return None

def format_block(row, bank_col, acct_col, ifsc_col, amount, label):
    parts = []
    if bank_col and row.get(bank_col): parts.append(f"Bank: {row[bank_col]}")
    if acct_col and row.get(acct_col): parts.append(f"A/c No: {row[acct_col]}")
    if ifsc_col and row.get(ifsc_col): parts.append(f"IFSC: {row[ifsc_col]}")
    parts.append(f"Amount {label}: ₹{int(amount):,}")
    return "\n".join(parts)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    sender_col = match_column(["Sender", "Account No./ (Wallet /PG/PA) Id"], df.columns)
    receiver_col = match_column(["Receiver", "Account No"], df.columns)
    amount_col = match_column(["Transaction Amount", "Amount"], df.columns)
    bank_col = match_column(["Bank/FIs", "Bank"], df.columns)
    ifsc_col = match_column(["IFSC Code", "Ifsc Code"], df.columns)

    if not sender_col or not receiver_col or not amount_col:
        st.error("❌ Missing required columns.")
        st.stop()

    df[amount_col] = pd.to_numeric(df[amount_col].astype(str).str.replace(",", "").str.replace("₹", ""), errors="coerce")
    df = df[df[amount_col] > 50000]

    victim = df[sender_col].value_counts().idxmax()

    wb = Workbook()
    ws = wb.active
    ws.title = "Flowchart"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=50)
    ws.cell(row=1, column=1, value=f"Victim: {victim}")

    seen_l1 = set()
    seen_l2 = set()
    col = 1

    layer1_df = df[df[sender_col] == victim].drop_duplicates(subset=receiver_col)
    for _, l1_row in layer1_df.iterrows():
        l1_receiver = l1_row[receiver_col]
        if l1_receiver in seen_l1:
            continue
        seen_l1.add(l1_receiver)

        row = 3
        l1_text = format_block(l1_row, bank_col, receiver_col, ifsc_col, l1_row[amount_col], "Sent")
        ws.cell(row=row, column=col, value=l1_text)
        ws.cell(row=row+1, column=col, value="↓")

        layer2_df = df[df[sender_col] == l1_receiver].drop_duplicates(subset=receiver_col)
        for _, l2_row in layer2_df.iterrows():
            l2_receiver = l2_row[receiver_col]
            if l2_receiver in seen_l2:
                continue
            seen_l2.add(l2_receiver)

            l2_text = format_block(l2_row, bank_col, receiver_col, ifsc_col, l2_row[amount_col], "Received")
            ws.cell(row=row+2, column=col, value=l2_text)
            ws.cell(row=row+3, column=col, value="↓")
            row += 4

            # Withdrawal under this L2
            wd_df = df[(df[sender_col] == l2_receiver) & (df[receiver_col].isna())]
            for _, wd in wd_df.iterrows():
                amt = wd[amount_col]
                wd_text = f"💸 Withdrawal Made\nFrom: Layer 2\nA/c No: {l2_receiver}\nAmount: ₹{int(amt):,}"
                ws.cell(row=row, column=col, value=wd_text)
                row += 2

        col += 2

    for c in ws.columns:
        for cell in c:
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    outname = uploaded_file.name.replace(".xlsx", "_final_unique_flowchart.xlsx")
    st.success("✅ Done! Download below:")
    st.download_button("📥 Download Unique Flowchart Excel", data=output, file_name=outname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
