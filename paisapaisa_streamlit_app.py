
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font

st.set_page_config(page_title="Paisa Paisa Final Flowchart", page_icon="📊", layout="wide")
st.markdown("<h1 style='text-align: center; color: #FFD700;'>📊 Final Stylized Flowchart: Victim → L1 → L2 → Withdrawal</h1>", unsafe_allow_html=True)

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

    # Styles
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    fill_blue = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
    fill_green = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_violet = PatternFill(start_color='E4DFEC', end_color='E4DFEC', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    # Merge victim row across columns
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=50)
    victim_cell = ws.cell(row=1, column=1, value=f"Victim: {victim}")
    victim_cell.fill = fill_blue
    victim_cell.alignment = align_center
    victim_cell.font = Font(bold=True)

    start_col = 1
    used_l1 = set()
    layer1_df = df[(df[sender_col] == victim) & (df[receiver_col] != victim)]

    for _, l1_row in layer1_df.iterrows():
        l1_receiver = l1_row[receiver_col]
        if l1_receiver in used_l1:
            continue
        used_l1.add(l1_receiver)

        row = 3
        ws.cell(row=row, column=start_col, value="↓").alignment = align_center
        row += 1

        l1_text = format_block(l1_row, bank_col, receiver_col, ifsc_col, l1_row[amount_col], "Sent")
        cell = ws.cell(row=row, column=start_col, value=l1_text)
        cell.fill = fill_green
        cell.alignment = align_center
        row += 2

        ws.cell(row=row, column=start_col, value="↓").alignment = align_center
        row += 1

        used_l2 = set()
        layer2_df = df[(df[sender_col] == l1_receiver) & (df[receiver_col] != l1_receiver)]
        for _, l2_row in layer2_df.iterrows():
            l2_receiver = l2_row[receiver_col]
            if l2_receiver in used_l2 or l2_receiver == victim or pd.isna(l2_receiver):
                continue
            used_l2.add(l2_receiver)

            l2_text = format_block(l2_row, bank_col, receiver_col, ifsc_col, l2_row[amount_col], "Received")
            cell = ws.cell(row=row, column=start_col, value=l2_text)
            cell.fill = fill_violet
            cell.alignment = align_center
            row += 2

            ws.cell(row=row, column=start_col, value="↓").alignment = align_center
            row += 1

            wd_df = df[(df[sender_col] == l2_receiver) & (df[receiver_col].isna())]
            for _, wd in wd_df.iterrows():
                amt = wd[amount_col]
                text = f"💸 Withdrawal Made\nFrom: Layer 2\nA/c No: {l2_receiver}\nAmount: ₹{int(amt):,}"
                cell = ws.cell(row=row, column=start_col, value=text)
                cell.fill = fill_yellow if amt <= 100000 else fill_red
                cell.alignment = align_center
                row += 2

        start_col += 2

    for col in ws.columns:
        for cell in col:
            cell.alignment = align_center

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    outname = uploaded_file.name.replace(".xlsx", "_styled_flowchart.xlsx")
    st.success("✅ Flowchart generated with full formatting.")
    st.download_button("📥 Download Flowchart Excel", data=output, file_name=outname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
