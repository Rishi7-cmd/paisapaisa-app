import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

def generate_flowchart(df):
    df = df.rename(columns={
        'Account No./ (Wallet /PG/PA) Id': 'Sender',
        'Account No': 'Receiver',
        'Transaction Amount': 'Amount',
        'Bank/FIs': 'Bank',
        'Ifsc Code': 'IFSC'
    })

    df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(",", "").str.replace("₹", ""), errors='coerce')
    df = df[df['Amount'] > 50000]

    victim = df['Sender'].value_counts().idxmax()
    layer1 = df[df['Sender'] == victim]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transaction Flow"

    blue = PatternFill(start_color="B7D5F4", end_color="B7D5F4", fill_type="solid")
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    violet = PatternFill(start_color="E4D7F5", end_color="E4D7F5", fill_type="solid")
    yellow = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    redfont = Font(color="9C0006")

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=50)
    ws.cell(row=row, column=1, value=f"Victim: {victim}").fill = blue
    row += 2

    for l1_acc in layer1['Receiver'].unique():
        ws.cell(row=row, column=2, value="↓")
        row += 1
        l1_txns = df[df['Sender'] == l1_acc]
        ws.cell(row=row, column=2, value=f"L1: {l1_acc}").fill = green
        row += 2

        for l2_acc in l1_txns['Receiver'].unique():
            ws.cell(row=row, column=3, value="↓")
            row += 1
            ws.cell(row=row, column=3, value=f"L2: {l2_acc}").fill = violet
            row += 1
            l2_wd = df[(df['Sender'] == l2_acc) & (df['Receiver'].isna())]
            for _, wd in l2_wd.iterrows():
                row += 1
                cell = ws.cell(row=row, column=4, value=f"Withdraw: ₹{wd['Amount']} {wd['Bank']} ({wd['IFSC']})")
                cell.fill = yellow
                if wd['Amount'] > 100000:
                    cell.font = redfont

        l1_wd = df[(df['Sender'] == l1_acc) & (df['Receiver'].isna())]
        for _, wd in l1_wd.iterrows():
            row += 1
            cell = ws.cell(row=row, column=3, value=f"Withdraw: ₹{wd['Amount']} {wd['Bank']} ({wd['IFSC']})")
            cell.fill = yellow
            if wd['Amount'] > 100000:
                cell.font = redfont

        row += 2

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit UI
st.set_page_config(page_title="Paisa Paisa", layout="wide")

st.markdown("""
    <style>
    body {
        background-image: url("https://images.unsplash.com/photo-1604079628044-b0c5e5033183");
        background-size: cover;
        background-attachment: fixed;
        color: white;
    }
    .main, .block-container {
        background: rgba(0, 0, 0, 0.7);
        padding: 2rem;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='color:#FFD700;'>🌟 Paisa Paisa Transaction Flow Analyzer 🚀</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload the input Excel file", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        output_bytes = generate_flowchart(df)
        st.success("Flowchart generated successfully!")
        st.download_button("📄 Download Flowchart Excel", data=output_bytes, file_name="Flowchart_Output.xlsx")
    except Exception as e:
        st.error(f"Error: {e}")
