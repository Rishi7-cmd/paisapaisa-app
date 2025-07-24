
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

if uploaded_file:
    input_df = pd.read_excel(uploaded_file)

    # FILTERS
    input_df = input_df[input_df["Transaction Amount"] > 50000]

    # Grouping and tracing logic
    victim = input_df["Sender account"].value_counts().idxmax()
    layer1_df = input_df[input_df["Sender account"] == victim]
    layer1_receivers = layer1_df["Receiver account"].unique()

    output_data = []
    withdrawals = []

    for l1 in layer1_receivers:
        l1_trans = input_df[input_df["Sender account"] == l1]
        layer2_receivers = l1_trans["Receiver account"].unique()
        l2_data = []

        for l2 in layer2_receivers:
            wd = input_df[(input_df["Sender account"] == l2) & (input_df["Receiver account"].isna())]
            if not wd.empty:
                amount = wd["Amount"].sum()
                highlight = "⚠️" if amount > 100000 else ""
                l2_data.append(f"{l2}\n💰 {amount}{highlight}")
                withdrawals.append((l2, amount))

        if l2_data:
            output_data.append([f"{victim}", f"{l1}", "\n\n".join(l2_data)])

    # Create Excel output
    wb = Workbook()
    ws = wb.active
    ws.title = "Layered Flow"

    headers = ["Victim", "Layer 1", "Layer 2 + Withdrawals"]
    ws.append(headers)

    for row in output_data:
        ws.append(row)

    # Highlight withdrawals > 1 lakh
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for idx, row in enumerate(ws.iter_rows(min_row=2, max_col=3), start=2):
        if "⚠️" in str(row[2].value):
            for cell in row:
                cell.fill = yellow_fill

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    output_filename = uploaded_file.name.replace(".xlsx", "_output.xlsx")
    st.success("✅ Analysis complete! Download your output Excel file below:")
    st.download_button(label="📥 Download Output Excel",
                       data=output,
                       file_name=output_filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
