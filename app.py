import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import win32com.client
import pythoncom

st.title("RME Quote Generator")

customers_db = pd.read_excel("customers.xlsx")
customers_db.columns = customers_db.columns.str.strip()

quote_number = st.text_input("Quote Number")
revision = st.text_input("Revision", "0")

selected_customer = st.selectbox(
    "Customer Contact",
    customers_db["Contact Name"]
)

customer_row = customers_db[
    customers_db["Contact Name"] == selected_customer
].iloc[0]

department = customer_row["Department"]
company = customer_row["Company"]
address = customer_row["Address"]
city_state = customer_row["City/State"]

st.subheader("Customer Details")
st.write(f"Name: {selected_customer}")
st.write(f"Department: {department}")
st.write(f"Company: {company}")
st.write(f"Address: {address}")
st.write(f"City/State: {city_state}")

scope = st.text_area("Scope of Work")

st.subheader("Items")

items = []

for i in range(10):

    st.markdown(f"### Item {i+1}")

    part_no = st.text_input(f"Part Number {i+1}", key=f"part{i}")
    description = st.text_input(f"Description {i+1}", key=f"desc{i}")

    qty = st.number_input(
        f"Qty {i+1}",
        min_value=0,
        value=0,
        key=f"qty{i}"
    )

    unit_price = st.number_input(
        f"Unit Price {i+1}",
        min_value=0.0,
        value=0.0,
        key=f"price{i}"
    )

    if qty > 0:
        total = qty * unit_price

        items.append({
            "part_no": part_no,
            "description": description,
            "qty": qty,
            "unit_price": unit_price,
            "total": total
        })

subtotal = sum(item["total"] for item in items)
gst = subtotal * 0.10
grand_total = subtotal + gst

st.subheader("Totals")
st.write(f"Subtotal: ${subtotal:,.2f}")
st.write(f"GST: ${gst:,.2f}")
st.write(f"Grand Total: ${grand_total:,.2f}")

if st.button("Generate Quote"):

    wb = load_workbook("rme_excel_template.xlsx")
    ws = wb.active

    ws["F8"] = quote_number
    ws["F8"].number_format = "General"

    ws["K8"] = revision
    ws["K8"].number_format = "General"

    ws["F9"] = datetime.today().strftime("%d/%m/%Y")

    ws["C13"] = selected_customer
    ws["C14"] = department
    ws["C15"] = company
    ws["C16"] = address
    ws["C17"] = city_state

    ws["B20"] = scope

    start_row = 26

    for index, item in enumerate(items):

        row = start_row + index

        ws[f"B{row}"] = item["part_no"]
        ws[f"B{row}"].number_format = "General"

        ws[f"C{row}"] = item["description"]

        ws[f"K{row}"] = item["qty"]
        ws[f"K{row}"].number_format = "General"

        ws[f"L{row}"] = item["unit_price"]
        ws[f"L{row}"].number_format = '$#,##0.00'

        ws[f"M{row}"] = item["total"]
        ws[f"M{row}"].number_format = '$#,##0.00'

    ws["L38"] = subtotal
    ws["L39"] = gst
    ws["L40"] = grand_total

    ws["L38"].number_format = '$#,##0.00'
    ws["L39"].number_format = '$#,##0.00'
    ws["L40"].number_format = '$#,##0.00'

    output_excel = f"output/RME_Quote_{quote_number}.xlsx"
    wb.save(output_excel)

    try:
        pythoncom.CoInitialize()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        workbook_path = fr"C:\Users\rohit\opencv_projects\RME QUOTE GENERATOR\{output_excel}"
        pdf_path = fr"C:\Users\rohit\opencv_projects\RME QUOTE GENERATOR\output\RME_Quote_{quote_number}.pdf"

        workbook = excel.Workbooks.Open(workbook_path)
        worksheet = workbook.Worksheets(1)

        worksheet.ExportAsFixedFormat(0, pdf_path)

        workbook.Close(False)
        excel.Quit()

        st.success(
            f"Excel and PDF Generated Successfully:\n{output_excel}"
        )

    except Exception as e:
        st.warning(
            "Excel quote was generated, but PDF export failed. "
            "Close any open Excel files and try again."
        )
        st.error(e)
