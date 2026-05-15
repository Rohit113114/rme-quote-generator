import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
from pyluach import dates
import gspread
from google.oauth2.service_account import Credentials

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


st.title("RME Quote Generator")


def generate_hebrew_quote_number():
    today = dates.GregorianDate.today()
    hebrew_date = today.to_heb()

    month_codes = {
        1: "NS",
        2: "IY",
        3: "SV",
        4: "TM",
        5: "AV",
        6: "EL",
        7: "TS",
        8: "CH",
        9: "KS",
        10: "TV",
        11: "SH",
        12: "AD",
        13: "A2"
    }

    return f"{hebrew_date.day:02d}{month_codes[hebrew_date.month]}{hebrew_date.year}"


def format_date(date_value):
    if date_value is None:
        return ""
    return date_value.strftime("%d/%m/%Y")


def connect_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(credentials)
    sheet = client.open("RME Quote Register").sheet1

    return sheet


customers_db = pd.read_excel("customers.xlsx")
customers_db.columns = customers_db.columns.str.strip()


st.subheader("Quote Details")

auto_quote_number = generate_hebrew_quote_number()

quote_number = st.text_input(
    "Quote Number",
    value=auto_quote_number
)

revision = st.text_input("Revision", "0")


st.subheader("Internal Workflow Tracking")

job_status = st.selectbox(
    "Job Status",
    [
        "Draft",
        "Released",
        "PO Received",
        "Items Delivered",
        "Invoice Sent",
        "Paid",
        "Completed",
        "Closed"
    ]
)

po_number = st.text_input("PO Number")
invoice_number = st.text_input("Invoice Number")

quote_released_date = st.date_input("Date Quote Released", value=None)
po_received_date = st.date_input("Date PO Received", value=None)
item_delivered_date = st.date_input("Date Item Delivered", value=None)
invoice_sent_date = st.date_input("Date Invoice Sent", value=None)
invoice_due_date = st.date_input("Invoice Due Date", value=None)
invoice_paid_date = st.date_input("Date Invoice Paid", value=None)
job_completed_date = st.date_input("Date Job Completed", value=None)


st.subheader("Customer Details")

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

st.write(f"Name: {selected_customer}")
st.write(f"Department: {department}")
st.write(f"Company: {company}")
st.write(f"Address: {address}")
st.write(f"City/State: {city_state}")


scope = st.text_area("Scope of Work")


st.subheader("Items")

item_count = st.number_input(
    "Number of item rows",
    min_value=1,
    max_value=20,
    value=3
)

items = []

for i in range(item_count):

    st.markdown(f"### Item {i + 1}")

    part_no = st.text_input(f"Part Number {i + 1}", key=f"part{i}")
    description = st.text_input(f"Description {i + 1}", key=f"desc{i}")

    qty = st.number_input(
        f"Qty {i + 1}",
        min_value=0,
        value=0,
        key=f"qty{i}"
    )

    unit_price = st.number_input(
        f"Unit Price {i + 1}",
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


def create_excel_quote():
    wb = load_workbook("rme_excel_template.xlsx")
    ws = wb.active

    ws["F8"] = quote_number
    ws["K8"] = revision
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
        ws[f"C{row}"] = item["description"]
        ws[f"K{row}"] = item["qty"]
        ws[f"L{row}"] = item["unit_price"]
        ws[f"M{row}"] = item["total"]

        ws[f"L{row}"].number_format = '$#,##0.00'
        ws[f"M{row}"].number_format = '$#,##0.00'

    ws["L38"] = subtotal
    ws["L39"] = gst
    ws["L40"] = grand_total

    ws["L38"].number_format = '$#,##0.00'
    ws["L39"].number_format = '$#,##0.00'
    ws["L40"].number_format = '$#,##0.00'

    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    return excel_buffer


def create_pdf_quote():
    pdf_buffer = BytesIO()

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Rail and Marine Engineering Pty Ltd", styles["Title"]))
    elements.append(Paragraph(
        "ACN 656374373 | ABN 82656374373 | Bibra Lake, Western Australia",
        styles["Normal"]
    ))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph(f"Quotation: {quote_number} Rev {revision}", styles["Heading2"]))
    elements.append(Paragraph(f"Date: {datetime.today().strftime('%d/%m/%Y')}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    customer_text = f"""
    <b>Customer</b><br/>
    Name: {selected_customer}<br/>
    Department: {department}<br/>
    Company: {company}<br/>
    Address: {address}<br/>
    City/State: {city_state}
    """
    elements.append(Paragraph(customer_text, styles["Normal"]))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b>Description of work, scope and conditions</b>", styles["Normal"]))
    elements.append(Paragraph(scope, styles["Normal"]))
    elements.append(Spacer(1, 12))

    table_data = [["RME P/N", "Description", "Qty", "$ per unit", "$ Value"]]

    for item in items:
        table_data.append([
            item["part_no"],
            item["description"],
            item["qty"],
            f"${item['unit_price']:,.2f}",
            f"${item['total']:,.2f}"
        ])

    table_data.append(["", "", "", "Sub Total", f"${subtotal:,.2f}"])
    table_data.append(["", "", "", "10% GST", f"${gst:,.2f}"])
    table_data.append(["", "", "", "Total including GST", f"${grand_total:,.2f}"])

    table = Table(table_data, colWidths=[90, 300, 60, 100, 100])

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.black),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (2, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph("Contact Details", styles["Heading3"]))
    elements.append(Paragraph(
        "Rohit Saini | Mechanical Engineer | +610481247284 | rohit@rmerail.com",
        styles["Normal"]
    ))

    doc.build(elements)

    pdf_buffer.seek(0)
    return pdf_buffer


if st.button("Generate Quote"):

    excel_file = create_excel_quote()
    pdf_file = create_pdf_quote()

    history_row = [
        quote_number,
        revision,
        datetime.today().strftime("%d/%m/%Y"),
        selected_customer,
        department,
        company,
        job_status,
        po_number,
        invoice_number,
        format_date(quote_released_date),
        format_date(po_received_date),
        format_date(item_delivered_date),
        format_date(invoice_sent_date),
        format_date(invoice_due_date),
        format_date(invoice_paid_date),
        format_date(job_completed_date),
        subtotal,
        gst,
        grand_total
    ]

    try:
        sheet = connect_google_sheet()
        sheet.append_row(history_row, value_input_option="USER_ENTERED")
        st.success("Quote generated and saved to Google Sheets")
    except Exception as e:
        st.warning("Quote generated, but Google Sheet save failed.")
        st.error(e)

    st.download_button(
        label="Download Quote Excel",
        data=excel_file,
        file_name=f"RME_Quote_{quote_number}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Download Quote PDF",
        data=pdf_file,
        file_name=f"RME_Quote_{quote_number}.pdf",
        mime="application/pdf"
    )


st.subheader("Quote Register")

try:
    sheet = connect_google_sheet()
    records = sheet.get_all_records()

    if records:
        register_df = pd.DataFrame(records)
        st.dataframe(register_df)

        csv_data = register_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="Download Quote Register",
            data=csv_data,
            file_name="rme_quote_register.csv",
            mime="text/csv"
        )
    else:
        st.write("No quote records found yet.")

except Exception:
    st.write("Quote register will appear here after Google Sheets connection is active.")
