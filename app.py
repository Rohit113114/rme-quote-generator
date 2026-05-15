import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, date
from io import BytesIO
from pyluach import dates
import gspread
from google.oauth2.service_account import Credentials
import smtplib
from email.message import EmailMessage
import time
import random

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


st.set_page_config(page_title="RME Commercial Dashboard", layout="wide")
st.title("RME Commercial Dashboard")


REGISTER_HEADERS = [
    "Quote Number", "Revision", "Created Date", "Customer", "Department",
    "Company", "Job Status", "PO Number", "Invoice Number",
    "Quote Released Date", "PO Received Date", "Item Delivered Date",
    "Invoice Sent Date", "Invoice Due Date", "Invoice Paid Date",
    "Job Completed Date", "Subtotal", "GST", "Total"
]

STATUS_OPTIONS = [
    "Draft", "Released", "PO Received", "Items Delivered",
    "Invoice Sent", "Paid", "Completed", "Closed"
]


def generate_hebrew_quote_number():
    today = dates.GregorianDate.today()
    hebrew_date = today.to_heb()

    month_codes = {
        1: "NS", 2: "IY", 3: "SV", 4: "TM", 5: "AV", 6: "EL",
        7: "TS", 8: "CH", 9: "KS", 10: "TV", 11: "SH",
        12: "AD", 13: "A2"
    }

    return f"{hebrew_date.day:02d}{month_codes[hebrew_date.month]}{hebrew_date.year}"


def format_date(date_value):
    if date_value is None:
        return ""
    return date_value.strftime("%d/%m/%Y")


def parse_date(value):
    try:
        if value is None or str(value).strip() == "":
            return None
        return datetime.strptime(str(value), "%d/%m/%Y").date()
    except Exception:
        return None


def connect_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    max_retries = 5
    maximum_backoff = 32

    for retry in range(max_retries):
        try:
            client = gspread.authorize(credentials)
            sheet = client.open("RME Quote Register").sheet1
            return sheet

        except Exception:
            wait_time = min(
                (2 ** retry) + random.uniform(0, 1),
                maximum_backoff
            )

            st.warning(
                f"Google Sheets is busy. Retrying in {wait_time:.1f} seconds..."
            )

            time.sleep(wait_time)

    raise Exception("Google Sheets connection failed after multiple retries.")


def get_register_dataframe():
    sheet = connect_google_sheet()
    records = sheet.get_all_records()

    if records:
        return pd.DataFrame(records)

    return pd.DataFrame(columns=REGISTER_HEADERS)


def update_register_row(quote_number, revision, update_values):
    sheet = connect_google_sheet()
    all_values = sheet.get_all_values()

    if not all_values:
        return False

    headers = all_values[0]

    quote_col_index = headers.index("Quote Number") + 1
    revision_col_index = headers.index("Revision") + 1

    row_to_update = None

    for row_number, row in enumerate(all_values[1:], start=2):
        quote_match = (
            len(row) >= quote_col_index
            and str(row[quote_col_index - 1]) == str(quote_number)
        )
        revision_match = (
            len(row) >= revision_col_index
            and str(row[revision_col_index - 1]) == str(revision)
        )

        if quote_match and revision_match:
            row_to_update = row_number
            break

    if row_to_update is None:
        return False

    for column_name, value in update_values.items():
        if column_name in headers:
            col_index = headers.index(column_name) + 1
            sheet.update_cell(row_to_update, col_index, value)

    return True


def send_email_with_pdf(to_email, subject, body, pdf_file, pdf_filename):
    if "smtp" not in st.secrets:
        raise Exception("SMTP email settings are not configured in Streamlit Secrets.")

    smtp_host = st.secrets["smtp"]["host"]
    smtp_port = int(st.secrets["smtp"]["port"])
    smtp_user = st.secrets["smtp"]["username"]
    smtp_password = st.secrets["smtp"]["password"]
    from_email = st.secrets["smtp"]["from_email"]

    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    msg.add_attachment(
        pdf_file.getvalue(),
        maintype="application",
        subtype="pdf",
        filename=pdf_filename
    )

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as smtp:
        smtp.login(smtp_user, smtp_password)
        smtp.send_message(msg)


customers_db = pd.read_excel("customers.xlsx")
customers_db.columns = customers_db.columns.str.strip()


tab_dashboard, tab_create, tab_update, tab_register = st.tabs(
    ["Dashboard", "Create New Quote", "Update Existing Quote", "Quote Register"]
)


with tab_dashboard:

    st.subheader("RME Quote Dashboard")

    try:

        dashboard_df = get_register_dataframe()

        if dashboard_df.empty:
            st.write("No quote data available.")

        else:

            for money_col in ["Subtotal", "GST", "Total"]:

                dashboard_df[money_col] = (
                    dashboard_df[money_col]
                    .astype(str)
                    .str.replace("$", "", regex=False)
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )

                dashboard_df[money_col] = pd.to_numeric(
                    dashboard_df[money_col],
                    errors="coerce"
                ).fillna(0)

            total_quotes = len(dashboard_df)

            total_revenue = dashboard_df["Total"].sum()

            paid_jobs = len(
                dashboard_df[
                    dashboard_df["Job Status"] == "Paid"
                ]
            )

            po_received = len(
                dashboard_df[
                    dashboard_df["Job Status"] == "PO Received"
                ]
            )

            invoice_sent = len(
                dashboard_df[
                    dashboard_df["Job Status"] == "Invoice Sent"
                ]
            )

            today_date = date.today()

            dashboard_df["Parsed Due Date"] = dashboard_df[
                "Invoice Due Date"
            ].apply(parse_date)

            dashboard_df["Paid Blank"] = (
                dashboard_df["Invoice Paid Date"]
                .astype(str)
                .str.strip() == ""
            )

            overdue_df = dashboard_df[
                (dashboard_df["Parsed Due Date"].notna()) &
                (dashboard_df["Parsed Due Date"] < today_date) &
                (dashboard_df["Paid Blank"])
            ]

            overdue_invoices = len(overdue_df)

            col1, col2, col3 = st.columns(3)

            col1.metric("Total Quotes", total_quotes)

            col2.metric(
                "Total Revenue",
                f"${total_revenue:,.2f}"
            )

            col3.metric("Paid Jobs", paid_jobs)

            col4, col5, col6 = st.columns(3)

            col4.metric("PO Received", po_received)

            col5.metric("Invoices Sent", invoice_sent)

            col6.metric("Overdue Invoices", overdue_invoices)

            st.subheader("Quote Search")

            search_quote = st.text_input(
                "Search Quote Number"
            )

            search_customer = st.text_input(
                "Search Customer"
            )

            search_po = st.text_input(
                "Search PO Number"
            )

            search_status = st.selectbox(
                "Filter Job Status",
                ["All"] + STATUS_OPTIONS
            )

            filtered_df = dashboard_df.copy()

            if search_quote:

                filtered_df = filtered_df[
                    filtered_df["Quote Number"]
                    .astype(str)
                    .str.contains(
                        search_quote,
                        case=False,
                        na=False
                    )
                ]

            if search_customer:

                filtered_df = filtered_df[
                    filtered_df["Customer"]
                    .astype(str)
                    .str.contains(
                        search_customer,
                        case=False,
                        na=False
                    )
                ]

            if search_po:

                filtered_df = filtered_df[
                    filtered_df["PO Number"]
                    .astype(str)
                    .str.contains(
                        search_po,
                        case=False,
                        na=False
                    )
                ]

            if search_status != "All":

                filtered_df = filtered_df[
                    filtered_df["Job Status"]
                    .astype(str) == search_status
                ]

            filtered_df = filtered_df.drop(
                columns=["Parsed Due Date", "Paid Blank"],
                errors="ignore"
            )

            st.subheader("Quote Results")

display_df = filtered_df.copy()

for money_col in ["Subtotal", "GST", "Total"]:

    display_df[money_col] = display_df[money_col].apply(
        lambda x: f"${x:,.2f}"
    )

st.dataframe(
    display_df,
    use_container_width=True
)

            if overdue_invoices > 0:

                st.subheader("Overdue Invoices")

                st.dataframe(
                    overdue_df.drop(
                        columns=[
                            "Parsed Due Date",
                            "Paid Blank"
                        ],
                        errors="ignore"
                    ),
                    use_container_width=True
                )

    except Exception as e:

        st.error("Dashboard failed to load.")
        st.error(e)


with tab_create:
    st.subheader("Quote Details")

    auto_quote_number = generate_hebrew_quote_number()

    quote_number = st.text_input("Quote Number", value=auto_quote_number)
    revision = st.text_input("Revision", "0")

    quote_reference = f"{quote_number}-R{revision}"
    st.info(f"Quote Reference: {quote_reference}")

    st.subheader("Internal Workflow Tracking")

    job_status = st.selectbox("Job Status", STATUS_OPTIONS)

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

    selected_customer = st.selectbox("Customer Contact", customers_db["Contact Name"])

    customer_row = customers_db[
        customers_db["Contact Name"] == selected_customer
    ].iloc[0]

    department = customer_row["Department"]
    company = customer_row["Company"]
    address = customer_row["Address"]
    city_state = customer_row["City/State"]

    customer_email = ""
    if "Email" in customers_db.columns:
        customer_email = str(customer_row["Email"])

    st.write(f"Name: {selected_customer}")
    st.write(f"Department: {department}")
    st.write(f"Company: {company}")
    st.write(f"Address: {address}")
    st.write(f"City/State: {city_state}")

    if customer_email:
        st.write(f"Email: {customer_email}")

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
            file_name=f"RME_Quote_{quote_reference}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download Quote PDF",
            data=pdf_file,
            file_name=f"RME_Quote_{quote_reference}.pdf",
            mime="application/pdf"
        )

        st.subheader("Email Quote")

        email_to = st.text_input("Send to email", value=customer_email)
        email_subject = st.text_input(
            "Email subject",
            value=f"RME Quotation {quote_reference}"
        )
        email_body = st.text_area(
            "Email body",
            value=f"""Hi,

Please find attached RME quotation {quote_reference}.

Regards,
Rohit Saini
Rail and Marine Engineering Pty Ltd"""
        )

        if st.button("Send Quote Email"):
            try:
                send_email_with_pdf(
                    email_to,
                    email_subject,
                    email_body,
                    pdf_file,
                    f"RME_Quote_{quote_reference}.pdf"
                )
                st.success("Email sent successfully.")
            except Exception as e:
                st.error("Email sending failed.")
                st.error(e)


with tab_update:
    st.subheader("Update Existing Quote")

    try:
        register_df = get_register_dataframe()

        if register_df.empty:
            st.write("No quote records found.")
        else:
            register_df["Quote Display"] = (
                register_df["Quote Number"].astype(str) +
                "-R" +
                register_df["Revision"].astype(str)
            )

            quote_options = register_df["Quote Display"].tolist()

            selected_quote_display = st.selectbox(
                "Select Quote",
                quote_options
            )

            selected_record = register_df[
                register_df["Quote Display"] == selected_quote_display
            ].iloc[0]

            selected_quote_number = str(selected_record.get("Quote Number", ""))
            selected_revision = str(selected_record.get("Revision", ""))

            st.write("Customer:", selected_record.get("Customer", ""))
            st.write("Company:", selected_record.get("Company", ""))
            st.write("Current Status:", selected_record.get("Job Status", ""))

            current_status = selected_record.get("Job Status", "Draft")

            if current_status not in STATUS_OPTIONS:
                current_status = "Draft"

            updated_job_status = st.selectbox(
                "Update Job Status",
                STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(current_status)
            )

            updated_po_number = st.text_input(
                "Update PO Number",
                value=str(selected_record.get("PO Number", ""))
            )

            updated_invoice_number = st.text_input(
                "Update Invoice Number",
                value=str(selected_record.get("Invoice Number", ""))
            )

            update_quote_released_date = st.text_input(
                "Quote Released Date",
                value=str(selected_record.get("Quote Released Date", ""))
            )

            update_po_received_date = st.text_input(
                "PO Received Date",
                value=str(selected_record.get("PO Received Date", ""))
            )

            update_item_delivered_date = st.text_input(
                "Item Delivered Date",
                value=str(selected_record.get("Item Delivered Date", ""))
            )

            update_invoice_sent_date = st.text_input(
                "Invoice Sent Date",
                value=str(selected_record.get("Invoice Sent Date", ""))
            )

            update_invoice_due_date = st.text_input(
                "Invoice Due Date",
                value=str(selected_record.get("Invoice Due Date", ""))
            )

            update_invoice_paid_date = st.text_input(
                "Invoice Paid Date",
                value=str(selected_record.get("Invoice Paid Date", ""))
            )

            update_job_completed_date = st.text_input(
                "Job Completed Date",
                value=str(selected_record.get("Job Completed Date", ""))
            )

            if st.button("Update Quote Register"):
                update_values = {
                    "Job Status": updated_job_status,
                    "PO Number": updated_po_number,
                    "Invoice Number": updated_invoice_number,
                    "Quote Released Date": update_quote_released_date,
                    "PO Received Date": update_po_received_date,
                    "Item Delivered Date": update_item_delivered_date,
                    "Invoice Sent Date": update_invoice_sent_date,
                    "Invoice Due Date": update_invoice_due_date,
                    "Invoice Paid Date": update_invoice_paid_date,
                    "Job Completed Date": update_job_completed_date
                }

                success = update_register_row(
                    selected_quote_number,
                    selected_revision,
                    update_values
                )

                if success:
                    st.success("Quote register updated successfully.")
                else:
                    st.error("Quote number/revision not found in register.")

    except Exception as e:
        st.error("Could not load quote register.")
        st.error(e)


with tab_register:
    st.subheader("Quote Register")

    try:
        register_df = get_register_dataframe()

        if not register_df.empty:
            st.dataframe(register_df, use_container_width=True)

            csv_data = register_df.to_csv(index=False).encode("utf-8")

            st.download_button(
                label="Download Quote Register",
                data=csv_data,
                file_name="rme_quote_register.csv",
                mime="text/csv"
            )
        else:
            st.write("No quote records found yet.")

    except Exception as e:
        st.write("Quote register will appear here after Google Sheets connection is active.")
        st.error(e)
