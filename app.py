import streamlit as st
import pandas as pd
from datetime import datetime
import sqlite3

st.set_page_config(page_title="DC Email System", layout="wide")

# ================= DATABASE =================
conn = sqlite3.connect("emails.db", check_same_thread=False)

conn.execute("""
CREATE TABLE IF NOT EXISTS emails (
id INTEGER PRIMARY KEY AUTOINCREMENT,
created_at TEXT,
project_code TEXT,
document_number TEXT,
document_type TEXT,
title TEXT,
revision TEXT,
status TEXT,
recipient TEXT,
sender_name TEXT,
company_name TEXT,
language TEXT,
subject TEXT,
body TEXT,
source TEXT
)
""")
conn.commit()

def save_email(**kwargs):
    data = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        **kwargs
    }

    cols = list(data.keys())
    values = list(data.values())
    placeholders = ",".join(["?"] * len(cols))

    conn.execute(
        f"INSERT INTO emails ({','.join(cols)}) VALUES ({placeholders})",
        values,
    )
    conn.commit()

def load_emails():
    return pd.read_sql("SELECT * FROM emails ORDER BY id DESC", conn)

# ================= DECISION ENGINE =================
def generate_email(data):
    status = data["status"]
    lang = data["language"]

    # Clean values
    project = data["project_code"] or "-"
    doc_no = data["document_number"] or "-"
    title = data["title"] or "-"
    rev = data["revision"] or "-"
    doc_type = data["document_type"]
    recipient = data["recipient"]
    sender = data["sender_name"]

    if lang == "English":

        if status in ["Pending", "Under Review"]:
            subject_prefix = "[REMINDER]"
            body = f"""Dear {recipient},

We are writing to follow up on the review status of the {doc_type.lower()} "{title}" ({doc_no}, {rev}).

Kindly provide your update at your earliest convenience to ensure project continuity.

Best regards,
{sender}"""

        elif status == "Approved":
            subject_prefix = "[APPROVED]"
            body = f"""Dear {recipient},

We are pleased to inform you that the {doc_type.lower()} "{title}" ({doc_no}, {rev}) has been approved.

Best regards,
{sender}"""

        elif status == "Rejected":
            subject_prefix = "[REJECTED]"
            body = f"""Dear {recipient},

The {doc_type.lower()} "{title}" ({doc_no}) has been rejected.

Kindly review and resubmit.

Best regards,
{sender}"""

        elif status == "Submitted":
            subject_prefix = "[SUBMIT]"
            body = f"""Dear {recipient},

Please find attached {doc_type} "{title}" ({doc_no}) for your review.

Best regards,
{sender}"""

        else:
            subject_prefix = "[INFO]"
            body = "Status update."

    else:
        subject_prefix = "[متابعة]"
        body = f"""السيد {recipient},

نود المتابعة على المستند "{title}" رقم {doc_no}.

مع التحية،
{sender}"""

    subject = f"{subject_prefix} {project} - {doc_type} - {doc_no} - {title} - {status}"
    subject = subject.replace("  ", " ")

    return subject, body

# ================= COPY BUTTON =================
def copy_block(label, text):
    st.code(text)
    st.button(label)

# ================= UI =================
menu = st.sidebar.radio("Menu", ["Single", "Bulk", "Dashboard", "History"])

# ================= SINGLE =================
if menu == "Single":

    st.title("📧 Single Email")

    col1, col2 = st.columns(2)

    with col1:
        project = st.text_input("Project Code", "NBT-24")
        doc_no = st.text_input("Document Number", "NBT-SUB-101")
        doc_type = st.selectbox("Document Type", ["Submittal", "RFI"])
        title = st.text_input("Title", "Sample Title")
        rev = st.text_input("Revision", "Rev.00")

    with col2:
        status = st.selectbox("Status", ["Pending", "Under Review", "Approved", "Rejected", "Submitted"])
        recipient = st.text_input("Recipient", "Consultant")
        sender = st.text_input("Sender Name", "Your Name")
        company = st.text_input("Company", "Your Company")
        lang = st.selectbox("Language", ["English", "Arabic"])

    if st.button("Generate Email"):

        if not project or not doc_no or not title or not sender:
            st.error("❌ Please fill required fields")
            st.stop()

        data = {
            "project_code": project,
            "document_number": doc_no,
            "document_type": doc_type,
            "title": title,
            "revision": rev,
            "status": status,
            "recipient": recipient,
            "sender_name": sender,
            "company_name": company,
            "language": lang
        }

        subject, body = generate_email(data)

        st.success("Generated")

        st.subheader("Preview")
        st.text(f"Subject: {subject}")
        st.text(body)

        st.subheader("Copy")
        copy_block("Copy Subject", subject)
        copy_block("Copy Body", body)

        save_email(
            project_code=project,
            document_number=doc_no,
            document_type=doc_type,
            title=title,
            revision=rev,
            status=status,
            recipient=recipient,
            sender_name=sender,
            company_name=company,
            language=lang,
            subject=subject,
            body=body,
            source="single"
        )

# ================= BULK =================
elif menu == "Bulk":

    st.title("📂 Bulk Emails")

    sender = st.text_input("Sender Name")
    company = st.text_input("Company")
    lang = st.selectbox("Language", ["English", "Arabic"])

    file = st.file_uploader("Upload Excel", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        st.dataframe(df)

        if st.button("Generate Emails"):

            results = []

            for _, row in df.iterrows():

                data = {
                    "project_code": row.get("Project Code", ""),
                    "document_number": row.get("Document Number", ""),
                    "document_type": row.get("Document Type", ""),
                    "title": row.get("Title", ""),
                    "revision": row.get("Revision", ""),
                    "status": row.get("Status", ""),
                    "recipient": "Consultant",
                    "sender_name": sender,
                    "company_name": company,
                    "language": lang
                }

                subject, body = generate_email(data)

                results.append({
                    "Doc": data["document_number"],
                    "Status": data["status"],
                    "Subject": subject
                })

                save_email(
                    project_code=data["project_code"],
                    document_number=data["document_number"],
                    document_type=data["document_type"],
                    title=data["title"],
                    revision=data["revision"],
                    status=data["status"],
                    recipient=data["recipient"],
                    sender_name=sender,
                    company_name=company,
                    language=lang,
                    subject=subject,
                    body=body,
                    source="bulk"
                )

            st.success("Done")
            st.dataframe(pd.DataFrame(results))

# ================= DASHBOARD =================
elif menu == "Dashboard":

    st.title("📊 Dashboard")

    df = load_emails()

    st.metric("Total Emails", len(df))
    st.write(df["status"].value_counts())

# ================= HISTORY =================
elif menu == "History":

    st.title("📁 History")

    df = load_emails()
    st.dataframe(df)

    if len(df) > 0:
        selected = st.selectbox("Select Email", df["id"])
        email = df[df["id"] == selected].iloc[0]

        st.subheader("Preview")
        st.text(email["subject"])
        st.text(email["body"])
