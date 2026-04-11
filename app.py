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

    if lang == "English":

        if status in ["Pending", "Under Review"]:
            subject_prefix = "[REMINDER]"
            body = f"""Dear {data['recipient']},

We are writing to follow up on the review status of the {data['document_type'].lower()} for "{data['title']}" ({data['document_number']}, {data['revision']}).

Kindly provide your update or the reviewed outcome at your earliest convenience to ensure the continuity of the project workflow.

Your prompt attention to this matter is highly appreciated.

Best regards,
{data['sender_name']}"""

        elif status == "Approved":
            subject_prefix = "[APPROVED]"
            body = f"""Dear {data['recipient']},

We are pleased to inform you that the {data['document_type'].lower()} "{data['title']}" ({data['document_number']}) has been approved.

Best regards,
{data['sender_name']}"""

        elif status == "Rejected":
            subject_prefix = "[REJECTED]"
            body = f"""Dear {data['recipient']},

Please be informed that the {data['document_type'].lower()} "{data['title']}" ({data['document_number']}) has been rejected.

Kindly review and resubmit.

Best regards,
{data['sender_name']}"""

        elif status == "Submitted":
            subject_prefix = "[SUBMIT]"
            body = f"""Dear {data['recipient']},

Please find attached {data['document_type']} {data['document_number']} for your review.

Best regards,
{data['sender_name']}"""

        else:
            subject_prefix = "[INFO]"
            body = "Status update."

    else:  # عربي
        subject_prefix = "[متابعة]"
        body = f"""السيد {data['recipient']},

نود المتابعة على حالة المستند "{data['title']}" رقم {data['document_number']}.

يرجى التكرم بالتحديث.

مع التحية،
{data['sender_name']}"""

    subject = f"{subject_prefix} {data['project_code']} - {data['document_type']} - {data['document_number']} - {data['title']} - {data['status']}"

    return subject, body


# ================= COPY BUTTON =================
def copy_button(text, label):
    st.code(text)
    st.button(label)


# ================= UI =================
menu = st.sidebar.radio("Menu", ["Single", "Bulk", "Dashboard", "History"])

# ================= SINGLE =================
if menu == "Single":

    st.title("📧 Single Email Generator")

    col1, col2 = st.columns(2)

    with col1:
        project = st.text_input("Project Code")
        doc_no = st.text_input("Document Number")
        doc_type = st.selectbox("Document Type", ["Submittal", "RFI"])
        title = st.text_input("Title")
        rev = st.text_input("Revision", "Rev.00")

    with col2:
        status = st.selectbox("Status", ["Pending", "Under Review", "Approved", "Rejected", "Submitted"])
        recipient = st.text_input("Recipient", "Consultant")
        sender = st.text_input("Sender Name")
        company = st.text_input("Company")
        lang = st.selectbox("Language", ["English", "Arabic"])

    if st.button("Generate Email"):

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
        copy_button(subject, "Copy Subject")
        copy_button(body, "Copy Body")

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
    company = st.text_input("Company Name")
    lang = st.selectbox("Language", ["English", "Arabic"])

    file = st.file_uploader("Upload Excel", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        st.dataframe(df)

        if st.button("Generate Emails"):

            results = []

            for _, row in df.iterrows():
                data = {
                    "project_code": row["Project Code"],
                    "document_number": row["Document Number"],
                    "document_type": row["Document Type"],
                    "title": row["Title"],
                    "revision": row["Revision"],
                    "status": row["Status"],
                    "recipient": "Consultant",
                    "sender_name": sender,
                    "company_name": company,
                    "language": lang
                }

                subject, body = generate_email(data)

                results.append({
                    "Doc": row["Document Number"],
                    "Status": row["Status"],
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

    st.title("📁 Saved Emails")

    df = load_emails()

    st.dataframe(df)

    selected = st.selectbox("Select Email", df["id"])

    email = df[df["id"] == selected].iloc[0]

    st.subheader("Preview")
    st.text(email["subject"])
    st.text(email["body"])
