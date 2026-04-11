import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(page_title="DC Email System", layout="wide")

DB = "emails.db"

# =========================================================
# DATABASE
# =========================================================
conn = sqlite3.connect(DB, check_same_thread=False)

def init_db():
    conn.execute("""
    CREATE TABLE IF NOT EXISTS emails (
        id INTEGER PRIMARY KEY,
        created_at TEXT,
        project_code TEXT,
        document_number TEXT,
        document_type TEXT,
        title TEXT,
        revision TEXT,
        status TEXT,
        recipient TEXT,
        subject TEXT,
        body TEXT,
        language TEXT,
        source TEXT
    )
    """)
    conn.commit()

init_db()

def save_email(**kwargs):
    cols = [
        "created_at","project_code","document_number","document_type",
        "title","revision","status","recipient","subject","body","language","source"
    ]

    values = []
    for c in cols:
        if c == "created_at":
            values.append(datetime.now().strftime("%Y-%m-%d %H:%M"))
        else:
            values.append(kwargs.get(c, ""))

    q = ",".join(["?"] * len(cols))

    conn.execute(
        f"INSERT INTO emails ({','.join(cols)}) VALUES ({q})",
        values
    )
    conn.commit()

def load_emails():
    return pd.read_sql("SELECT * FROM emails ORDER BY id DESC", conn)

# =========================================================
# DECISION ENGINE
# =========================================================
def subject_builder(p, t, no, title, status):
    prefix = {
        "Pending": "[REMINDER]",
        "Under Review": "[REMINDER]",
        "Approved": "[APPROVED]",
        "Approved with Comments": "[COMMENTS]",
        "Rejected": "[REJECTED]",
        "Submitted": "[SUBMIT]"
    }
    return f"{prefix.get(status,'')} {p} - {t} - {no} - {title} - {status}"

def email_body(status, title, no, rev, recipient, sender, lang):
    if lang == "English":
        if status in ["Pending","Under Review"]:
            return f"""Dear {recipient},

We are writing to follow up on "{title}" ({no}, {rev}).

Kindly provide your update.

Best regards,
{sender}"""

        if status == "Approved":
            return f"""Dear {recipient},

"{title}" ({no}) has been approved.

Best regards,
{sender}"""

        if status == "Rejected":
            return f"""Dear {recipient},

"{title}" ({no}) has been rejected.

Please revise and resubmit.

Best regards,
{sender}"""

        if status == "Submitted":
            return f"""Dear {recipient},

Please find attached "{title}" ({no}) for your review.

Best regards,
{sender}"""

    else:
        return f"""السادة {recipient}،

بخصوص "{title}" ({no})

يرجى المراجعة.

مع تحياتي،
{sender}"""

# =========================================================
# COPY BUTTON
# =========================================================
def copy_btn(text):
    st.code(text)

# =========================================================
# UI
# =========================================================
menu = st.sidebar.radio("Menu", ["Single","Bulk","Dashboard","History"])

# =========================================================
# SINGLE
# =========================================================
if menu == "Single":

    st.header("Single Email")

    col1, col2 = st.columns(2)

    with col1:
        p = st.text_input("Project Code")
        no = st.text_input("Document No")
        t = st.selectbox("Type", ["Submittal","RFI","Drawing"])
        title = st.text_input("Title")
        rev = st.text_input("Revision")

    with col2:
        status = st.selectbox("Status", [
            "Pending","Under Review","Approved",
            "Approved with Comments","Rejected","Submitted"
        ])
        recipient = st.text_input("Recipient", "Consultant")
        sender = st.text_input("Sender")
        lang = st.selectbox("Language", ["English","Arabic"])

    if st.button("Generate"):

        subject = subject_builder(p,t,no,title,status)
        body = email_body(status,title,no,rev,recipient,sender,lang)

        save_email(
            project_code=p,
            document_number=no,
            document_type=t,
            title=title,
            revision=rev,
            status=status,
            recipient=recipient,
            subject=subject,
            body=body,
            language=lang,
            source="single"
        )

        st.success("Generated")

        st.subheader("Subject")
        copy_btn(subject)

        st.subheader("Body")
        copy_btn(body)

# =========================================================
# BULK
# =========================================================
elif menu == "Bulk":

    st.header("Bulk Emails")

    sender = st.text_input("Sender Name")
    lang = st.selectbox("Language", ["English","Arabic"])

    file = st.file_uploader("Upload Excel", type="xlsx")

    if file:
        df = pd.read_excel(file)

        st.dataframe(df)

        if st.button("Generate Emails"):

            results = []

            for _, r in df.iterrows():

                subject = subject_builder(
                    r["Project Code"],
                    r["Document Type"],
                    r["Document Number"],
                    r["Title"],
                    r["Status"]
                )

                body = email_body(
                    r["Status"],
                    r["Title"],
                    r["Document Number"],
                    r["Revision"],
                    "Consultant",
                    sender,
                    lang
                )

                save_email(
                    project_code=r["Project Code"],
                    document_number=r["Document Number"],
                    document_type=r["Document Type"],
                    title=r["Title"],
                    revision=r["Revision"],
                    status=r["Status"],
                    recipient="Consultant",
                    subject=subject,
                    body=body,
                    language=lang,
                    source="bulk"
                )

                results.append({
                    "Doc": r["Document Number"],
                    "Status": r["Status"],
                    "Subject": subject
                })

            st.success("Done")

            st.dataframe(pd.DataFrame(results))

# =========================================================
# DASHBOARD
# =========================================================
elif menu == "Dashboard":

    st.header("Dashboard")

    df = load_emails()

    if not df.empty:
        st.metric("Total", len(df))

        st.bar_chart(df["status"].value_counts())

# =========================================================
# HISTORY
# =========================================================
elif menu == "History":

    st.header("Saved Emails")

    df = load_emails()

    st.dataframe(df)

    for _, r in df.head(10).iterrows():
        with st.expander(r["subject"]):
            st.code(r["body"])
