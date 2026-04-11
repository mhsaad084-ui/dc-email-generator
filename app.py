import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="DC Email Generator", layout="wide")

st.title("📧 Engineering Document Control Email Automation Tool")

# ================= HELPERS =================

def detect_email_type(status):
    if status in ["Pending", "Under Review"]:
        return "REMINDER"
    elif status == "Approved":
        return "APPROVED"
    elif status == "Approved with Comments":
        return "APPROVED WITH COMMENTS"
    elif status == "Rejected":
        return "REJECTED"
    elif status == "Submitted":
        return "SUBMISSION"
    else:
        return "GENERAL"

def build_subject(project_code, doc_type, doc_number, title, status):
    email_type = detect_email_type(status)

    prefix_map = {
        "REMINDER": "[REMINDER]",
        "APPROVED": "[APPROVED]",
        "APPROVED WITH COMMENTS": "[COMMENTS]",
        "REJECTED": "[REJECTED]",
        "SUBMISSION": "[SUBMIT]",
        "GENERAL": ""
    }

    prefix = prefix_map.get(email_type, "")

    return f"{prefix} {project_code} - {doc_type} - {doc_number} - {title} - {status}"

def build_email(recipient, doc_type, doc_number, title, revision, status, sender_name, language):

    if language == "English":

        if status in ["Pending", "Under Review"]:
            body = f"""Dear {recipient},

We are writing to follow up on the review status of the {doc_type.lower()} for {title} ({doc_number}, {revision}).

Kindly provide an update or the reviewed outcome for this submission at your earliest convenience to ensure the continuity of the project workflow.

Your prompt attention to this matter is highly appreciated.

Best regards,
{sender_name if sender_name else "[Sender Name]"}"""

        elif status == "Approved":
            body = f"""Dear {recipient},

We are pleased to inform you that the {doc_type.lower()} for {title} ({doc_number}, {revision}) has been approved.

Please proceed accordingly.

Best regards,
{sender_name}"""

        elif status == "Approved with Comments":
            body = f"""Dear {recipient},

Please be informed that the {doc_type.lower()} for {title} ({doc_number}, {revision}) has been approved with comments.

Kindly address all comments and proceed accordingly.

Best regards,
{sender_name}"""

        elif status == "Rejected":
            body = f"""Dear {recipient},

Please be informed that the {doc_type.lower()} for {title} ({doc_number}, {revision}) has been rejected.

Kindly revise and resubmit after addressing all comments.

Best regards,
{sender_name}"""

        else:
            body = f"""Dear {recipient},

Please find attached {doc_type} {doc_number} for your review.

Best regards,
{sender_name}"""

    else:  # Arabic

        if status in ["Pending", "Under Review"]:
            body = f"""السيد / {recipient}،

نود المتابعة بخصوص حالة مراجعة {doc_type} الخاص بـ {title} ({doc_number}، {revision}).

برجاء التكرم بتزويدنا بالتحديث في أقرب وقت ممكن لضمان استمرارية سير المشروع.

شاكرين تعاونكم.

مع خالص التحية،
{sender_name}"""

        elif status == "Approved":
            body = f"""السيد / {recipient}،

نحيطكم علمًا بأنه تم اعتماد {doc_type} الخاص بـ {title} ({doc_number}، {revision}).

يرجى المتابعة وفقًا لذلك.

مع خالص التحية،
{sender_name}"""

        elif status == "Rejected":
            body = f"""السيد / {recipient}،

نحيطكم علمًا بأنه تم رفض {doc_type} الخاص بـ {title} ({doc_number}، {revision}).

يرجى تعديل المستند وإعادة الإرسال بعد معالجة جميع الملاحظات.

مع خالص التحية،
{sender_name}"""

        else:
            body = f"""السيد / {recipient}،

يرجى مراجعة المستند المرفق.

مع خالص التحية،
{sender_name}"""

    return body

# ================= TABS =================
tab1, tab2 = st.tabs(["Single Email", "Bulk Emails"])

# ================= SINGLE =================
with tab1:

    col1, col2 = st.columns(2)

    with col1:
        project_code = st.text_input("Project Code")
        doc_number = st.text_input("Document Number")
        doc_type = st.selectbox("Document Type", ["Submittal", "RFI", "Drawing"])
        title = st.text_input("Title")
        revision = st.text_input("Revision")

    with col2:
        status = st.selectbox("Status", ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"])
        recipient = st.text_input("Recipient Name")
        sender_name = st.text_input("Sender Name")
        language = st.selectbox("Language", ["English", "Arabic"])

    if st.button("Generate Email"):

        if not project_code or not doc_number or not title:
            st.warning("Fill required fields")
        else:
            recipient_name = recipient if recipient else "Consultant"

            subject = build_subject(project_code, doc_type, doc_number, title, status)
            body = build_email(recipient_name, doc_type, doc_number, title, revision, status, sender_name, language)

            st.success("Email Generated")

            st.code(subject)
            st.code(body)

# ================= BULK =================
with tab2:

    file = st.file_uploader("Upload Excel", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        st.dataframe(df)

        status_filter = st.selectbox("Filter by Status", ["All"] + list(df["Status"].unique()))

        if status_filter != "All":
            df = df[df["Status"] == status_filter]

        if st.button("Generate Emails from Excel"):

            results = []

            for _, row in df.iterrows():

                subject = build_subject(
                    row["Project Code"],
                    row["Document Type"],
                    row["Document Number"],
                    row["Title"],
                    row["Status"]
                )

                body = build_email(
                    "Consultant",
                    row["Document Type"],
                    row["Document Number"],
                    row["Title"],
                    row["Revision"],
                    row["Status"],
                    sender_name,
                    "English"
                )

                results.append(subject + "\n\n" + body)

            st.success("Emails Generated")

            for r in results:
                st.code(r)

            csv = pd.DataFrame(results, columns=["Emails"]).to_csv(index=False)

            st.download_button("Download CSV", csv, "emails.csv")
