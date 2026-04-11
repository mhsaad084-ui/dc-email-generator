import html
import sqlite3
from io import StringIO
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Engineering Document Control Email System",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# DATABASE
# =========================================================
DB_NAME = "emails.db"


def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False)


conn = get_connection()


def init_db():
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS saved_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            project_code TEXT,
            document_number TEXT,
            document_type TEXT,
            discipline TEXT,
            title TEXT,
            revision TEXT,
            status TEXT,
            action_required TEXT,
            submitted_by TEXT,
            submitted_to TEXT,
            sent_date TEXT,
            received_date TEXT,
            recipient TEXT,
            sender_name TEXT,
            company_name TEXT,
            language TEXT,
            email_type TEXT,
            subject TEXT,
            body TEXT,
            source_type TEXT
        )
        """
    )
    conn.commit()


init_db()


def save_email_to_db(
    project_code,
    document_number,
    document_type,
    discipline,
    title,
    revision,
    status,
    action_required,
    submitted_by,
    submitted_to,
    sent_date,
    received_date,
    recipient,
    sender_name,
    company_name,
    language,
    email_type,
    subject,
    body,
    source_type="single",
):
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO saved_emails (
            created_at, project_code, document_number, document_type, discipline,
            title, revision, status, action_required, submitted_by, submitted_to,
            sent_date, received_date, recipient, sender_name, company_name,
            language, email_type, subject, body, source_type
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            project_code,
            document_number,
            document_type,
            discipline,
            title,
            revision,
            status,
            action_required,
            submitted_by,
            submitted_to,
            sent_date,
            received_date,
            recipient,
            sender_name,
            company_name,
            language,
            email_type,
            subject,
            body,
            source_type,
        ),
    )
    conn.commit()


def load_saved_emails():
    return pd.read_sql_query(
        """
        SELECT *
        FROM saved_emails
        ORDER BY id DESC
        """,
        conn,
    )


def delete_all_saved_emails():
    cursor = conn.cursor()
    cursor.execute("DELETE FROM saved_emails")
    conn.commit()


# =========================================================
# UTILS
# =========================================================
def norm(value, default=""):
    if pd.isna(value):
        return default
    return str(value).strip()


def copy_button(text: str, label: str, key: str):
    escaped_text = html.escape(text)
    button_id = f"btn_{key}"
    components.html(
        f"""
        <div style="margin-bottom: 8px;">
            <button id="{button_id}"
                style="
                    background:#2563eb;
                    color:white;
                    border:none;
                    padding:8px 14px;
                    border-radius:8px;
                    cursor:pointer;
                    font-size:14px;
                    font-weight:600;
                ">
                {html.escape(label)}
            </button>
            <span id="{button_id}_msg" style="margin-left:10px;color:#16a34a;font-size:13px;"></span>
        </div>
        <script>
            const btn = document.getElementById("{button_id}");
            const msg = document.getElementById("{button_id}_msg");
            btn.onclick = async function() {{
                try {{
                    await navigator.clipboard.writeText(`{escaped_text}`);
                    msg.innerText = "Copied";
                    setTimeout(() => msg.innerText = "", 1500);
                }} catch (e) {{
                    msg.innerText = "Copy failed";
                }}
            }};
        </script>
        """,
        height=45,
    )


def render_email_card(subject: str, recipient: str, body: str):
    st.markdown(
        f"""
        <div style="
            background:#ffffff;
            padding:22px;
            border-radius:14px;
            border:1px solid #e5e7eb;
            box-shadow:0 4px 14px rgba(0,0,0,0.08);
            color:#111827;
            margin-bottom:18px;
        ">
            <div style="font-size:18px; font-weight:700; margin-bottom:8px;">
                {html.escape(subject)}
            </div>
            <div style="font-size:13px; color:#6b7280; margin-bottom:18px;">
                To: {html.escape(recipient)}
            </div>
            <div style="font-size:15px; line-height:1.8; white-space:pre-wrap;">
                {html.escape(body)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# DECISION ENGINE
# =========================================================
def get_default_recipient(status: str, submitted_by: str = "", submitted_to: str = "") -> str:
    submitted_by_l = submitted_by.lower()
    submitted_to_l = submitted_to.lower()

    internal_keywords = [
        "team",
        "department",
        "technical office",
        "project control",
        "internal",
        "engineering",
    ]

    if any(k in submitted_by_l for k in internal_keywords) and any(
        k in submitted_to_l for k in internal_keywords
    ):
        return "Team"

    if status in ["Under Review", "Pending"]:
        return "Consultant"
    if status in ["Approved with Comments", "Rejected", "Approved"]:
        return "Contractor"
    if status == "Submitted":
        return "Client"
    return "Team"


def detect_email_type(status: str, action_required: str) -> str:
    if status in ["Under Review", "Pending"] or action_required in ["Follow-up", "Urgent Follow-up"]:
        return "FOLLOW_UP"

    if status in ["Approved with Comments", "Rejected"] and action_required == "Resubmit":
        return "RESUBMIT_REQUEST"

    if action_required == "Missing Attachment Correction":
        return "MISSING_ATTACHMENT"

    if status == "Submitted":
        return "SUBMISSION"

    if status == "Approved":
        return "APPROVED_NOTICE"

    return "GENERAL"


def get_priority_tag(status: str, action_required: str, urgent: bool = False) -> str:
    if urgent or action_required == "Urgent Follow-up":
        return "[URGENT]"
    if status in ["Pending", "Under Review"]:
        return "[REMINDER]"
    if status == "Approved":
        return "[APPROVED]"
    if status == "Approved with Comments":
        return "[COMMENTS]"
    if status == "Rejected":
        return "[REJECTED]"
    if status == "Submitted":
        return "[SUBMISSION]"
    return ""


def get_status_action_text(status: str, action_required: str) -> str:
    if action_required and action_required not in ["None", "No Action", ""]:
        return action_required
    return status


def build_subject(
    project_code: str,
    document_type: str,
    document_number: str,
    title: str,
    status: str,
    action_required: str,
    urgent: bool = False,
) -> str:
    prefix = get_priority_tag(status, action_required, urgent)
    status_or_action = get_status_action_text(status, action_required)
    base = f"{project_code} - {document_type} - {document_number} - {title} - {status_or_action}"
    return f"{prefix} {base}".strip()


# =========================================================
# TEMPLATES
# =========================================================
def build_follow_up_body_en(
    recipient, document_type, document_number, title, revision, status, sent_date, sender_name, company_name, urgent
):
    if document_type == "RFI":
        intro = f'We are writing to follow up on RFI No. {document_number} regarding "{title}" ({revision}).'
        action = "Kindly provide your response or clarification at your earliest convenience to ensure the continuity of the project workflow."
    elif document_type in ["Submittal", "Drawing"]:
        intro = f'We are writing to follow up on the review status of the {document_type.lower()} for "{title}" ({document_number}, {revision}).'
        action = "Kindly provide an update or the reviewed outcome for this submission at your earliest convenience to ensure the continuity of the project workflow."
    elif document_type == "Report":
        intro = f'We are following up on Report No. {document_number} regarding "{title}" ({revision}).'
        action = "Kindly provide your update at your earliest convenience."
    else:
        intro = f'We are writing to follow up on {document_type} No. {document_number} regarding "{title}" ({revision}).'
        action = "Kindly provide an update at your earliest convenience."

    status_line = ""
    if sent_date:
        status_line = f'\n\nThe subject document was submitted on {sent_date} and is currently "{status}".'

    urgent_line = ""
    if urgent:
        urgent_line = "\n\nThis matter is considered urgent and requires your immediate attention."

    return f"""Dear {recipient},

{intro}{status_line}

{action}

Your prompt attention to this matter is highly appreciated.{urgent_line}

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_resubmit_body_en(
    recipient, document_type, document_number, title, revision, status, sender_name, company_name
):
    if status == "Approved with Comments":
        line = f'Please be informed that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been reviewed and marked as "Approved with Comments".'
        request = "Kindly address all comments and resubmit the revised document for further review and approval."
    else:
        line = f'Please be informed that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been reviewed and marked as "Rejected".'
        request = "Kindly revise and resubmit the document after addressing all comments."

    return f"""Dear {recipient},

{line}

{request}

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_submission_body_en(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    if document_type == "Report":
        main = f'Please find attached {document_type} No. {document_number} regarding "{title}" ({revision}) submitted for your review and record.'
    else:
        main = f'Please find attached {document_type} No. {document_number} regarding "{title}" ({revision}), submitted for your review.'

    if document_type == "Report":
        closing_request = ""
    else:
        closing_request = "\n\nKindly review and provide your comments or approval at your earliest convenience."

    return f"""Dear {recipient},

{main}{closing_request}

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_approved_notice_body_en(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""Dear {recipient},

We are pleased to inform you that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been approved.

Please proceed accordingly.

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_missing_attachment_body_en(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""Dear {recipient},

Please accept our apologies.

Please find the correct attachment for {document_type} No. {document_number} regarding "{title}" ({revision}).

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_general_body_en(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""Dear {recipient},

Please proceed regarding {document_type} No. {document_number} concerning "{title}" ({revision}).

Best regards,
{sender_name}
Document Control Department
{company_name}"""


def build_follow_up_body_ar(
    recipient, document_type, document_number, title, revision, status, sent_date, sender_name, company_name, urgent
):
    if document_type == "RFI":
        intro = f'نود المتابعة بخصوص RFI رقم {document_number} الخاص بـ "{title}" ({revision}).'
        action = "نرجو التكرم بتزويدنا بالرد أو الإيضاح في أقرب وقت ممكن لضمان استمرارية سير العمل بالمشروع."
    elif document_type in ["Submittal", "Drawing"]:
        intro = f'نود المتابعة بخصوص حالة مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}).'
        action = "نرجو التكرم بتزويدنا بالتحديث أو نتيجة المراجعة في أقرب وقت ممكن لضمان استمرارية سير العمل بالمشروع."
    elif document_type == "Report":
        intro = f'نود المتابعة بخصوص Report رقم {document_number} الخاص بـ "{title}" ({revision}).'
        action = "نرجو التكرم بتزويدنا بالتحديث في أقرب وقت ممكن."
    else:
        intro = f'نود المتابعة بخصوص {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}).'
        action = "نرجو التكرم بتزويدنا بالتحديث في أقرب وقت ممكن."

    status_line = ""
    if sent_date:
        status_line = f'\n\nتم إرسال المستند بتاريخ {sent_date} ولا يزال بحالة "{status}".'

    urgent_line = ""
    if urgent:
        urgent_line = "\n\nيرجى العلم بأن هذا الموضوع عاجل ويتطلب اهتمامكم الفوري."

    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{intro}{status_line}

{action}

شاكرين حسن تعاونكم.{urgent_line}

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_resubmit_body_ar(
    recipient, document_type, document_number, title, revision, status, sender_name, company_name
):
    if status == "Approved with Comments":
        line = f'نحيطكم علمًا بأنه تمت مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}) وتم اعتماده مع ملاحظات.'
        request = "يرجى مراجعة جميع الملاحظات وإعادة تقديم المستند بعد التعديل لاستكمال المراجعة والاعتماد."
    else:
        line = f'نحيطكم علمًا بأنه تمت مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}) وتم رفضه.'
        request = "يرجى تعديل المستند وإعادة التقديم بعد معالجة جميع الملاحظات."

    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{line}

{request}

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_submission_body_ar(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    if document_type == "Report":
        main = f'نرفق لسيادتكم {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) للمراجعة والحفظ.'
    else:
        main = f'نرفق لسيادتكم {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) للمراجعة.'

    if document_type == "Report":
        closing_request = ""
    else:
        closing_request = "\n\nنرجو التكرم بالمراجعة وإفادتنا بالملاحظات أو الاعتماد في أقرب وقت ممكن."

    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{main}{closing_request}

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_approved_notice_body_ar(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأنه تم اعتماد {document_type} الخاص بـ "{title}" ({document_number}، {revision}).

يرجى المتابعة واتخاذ اللازم وفقًا لذلك.

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_missing_attachment_body_ar(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نعتذر عن السهو السابق.

نرفق لسيادتكم المرفق الصحيح الخاص بـ {document_type} رقم {document_number} بعنوان "{title}" ({revision}).

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_general_body_ar(
    recipient, document_type, document_number, title, revision, sender_name, company_name
):
    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

يرجى التكرم بمراجعة {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) واتخاذ اللازم.

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def build_email_body(
    recipient,
    project_code,
    document_number,
    document_type,
    discipline,
    title,
    revision,
    status,
    action_required,
    sender_name,
    company_name,
    language,
    sent_date="",
    email_type="GENERAL",
    urgent=False,
):
    sender_name = sender_name if sender_name else "[Sender Name]"
    company_name = company_name if company_name else "[Company Name]"

    if language == "English":
        if email_type == "FOLLOW_UP":
            return build_follow_up_body_en(
                recipient, document_type, document_number, title, revision, status, sent_date, sender_name, company_name, urgent
            )
        if email_type == "RESUBMIT_REQUEST":
            return build_resubmit_body_en(
                recipient, document_type, document_number, title, revision, status, sender_name, company_name
            )
        if email_type == "SUBMISSION":
            return build_submission_body_en(
                recipient, document_type, document_number, title, revision, sender_name, company_name
            )
        if email_type == "APPROVED_NOTICE":
            return build_approved_notice_body_en(
                recipient, document_type, document_number, title, revision, sender_name, company_name
            )
        if email_type == "MISSING_ATTACHMENT":
            return build_missing_attachment_body_en(
                recipient, document_type, document_number, title, revision, sender_name, company_name
            )
        return build_general_body_en(
            recipient, document_type, document_number, title, revision, sender_name, company_name
        )

    if email_type == "FOLLOW_UP":
        return build_follow_up_body_ar(
            recipient, document_type, document_number, title, revision, status, sent_date, sender_name, company_name, urgent
        )
    if email_type == "RESUBMIT_REQUEST":
        return build_resubmit_body_ar(
            recipient, document_type, document_number, title, revision, status, sender_name, company_name
        )
    if email_type == "SUBMISSION":
        return build_submission_body_ar(
            recipient, document_type, document_number, title, revision, sender_name, company_name
        )
    if email_type == "APPROVED_NOTICE":
        return build_approved_notice_body_ar(
            recipient, document_type, document_number, title, revision, sender_name, company_name
        )
    if email_type == "MISSING_ATTACHMENT":
        return build_missing_attachment_body_ar(
            recipient, document_type, document_number, title, revision, sender_name, company_name
        )
    return build_general_body_ar(
        recipient, document_type, document_number, title, revision, sender_name, company_name
    )


# =========================================================
# APP UI
# =========================================================
st.title("📧 Engineering Document Control Email System")

page = st.sidebar.radio(
    "System Navigation",
    ["Single Email", "Bulk Emails", "Dashboard", "Saved Emails"]
)

# =========================================================
# SINGLE EMAIL
# =========================================================
if page == "Single Email":
    st.header("Single Email Generator")

    c1, c2 = st.columns(2)

    with c1:
        project_code = st.text_input("Project Code")
        document_number = st.text_input("Document Number")
        document_type = st.selectbox("Document Type", ["Submittal", "RFI", "Drawing", "Report"])
        discipline = st.text_input("Discipline")
        title = st.text_input("Title")
        revision = st.text_input("Revision", placeholder="Rev.00")
        sent_date = st.text_input("Sent Date", placeholder="April 16, 2026")

    with c2:
        status = st.selectbox(
            "Status",
            ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"]
        )
        action_required = st.selectbox(
            "Action Required",
            ["Follow-up", "Urgent Follow-up", "Resubmit", "No Action", "Missing Attachment Correction"]
        )
        submitted_by = st.text_input("Submitted By")
        submitted_to = st.text_input("Submitted To")
        recipient_name = st.text_input("Recipient Name")
        sender_name = st.text_input("Sender Name")
        company_name = st.text_input("Company Name")
        language = st.selectbox("Language", ["English", "Arabic"])

    received_date = st.text_input("Received Date", placeholder="Optional")
    urgent = st.checkbox("Mark as Urgent")

    if project_code and document_number and title:
        st.info(f"Preview Subject: {project_code} - {document_type} - {document_number} - {title}")

    status_colors = {
        "Pending": "orange",
        "Under Review": "blue",
        "Approved": "green",
        "Approved with Comments": "purple",
        "Rejected": "red",
        "Submitted": "teal",
    }

    st.markdown(
        f"<span style='color:{status_colors.get(status, 'gray')}; font-weight:bold;'>Status: {status}</span>",
        unsafe_allow_html=True,
    )

    btn1, btn2 = st.columns(2)
    generate_clicked = btn1.button("Generate Email", type="primary")
    reset_clicked = btn2.button("Reset Form")

    if reset_clicked:
        st.rerun()

    if generate_clicked:
        if not project_code:
            st.warning("Project Code is required")
        elif not document_number:
            st.warning("Document Number is required")
        elif not title:
            st.warning("Title is required")
        elif not revision:
            st.warning("Revision is required")
        elif not sender_name:
            st.warning("Sender Name is required")
        else:
            email_type = detect_email_type(status, action_required)
            default_recipient = get_default_recipient(status, submitted_by, submitted_to)
            recipient = recipient_name if recipient_name else default_recipient

            subject = build_subject(
                project_code=project_code,
                document_type=document_type,
                document_number=document_number,
                title=title,
                status=status,
                action_required=action_required,
                urgent=urgent,
            )

            body = build_email_body(
                recipient=recipient,
                project_code=project_code,
                document_number=document_number,
                document_type=document_type,
                discipline=discipline,
                title=title,
                revision=revision,
                status=status,
                action_required=action_required,
                sender_name=sender_name,
                company_name=company_name,
                language=language,
                sent_date=sent_date,
                email_type=email_type,
                urgent=urgent,
            )

            save_email_to_db(
                project_code=project_code,
                document_number=document_number,
                document_type=document_type,
                discipline=discipline,
                title=title,
                revision=revision,
                status=status,
                action_required=action_required,
                submitted_by=submitted_by,
                submitted_to=submitted_to,
                sent_date=sent_date,
                received_date=received_date,
                recipient=recipient,
                sender_name=sender_name,
                company_name=company_name,
                language=language,
                email_type=email_type,
                subject=subject,
                body=body,
                source_type="single",
            )

            st.success("Email generated and saved successfully")
            render_email_card(subject, recipient, body)

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### Subject")
                st.code(subject, language="text")
                copy_button(subject, "Copy Subject", "single_subject")
            with col2:
                st.markdown("### Body")
                st.code(body, language="text")
                copy_button(body, "Copy Body", "single_body")

# =========================================================
# BULK EMAILS
# =========================================================
elif page == "Bulk Emails":
    st.header("Bulk Email Generator (Excel)")

    sender_name_bulk = st.text_input("Sender Name for Bulk Emails")
    company_name_bulk = st.text_input("Company Name for Bulk Emails")
    language_bulk = st.selectbox("Language for Bulk Emails", ["English", "Arabic"])

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.stop()

        required_columns = [
            "Project Code",
            "Document Number",
            "Document Type",
            "Title",
            "Revision",
            "Status",
            "Action Required",
        ]

        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            st.error(f"Missing required columns: {', '.join(missing_cols)}")
            st.stop()

        optional_defaults = {
            "Discipline": "",
            "Submitted By": "",
            "Submitted To": "",
            "Sent Date": "",
            "Received Date": "",
        }

        for col, default_val in optional_defaults.items():
            if col not in df.columns:
                df[col] = default_val

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        status_filter = st.selectbox("Filter by Status", ["All"] + list(df["Status"].dropna().unique()))
        if status_filter != "All":
            df = df[df["Status"] == status_filter]

        if st.button("Generate Emails from Excel", type="primary"):
            results = []

            for _, row in df.iterrows():
                row_project_code = norm(row["Project Code"])
                row_document_number = norm(row["Document Number"])
                row_document_type = norm(row["Document Type"])
                row_title = norm(row["Title"])
                row_revision = norm(row["Revision"])
                row_status = norm(row["Status"])
                row_action_required = norm(row["Action Required"])
                row_discipline = norm(row["Discipline"])
                row_submitted_by = norm(row["Submitted By"])
                row_submitted_to = norm(row["Submitted To"])
                row_sent_date = norm(row["Sent Date"])
                row_received_date = norm(row["Received Date"])

                row_email_type = detect_email_type(row_status, row_action_required)
                row_recipient = get_default_recipient(row_status, row_submitted_by, row_submitted_to)

                row_subject = build_subject(
                    project_code=row_project_code,
                    document_type=row_document_type,
                    document_number=row_document_number,
                    title=row_title,
                    status=row_status,
                    action_required=row_action_required,
                    urgent=False,
                )

                row_body = build_email_body(
                    recipient=row_recipient,
                    project_code=row_project_code,
                    document_number=row_document_number,
                    document_type=row_document_type,
                    discipline=row_discipline,
                    title=row_title,
                    revision=row_revision,
                    status=row_status,
                    action_required=row_action_required,
                    sender_name=sender_name_bulk if sender_name_bulk else "[Sender Name]",
                    company_name=company_name_bulk if company_name_bulk else "[Company Name]",
                    language=language_bulk,
                    sent_date=row_sent_date,
                    email_type=row_email_type,
                    urgent=False,
                )

                save_email_to_db(
                    project_code=row_project_code,
                    document_number=row_document_number,
                    document_type=row_document_type,
                    discipline=row_discipline,
                    title=row_title,
                    revision=row_revision,
                    status=row_status,
                    action_required=row_action_required,
                    submitted_by=row_submitted_by,
                    submitted_to=row_submitted_to,
                    sent_date=row_sent_date,
                    received_date=row_received_date,
                    recipient=row_recipient,
                    sender_name=sender_name_bulk if sender_name_bulk else "[Sender Name]",
                    company_name=company_name_bulk if company_name_bulk else "[Company Name]",
                    language=language_bulk,
                    email_type=row_email_type,
                    subject=row_subject,
                    body=row_body,
                    source_type="bulk",
                )

                results.append(
                    {
                        "Document Number": row_document_number,
                        "Status": row_status,
                        "Recipient": row_recipient,
                        "Subject": row_subject,
                        "Body": row_body,
                    }
                )

            result_df = pd.DataFrame(results)
            st.success("Bulk emails generated and saved successfully")
            st.dataframe(result_df[["Document Number", "Status", "Recipient", "Subject"]], use_container_width=True)

            st.subheader("Email Preview")
            for idx, item in enumerate(results):
                render_email_card(item["Subject"], item["Recipient"], item["Body"])

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Subject**")
                    st.code(item["Subject"], language="text")
                    copy_button(item["Subject"], "Copy Subject", f"bulk_subject_{idx}")
                with c2:
                    st.markdown("**Body**")
                    st.code(item["Body"], language="text")
                    copy_button(item["Body"], "Copy Body", f"bulk_body_{idx}")

            csv_buffer = StringIO()
            result_df.to_csv(csv_buffer, index=False)

            st.download_button(
                label="📥 Download Emails as CSV",
                data=csv_buffer.getvalue(),
                file_name="generated_emails.csv",
                mime="text/csv",
            )

# =========================================================
# DASHBOARD
# =========================================================
elif page == "Dashboard":
    st.header("Dashboard Tracking")

    saved_df = load_saved_emails()

    total_emails = len(saved_df)
    pending_count = len(saved_df[saved_df["status"] == "Pending"])
    review_count = len(saved_df[saved_df["status"] == "Under Review"])
    approved_count = len(saved_df[saved_df["status"] == "Approved"])
    comments_count = len(saved_df[saved_df["status"] == "Approved with Comments"])
    rejected_count = len(saved_df[saved_df["status"] == "Rejected"])
    submitted_count = len(saved_df[saved_df["status"] == "Submitted"])

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Total", total_emails)
    c2.metric("Pending", pending_count)
    c3.metric("Under Review", review_count)
    c4.metric("Approved", approved_count)
    c5.metric("Comments", comments_count)
    c6.metric("Rejected", rejected_count)
    c7.metric("Submitted", submitted_count)

    st.subheader("Status Distribution")
    if not saved_df.empty:
        st.bar_chart(saved_df["status"].value_counts())
    else:
        st.info("No saved emails yet.")

    st.subheader("Recent Activity")
    if not saved_df.empty:
        st.dataframe(
            saved_df[
                [
                    "created_at",
                    "project_code",
                    "document_number",
                    "document_type",
                    "title",
                    "status",
                    "recipient",
                    "language",
                    "source_type",
                ]
            ].head(20),
            use_container_width=True,
        )

# =========================================================
# SAVED EMAILS
# =========================================================
elif page == "Saved Emails":
    st.header("Saved Emails / History")

    saved_df = load_saved_emails()

    if saved_df.empty:
        st.info("No saved emails found.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Filter by Status", ["All"] + sorted(saved_df["status"].dropna().unique().tolist()))
        with col2:
            language_filter = st.selectbox("Filter by Language", ["All"] + sorted(saved_df["language"].dropna().unique().tolist()))

        filtered_df = saved_df.copy()

        if status_filter != "All":
            filtered_df = filtered_df[filtered_df["status"] == status_filter]

        if language_filter != "All":
            filtered_df = filtered_df[filtered_df["language"] == language_filter]

        st.dataframe(
            filtered_df[
                [
                    "created_at",
                    "project_code",
                    "document_number",
                    "document_type",
                    "title",
                    "status",
                    "recipient",
                    "language",
                    "source_type",
                ]
            ],
            use_container_width=True,
        )

        st.subheader("Open Saved Email")
        for idx, row in filtered_df.head(20).iterrows():
            with st.expander(f"{row['document_number']} | {row['title']} | {row['status']}"):
                render_email_card(row["subject"], row["recipient"], row["body"])

                a1, a2 = st.columns(2)
                with a1:
                    st.markdown("**Subject**")
                    st.code(row["subject"], language="text")
                    copy_button(row["subject"], "Copy Subject", f"saved_subject_{idx}")
                with a2:
                    st.markdown("**Body**")
                    st.code(row["body"], language="text")
                    copy_button(row["body"], "Copy Body", f"saved_body_{idx}")

        st.divider()
        if st.button("Delete All Saved Emails"):
            delete_all_saved_emails()
            st.success("All saved emails deleted")
            st.rerun()
