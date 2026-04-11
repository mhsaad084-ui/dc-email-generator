import html
import io
import sqlite3
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
            subject TEXT,
            body TEXT,
            language TEXT,
            source TEXT
        )
        """
    )
    conn.commit()


init_db()


def save_email(**kwargs):
    cols = [
        "created_at",
        "project_code",
        "document_number",
        "document_type",
        "title",
        "revision",
        "status",
        "recipient",
        "subject",
        "body",
        "language",
        "source",
    ]

    values = []
    for c in cols:
        if c == "created_at":
            values.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            values.append(kwargs.get(c, ""))  # safe fallback

    placeholders = ",".join(["?"] * len(cols))
    conn.execute(
        f"INSERT INTO emails ({','.join(cols)}) VALUES ({placeholders})",
        values,
    )
    conn.commit()


def load_emails():
    return pd.read_sql("SELECT * FROM emails ORDER BY id DESC", conn)


# =========================================================
# HELPERS
# =========================================================
def norm(value, default=""):
    if pd.isna(value):
        return default
    return str(value).strip()


def subject_builder(project_code, doc_type, doc_no, title, status):
    prefix_map = {
        "Pending": "[REMINDER]",
        "Under Review": "[REMINDER]",
        "Approved": "[APPROVED]",
        "Approved with Comments": "[COMMENTS]",
        "Rejected": "[REJECTED]",
        "Submitted": "[SUBMIT]",
    }
    prefix = prefix_map.get(status, "[NOTICE]")
    return f"{prefix} {project_code} - {doc_type} - {doc_no} - {title} - {status}"


def get_default_recipient(status):
    if status in ["Pending", "Under Review"]:
        return "Consultant"
    if status in ["Approved", "Approved with Comments", "Rejected"]:
        return "Contractor"
    if status == "Submitted":
        return "Client"
    return "Team"


def email_body(status, title, doc_no, rev, recipient, sender, lang, doc_type):
    sender = sender if sender else "[Sender Name]"

    if lang == "English":
        if status in ["Pending", "Under Review"]:
            return f"""Dear {recipient},

We are writing to follow up on the review status of the {doc_type.lower()} for "{title}" ({doc_no}, {rev}).

Kindly provide your update or the reviewed outcome at your earliest convenience to ensure the continuity of the project workflow.

Your prompt attention to this matter is highly appreciated.

Best regards,
{sender}"""

        if status == "Approved":
            return f"""Dear {recipient},

We are pleased to inform you that the {doc_type.lower()} for "{title}" ({doc_no}, {rev}) has been approved.

Please proceed accordingly.

Best regards,
{sender}"""

        if status == "Approved with Comments":
            return f"""Dear {recipient},

Please be informed that the {doc_type.lower()} for "{title}" ({doc_no}, {rev}) has been approved with comments.

Kindly address all comments and proceed accordingly.

Best regards,
{sender}"""

        if status == "Rejected":
            return f"""Dear {recipient},

Please be informed that the {doc_type.lower()} for "{title}" ({doc_no}, {rev}) has been rejected.

Kindly revise and resubmit after addressing all comments.

Best regards,
{sender}"""

        if status == "Submitted":
            return f"""Dear {recipient},

Please find attached {doc_type} No. {doc_no} regarding "{title}" ({rev}), submitted for your review.

Kindly review and provide your comments or approval at your earliest convenience.

Best regards,
{sender}"""

        return f"""Dear {recipient},

Please proceed regarding {doc_type} No. {doc_no}.

Best regards,
{sender}"""

    # Arabic
    if status in ["Pending", "Under Review"]:
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نود المتابعة بخصوص حالة مراجعة {doc_type} الخاص بـ "{title}" ({doc_no}، {rev}).

نرجو التكرم بتزويدنا بالتحديث أو نتيجة المراجعة في أقرب وقت ممكن لضمان استمرارية سير العمل بالمشروع.

شاكرين حسن تعاونكم.

وتفضلوا بقبول فائق الاحترام،
{sender}"""

    if status == "Approved":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأنه تم اعتماد {doc_type} الخاص بـ "{title}" ({doc_no}، {rev}).

يرجى المتابعة واتخاذ اللازم وفقًا لذلك.

وتفضلوا بقبول فائق الاحترام،
{sender}"""

    if status == "Approved with Comments":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأنه تمت مراجعة {doc_type} الخاص بـ "{title}" ({doc_no}، {rev}) وتم اعتماده مع ملاحظات.

يرجى مراجعة الملاحظات واستكمال اللازم وفقًا لها.

وتفضلوا بقبول فائق الاحترام،
{sender}"""

    if status == "Rejected":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأنه تم رفض {doc_type} الخاص بـ "{title}" ({doc_no}، {rev}).

يرجى تعديل المستند وإعادة الإرسال بعد معالجة جميع الملاحظات.

وتفضلوا بقبول فائق الاحترام،
{sender}"""

    if status == "Submitted":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نرفق لسيادتكم {doc_type} رقم {doc_no} الخاص بـ "{title}" ({rev}) للمراجعة.

نرجو التكرم بالمراجعة وإفادتنا بالملاحظات أو الاعتماد في أقرب وقت ممكن.

وتفضلوا بقبول فائق الاحترام،
{sender}"""

    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

يرجى مراجعة المستند واتخاذ اللازم.

وتفضلوا بقبول فائق الاحترام،
{sender}"""


def render_email_card(subject, body):
    st.markdown(
        f"""
        <div style="
            background:#1e1e1e;
            padding:20px;
            border-radius:10px;
            margin-bottom:12px;
            border-left:5px solid #00c853;
        ">
            <div style="font-weight:700; font-size:16px; margin-bottom:10px;">
                Subject: {html.escape(subject)}
            </div>
            <pre style="
                white-space: pre-wrap;
                font-family: inherit;
                font-size: 14px;
                line-height: 1.7;
                margin: 0;
            ">{html.escape(body)}</pre>
        </div>
        """,
        unsafe_allow_html=True,
    )


def copy_button(text: str, label: str, key: str):
    escaped = html.escape(text)
    button_id = f"copy_{key}"

    components.html(
        f"""
        <div style="margin-top:6px; margin-bottom:10px;">
            <button id="{button_id}"
                style="
                    background:#2563eb;
                    color:white;
                    border:none;
                    padding:8px 14px;
                    border-radius:8px;
                    cursor:pointer;
                    font-size:13px;
                    font-weight:600;">
                {html.escape(label)}
            </button>
            <span id="{button_id}_msg" style="margin-left:8px;color:#22c55e;font-size:12px;"></span>
        </div>
        <script>
            const btn = document.getElementById("{button_id}");
            const msg = document.getElementById("{button_id}_msg");
            btn.onclick = async function() {{
                try {{
                    await navigator.clipboard.writeText(`{escaped}`);
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


def convert_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Emails")
    return output.getvalue()


# =========================================================
# UI
# =========================================================
st.title("📧 Engineering Document Control Email System")

menu = st.sidebar.radio("Menu", ["Single", "Bulk", "Dashboard", "History"])

# =========================================================
# SINGLE
# =========================================================
if menu == "Single":
    st.header("Single Email")

    col1, col2 = st.columns(2)

    with col1:
        project_code = st.text_input("Project Code")
        doc_no = st.text_input("Document No")
        doc_type = st.selectbox("Type", ["Submittal", "RFI", "Drawing", "Report"])
        title = st.text_input("Title")
        rev = st.text_input("Revision", placeholder="Rev.00")

    with col2:
        status = st.selectbox(
            "Status",
            ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )
        recipient = st.text_input("Recipient")
        sender = st.text_input("Sender")
        lang = st.selectbox("Language", ["English", "Arabic"])

    if project_code and doc_no and title:
        st.info(f"Preview Subject: {project_code} - {doc_type} - {doc_no} - {title}")

    if st.button("Generate", type="primary"):
        if not project_code:
            st.warning("Project Code is required")
        elif not doc_no:
            st.warning("Document No is required")
        elif not title:
            st.warning("Title is required")
        elif not rev:
            st.warning("Revision is required")
        elif not sender:
            st.warning("Sender is required")
        else:
            recipient_final = recipient if recipient else get_default_recipient(status)
            subject = subject_builder(project_code, doc_type, doc_no, title, status)
            body = email_body(status, title, doc_no, rev, recipient_final, sender, lang, doc_type)

            save_email(
                project_code=project_code,
                document_number=doc_no,
                document_type=doc_type,
                title=title,
                revision=rev,
                status=status,
                recipient=recipient_final,
                subject=subject,
                body=body,
                language=lang,
                source="single",
            )

            st.success("Generated")
            render_email_card(subject, body)

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Subject")
                st.code(subject, language="text")
                copy_button(subject, "📋 Copy Subject", "single_subject")
            with c2:
                st.subheader("Body")
                st.code(body, language="text")
                copy_button(body, "📋 Copy Body", "single_body")

# =========================================================
# BULK
# =========================================================
elif menu == "Bulk":
    st.header("Bulk Emails")

    sender = st.text_input("Sender Name")
    lang = st.selectbox("Language", ["English", "Arabic"])

    file = st.file_uploader("Upload Excel", type="xlsx")

    if file:
        df = pd.read_excel(file)
        st.dataframe(df, use_container_width=True)

        status_filter = st.selectbox(
            "Filter by Status",
            ["All", "Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )

        if st.button("Generate Emails", type="primary"):
            results = []

            for _, r in df.iterrows():
                row_status = norm(r.get("Status", ""))

                if status_filter != "All" and row_status != status_filter:
                    continue

                row_project = norm(r.get("Project Code", ""))
                row_doc_type = norm(r.get("Document Type", ""))
                row_doc_no = norm(r.get("Document Number", ""))
                row_title = norm(r.get("Title", ""))
                row_rev = norm(r.get("Revision", ""))

                recipient_final = get_default_recipient(row_status)
                subject = subject_builder(row_project, row_doc_type, row_doc_no, row_title, row_status)
                body = email_body(row_status, row_title, row_doc_no, row_rev, recipient_final, sender, lang, row_doc_type)

                save_email(
                    project_code=row_project,
                    document_number=row_doc_no,
                    document_type=row_doc_type,
                    title=row_title,
                    revision=row_rev,
                    status=row_status,
                    recipient=recipient_final,
                    subject=subject,
                    body=body,
                    language=lang,
                    source="bulk",
                )

                results.append(
                    {
                        "Doc": row_doc_no,
                        "Status": row_status,
                        "Recipient": recipient_final,
                        "Subject": subject,
                        "Body": body,
                    }
                )

            st.success("Done")

            result_df = pd.DataFrame(results)
            if not result_df.empty:
                st.dataframe(result_df[["Doc", "Status", "Recipient", "Subject"]], use_container_width=True)

                st.subheader("Preview")
                for i, item in enumerate(results):
                    render_email_card(item["Subject"], item["Body"])

                    a, b = st.columns(2)
                    with a:
                        st.markdown("**Subject**")
                        st.code(item["Subject"], language="text")
                        copy_button(item["Subject"], "📋 Copy Subject", f"bulk_subject_{i}")
                    with b:
                        st.markdown("**Body**")
                        st.code(item["Body"], language="text")
                        copy_button(item["Body"], "📋 Copy Body", f"bulk_body_{i}")

                csv_data = result_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Download CSV",
                    data=csv_data,
                    file_name="emails.csv",
                    mime="text/csv",
                )

                excel_data = convert_excel(result_df)
                st.download_button(
                    "⬇️ Download Excel",
                    data=excel_data,
                    file_name="emails.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("No rows matched the selected filter.")

# =========================================================
# DASHBOARD
# =========================================================
elif menu == "Dashboard":
    st.header("Dashboard")

    df = load_emails()

    if not df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total", len(df))
        c2.metric("Pending", len(df[df["status"] == "Pending"]))
        c3.metric("Under Review", len(df[df["status"] == "Under Review"]))
        c4.metric("Rejected", len(df[df["status"] == "Rejected"]))

        st.subheader("Status Distribution")
        st.bar_chart(df["status"].value_counts())

        st.subheader("Recent Activity")
        st.dataframe(df.head(20), use_container_width=True)
    else:
        st.info("No data yet.")

# =========================================================
# HISTORY
# =========================================================
elif menu == "History":
    st.header("Saved Emails")

    df = load_emails()

    if df.empty:
        st.info("No saved emails yet.")
    else:
        st.dataframe(df, use_container_width=True)

        st.subheader("Open Saved Emails")
        for i, r in df.head(10).iterrows():
            with st.expander(r["subject"]):
                render_email_card(r["subject"], r["body"])

                x, y = st.columns(2)
                with x:
                    st.markdown("**Subject**")
                    st.code(r["subject"], language="text")
                    copy_button(r["subject"], "📋 Copy Subject", f"hist_subject_{i}")
                with y:
                    st.markdown("**Body**")
                    st.code(r["body"], language="text")
                    copy_button(r["body"], "📋 Copy Body", f"hist_body_{i}")
