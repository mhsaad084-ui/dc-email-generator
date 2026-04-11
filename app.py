import html
import io
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Engineering Document Control Email System",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_NAME = "emails.db"
conn = sqlite3.connect(DB_NAME, check_same_thread=False)


def init_db():
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS emails (
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
            workflow_action TEXT,
            stage TEXT,
            submitted_by TEXT,
            submitted_to TEXT,
            sent_date TEXT,
            received_date TEXT,
            recipient TEXT,
            sender_name TEXT,
            company_name TEXT,
            language TEXT,
            subject TEXT,
            body TEXT,
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
        "discipline",
        "title",
        "revision",
        "status",
        "action_required",
        "workflow_action",
        "stage",
        "submitted_by",
        "submitted_to",
        "sent_date",
        "received_date",
        "recipient",
        "sender_name",
        "company_name",
        "language",
        "subject",
        "body",
        "source",
    ]

    values = []
    for c in cols:
        if c == "created_at":
            values.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            values.append(kwargs.get(c, ""))

    placeholders = ",".join(["?"] * len(cols))
    conn.execute(
        f"INSERT INTO emails ({','.join(cols)}) VALUES ({placeholders})",
        values,
    )
    conn.commit()


def load_emails():
    return pd.read_sql("SELECT * FROM emails ORDER BY id DESC", conn)


def update_email_status(email_id: int, new_status: str):
    workflow_action = workflow_engine(new_status)[1]
    conn.execute(
        "UPDATE emails SET status = ?, stage = ?, workflow_action = ? WHERE id = ?",
        (new_status, new_status, workflow_action, email_id),
    )
    conn.commit()


def delete_all_emails():
    conn.execute("DELETE FROM emails")
    conn.commit()


def norm(value, default=""):
    if pd.isna(value):
        return default
    return str(value).strip()


def get_default_recipient(status: str, submitted_by: str = "", submitted_to: str = "") -> str:
    submitted_by_l = submitted_by.lower()
    submitted_to_l = submitted_to.lower()
    internal_keywords = ["team", "department", "technical office", "internal", "engineering"]

    if any(k in submitted_by_l for k in internal_keywords) and any(k in submitted_to_l for k in internal_keywords):
        return "Team"

    if status in ["Pending", "Under Review"]:
        return "Consultant"
    if status in ["Approved", "Approved with Comments", "Rejected"]:
        return "Contractor"
    if status == "Submitted":
        return "Client"
    return "Team"


def workflow_engine(status: str):
    if status == "Submitted":
        return "Consultant", "Review Required"
    if status == "Under Review":
        return "Consultant", "Follow-up"
    if status == "Approved":
        return "Contractor", "Proceed"
    if status == "Approved with Comments":
        return "Contractor", "Revise with Comments"
    if status == "Rejected":
        return "Contractor", "Revise & Resubmit"
    if status == "Pending":
        return "Consultant", "Urgent Follow-up"
    return "Team", "Info"


def detect_email_type(status: str, action_required: str) -> str:
    if status in ["Pending", "Under Review"] or action_required in ["Follow-up", "Urgent Follow-up"]:
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


def get_priority_tag(status: str, action_required: str) -> str:
    if action_required == "Urgent Follow-up" or status == "Pending":
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


def build_subject(project_code, document_type, document_number, title, status, action_required):
    prefix = get_priority_tag(status, action_required)
    suffix = action_required if action_required not in ["", "None", "No Action"] else status
    subject = f"{project_code} - {document_type} - {document_number} - {title} - {suffix}"
    return f"{prefix} {subject}".strip()


def build_email_body(
    recipient,
    document_type,
    document_number,
    title,
    revision,
    status,
    action_required,
    sender_name,
    company_name,
    language,
    sent_date="",
):
    sender_name = sender_name or "[Sender Name]"
    company_name = company_name or "[Company Name]"
    email_type = detect_email_type(status, action_required)

    if language == "English":
        if email_type == "FOLLOW_UP":
            if document_type == "RFI":
                intro = f'We are writing to follow up on RFI No. {document_number} regarding "{title}" ({revision}).'
                ask = "Kindly provide your response or clarification at your earliest convenience to ensure the continuity of the project workflow."
            elif document_type in ["Submittal", "Drawing"]:
                intro = f'We are writing to follow up on the review status of the {document_type.lower()} for "{title}" ({document_number}, {revision}).'
                ask = "Kindly provide an update or the reviewed outcome for this submission at your earliest convenience to ensure the continuity of the project workflow."
            else:
                intro = f'We are writing to follow up on {document_type} No. {document_number} regarding "{title}" ({revision}).'
                ask = "Kindly provide your update at your earliest convenience."

            sent_line = f'\n\nThe subject document was submitted on {sent_date} and is currently "{status}".' if sent_date else ""

            return f"""Dear {recipient},

{intro}{sent_line}

{ask}

Your prompt attention to this matter is highly appreciated.

Best regards,
{sender_name}
Document Control Department
{company_name}"""

        if email_type == "RESUBMIT_REQUEST":
            if status == "Approved with Comments":
                main = f'Please be informed that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been reviewed and marked as "Approved with Comments".'
                req = "Kindly address all comments and resubmit the revised document for further review and approval."
            else:
                main = f'Please be informed that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been reviewed and marked as "Rejected".'
                req = "Kindly revise and resubmit the document after addressing all comments."

            return f"""Dear {recipient},

{main}

{req}

Best regards,
{sender_name}
Document Control Department
{company_name}"""

        if email_type == "SUBMISSION":
            if document_type == "Report":
                main = f'Please find attached {document_type} No. {document_number} regarding "{title}" ({revision}) submitted for your review and record.'
                follow = ""
            else:
                main = f'Please find attached {document_type} No. {document_number} regarding "{title}" ({revision}), submitted for your review.'
                follow = "\n\nKindly review and provide your comments or approval at your earliest convenience."

            return f"""Dear {recipient},

{main}{follow}

Best regards,
{sender_name}
Document Control Department
{company_name}"""

        if email_type == "APPROVED_NOTICE":
            return f"""Dear {recipient},

We are pleased to inform you that the {document_type.lower()} for "{title}" ({document_number}, {revision}) has been approved.

Please proceed accordingly.

Best regards,
{sender_name}
Document Control Department
{company_name}"""

        if email_type == "MISSING_ATTACHMENT":
            return f"""Dear {recipient},

Please accept our apologies.

Please find the correct attachment for {document_type} No. {document_number} regarding "{title}" ({revision}).

Best regards,
{sender_name}
Document Control Department
{company_name}"""

        return f"""Dear {recipient},

Please proceed regarding {document_type} No. {document_number} concerning "{title}" ({revision}).

Best regards,
{sender_name}
Document Control Department
{company_name}"""

    if email_type == "FOLLOW_UP":
        if document_type == "RFI":
            intro = f'نود المتابعة بخصوص RFI رقم {document_number} الخاص بـ "{title}" ({revision}).'
            ask = "نرجو التكرم بتزويدنا بالرد أو الإيضاح في أقرب وقت ممكن لضمان استمرارية سير العمل بالمشروع."
        elif document_type in ["Submittal", "Drawing"]:
            intro = f'نود المتابعة بخصوص حالة مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}).'
            ask = "نرجو التكرم بتزويدنا بالتحديث أو نتيجة المراجعة في أقرب وقت ممكن لضمان استمرارية سير العمل بالمشروع."
        else:
            intro = f'نود المتابعة بخصوص {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}).'
            ask = "نرجو التكرم بتزويدنا بالتحديث في أقرب وقت ممكن."

        sent_line = f'\n\nتم إرسال المستند بتاريخ {sent_date} ولا يزال بحالة "{status}".' if sent_date else ""

        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{intro}{sent_line}

{ask}

شاكرين حسن تعاونكم.

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""

    if email_type == "RESUBMIT_REQUEST":
        if status == "Approved with Comments":
            main = f'نحيطكم علمًا بأنه تمت مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}) وتم اعتماده مع ملاحظات.'
            req = "يرجى مراجعة جميع الملاحظات وإعادة تقديم المستند بعد التعديل لاستكمال المراجعة والاعتماد."
        else:
            main = f'نحيطكم علمًا بأنه تمت مراجعة {document_type} الخاص بـ "{title}" ({document_number}، {revision}) وتم رفضه.'
            req = "يرجى تعديل المستند وإعادة التقديم بعد معالجة جميع الملاحظات."

        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{main}

{req}

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""

    if email_type == "SUBMISSION":
        if document_type == "Report":
            main = f'نرفق لسيادتكم {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) للمراجعة والحفظ.'
            follow = ""
        else:
            main = f'نرفق لسيادتكم {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) للمراجعة.'
            follow = "\n\nنرجو التكرم بالمراجعة وإفادتنا بالملاحظات أو الاعتماد في أقرب وقت ممكن."

        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

{main}{follow}

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""

    if email_type == "APPROVED_NOTICE":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأنه تم اعتماد {document_type} الخاص بـ "{title}" ({document_number}، {revision}).

يرجى المتابعة واتخاذ اللازم وفقًا لذلك.

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""

    if email_type == "MISSING_ATTACHMENT":
        return f"""السادة/ {recipient}،

تحية طيبة وبعد،

نعتذر عن السهو السابق.

نرفق لسيادتكم المرفق الصحيح الخاص بـ {document_type} رقم {document_number} بعنوان "{title}" ({revision}).

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""

    return f"""السادة/ {recipient}،

تحية طيبة وبعد،

يرجى التكرم بمراجعة {document_type} رقم {document_number} الخاص بـ "{title}" ({revision}) واتخاذ اللازم.

وتفضلوا بقبول فائق الاحترام،
{sender_name}
قسم ضبط المستندات
{company_name}"""


def render_email_card(subject: str, body: str):
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


def convert_excel(df: pd.DataFrame):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Emails")
    return output.getvalue()


st.title("📧 Engineering Document Control Email System")

menu = st.sidebar.radio("Menu", ["Single", "Bulk", "Dashboard", "History"])

if menu == "Single":
    st.header("Single Email")

    c1, c2 = st.columns(2)

    with c1:
        project_code = st.text_input("Project Code")
        doc_no = st.text_input("Document No")
        doc_type = st.selectbox("Type", ["Submittal", "RFI", "Drawing", "Report"])
        discipline = st.text_input("Discipline")
        title = st.text_input("Title")
        rev = st.text_input("Revision", placeholder="Rev.00")
        sent_date = st.text_input("Sent Date", placeholder="April 16, 2026")
        received_date = st.text_input("Received Date", placeholder="Optional")

    with c2:
        status = st.selectbox(
            "Status",
            ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )
        action_required = st.selectbox(
            "Action Required",
            ["Follow-up", "Urgent Follow-up", "Resubmit", "No Action", "Missing Attachment Correction"],
        )
        submitted_by = st.text_input("Submitted By")
        submitted_to = st.text_input("Submitted To")
        recipient = st.text_input("Recipient")
        sender = st.text_input("Sender")
        company_name = st.text_input("Company Name")
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
            recipient_final = recipient if recipient else get_default_recipient(status, submitted_by, submitted_to)
            workflow_action = workflow_engine(status)[1]
            subject = build_subject(project_code, doc_type, doc_no, title, status, action_required)
            body = build_email_body(
                recipient=recipient_final,
                document_type=doc_type,
                document_number=doc_no,
                title=title,
                revision=rev,
                status=status,
                action_required=action_required,
                sender_name=sender,
                company_name=company_name,
                language=lang,
                sent_date=sent_date,
            )

            save_email(
                project_code=project_code,
                document_number=doc_no,
                document_type=doc_type,
                discipline=discipline,
                title=title,
                revision=rev,
                status=status,
                action_required=action_required,
                workflow_action=workflow_action,
                stage=status,
                submitted_by=submitted_by,
                submitted_to=submitted_to,
                sent_date=sent_date,
                received_date=received_date,
                recipient=recipient_final,
                sender_name=sender,
                company_name=company_name,
                language=lang,
                subject=subject,
                body=body,
                source="single",
            )

            st.success("Generated")
            render_email_card(subject, body)

            a, b = st.columns(2)
            with a:
                st.subheader("Subject")
                st.code(subject, language="text")
                copy_button(subject, "📋 Copy Subject", "single_subject")
            with b:
                st.subheader("Body")
                st.code(body, language="text")
                copy_button(body, "📋 Copy Body", "single_body")

elif menu == "Bulk":
    st.header("Bulk Emails")

    sender = st.text_input("Sender Name")
    company_name = st.text_input("Company Name")
    lang = st.selectbox("Language", ["English", "Arabic"])

    file = st.file_uploader("Upload Excel", type="xlsx")

    if file:
        df = pd.read_excel(file)

        required_cols = [
            "Project Code",
            "Document Number",
            "Document Type",
            "Title",
            "Revision",
            "Status",
            "Action Required",
        ]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"Missing columns: {', '.join(missing)}")
            st.stop()

        for optional in ["Discipline", "Submitted By", "Submitted To", "Sent Date", "Received Date"]:
            if optional not in df.columns:
                df[optional] = ""

        st.dataframe(df, use_container_width=True)

        status_filter = st.selectbox(
            "Filter by Status",
            ["All", "Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )

        if st.button("Generate Emails", type="primary"):
            results = []

            for idx, r in df.iterrows():
                row_status = norm(r.get("Status", ""))

                if status_filter != "All" and row_status != status_filter:
                    continue

                row_project = norm(r.get("Project Code", ""))
                row_doc_type = norm(r.get("Document Type", ""))
                row_doc_no = norm(r.get("Document Number", ""))
                row_title = norm(r.get("Title", ""))
                row_rev = norm(r.get("Revision", ""))
                row_action = norm(r.get("Action Required", ""))
                row_discipline = norm(r.get("Discipline", ""))
                row_submitted_by = norm(r.get("Submitted By", ""))
                row_submitted_to = norm(r.get("Submitted To", ""))
                row_sent_date = norm(r.get("Sent Date", ""))
                row_received_date = norm(r.get("Received Date", ""))

                recipient_final = get_default_recipient(row_status, row_submitted_by, row_submitted_to)
                workflow_action = workflow_engine(row_status)[1]
                subject = build_subject(row_project, row_doc_type, row_doc_no, row_title, row_status, row_action)
                body = build_email_body(
                    recipient=recipient_final,
                    document_type=row_doc_type,
                    document_number=row_doc_no,
                    title=row_title,
                    revision=row_rev,
                    status=row_status,
                    action_required=row_action,
                    sender_name=sender,
                    company_name=company_name,
                    language=lang,
                    sent_date=row_sent_date,
                )

                save_email(
                    project_code=row_project,
                    document_number=row_doc_no,
                    document_type=row_doc_type,
                    discipline=row_discipline,
                    title=row_title,
                    revision=row_rev,
                    status=row_status,
                    action_required=row_action,
                    workflow_action=workflow_action,
                    stage=row_status,
                    submitted_by=row_submitted_by,
                    submitted_to=row_submitted_to,
                    sent_date=row_sent_date,
                    received_date=row_received_date,
                    recipient=recipient_final,
                    sender_name=sender,
                    company_name=company_name,
                    language=lang,
                    subject=subject,
                    body=body,
                    source="bulk",
                )

                results.append(
                    {
                        "Doc": row_doc_no,
                        "Status": row_status,
                        "Recipient": recipient_final,
                        "Workflow Action": workflow_action,
                        "Subject": subject,
                        "Body": body,
                    }
                )

            st.success("Done")

            result_df = pd.DataFrame(results)
            if not result_df.empty:
                st.dataframe(
                    result_df[["Doc", "Status", "Recipient", "Workflow Action", "Subject"]],
                    use_container_width=True,
                )

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

elif menu == "Dashboard":
    st.header("Dashboard")

    df = load_emails()

    if not df.empty:
        a, b, c, d = st.columns(4)
        a.metric("Total", len(df))
        b.metric("Pending", len(df[df["status"] == "Pending"]))
        c.metric("Under Review", len(df[df["status"] == "Under Review"]))
        d.metric("Rejected", len(df[df["status"] == "Rejected"]))

        st.subheader("Status Distribution")
        st.bar_chart(df["status"].value_counts())

        st.subheader("Workflow Actions")
        st.bar_chart(df["workflow_action"].value_counts())

        st.subheader("Recent Activity")
        st.dataframe(df.head(20), use_container_width=True)
    else:
        st.info("No data yet.")

elif menu == "History":
    st.header("Saved Emails")

    df = load_emails()

    if df.empty:
        st.info("No saved emails yet.")
    else:
        search_doc = st.text_input("Search Document Number")
        status_filter = st.selectbox(
            "Filter Status",
            ["All", "Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )

        filtered = df.copy()

        if search_doc:
            filtered = filtered[filtered["document_number"].str.contains(search_doc, case=False, na=False)]

        if status_filter != "All":
            filtered = filtered[filtered["status"] == status_filter]

        st.dataframe(filtered, use_container_width=True)

        st.subheader("Open Saved Emails")
        for i, r in filtered.head(20).iterrows():
            with st.expander(f"{r['document_number']} | {r['title']} | {r['status']}"):
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

                new_status = st.selectbox(
                    "Change Status",
                    ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"],
                    key=f"status_{i}",
                    index=["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"].index(r["status"])
                    if r["status"] in ["Pending", "Under Review", "Approved", "Approved with Comments", "Rejected", "Submitted"]
                    else 0,
                )

                if st.button("Update Status", key=f"btn_{i}"):
                    update_email_status(int(r["id"]), new_status)
                    st.success("Updated")
                    st.rerun()

        if st.button("Delete All History"):
            delete_all_emails()
            st.success("History deleted")
            st.rerun()
