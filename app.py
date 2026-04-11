import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="DC Email Generator", layout="wide")

# ---------- Helpers ----------
def get_default_recipient(status: str) -> str:
    if status in ["Under Review", "Pending"]:
        return "Consultant"
    if status in ["Approved with Comments", "Rejected", "Approved"]:
        return "Contractor"
    return "Client"


def build_subject(project_code, doc_type, doc_number, title, email_type, urgent=False):
    subject = f"{project_code} - {doc_type} - {doc_number} - {title} - {email_type}"
    if urgent:
        subject = f"[URGENT] {subject}"
    return subject


def build_email_body(
    recipient,
    doc_type,
    doc_number,
    title,
    revision,
    sent_date,
    status,
    email_type,
    sender_name,
    company_name,
    urgent=False,
):
    sent_date_str = sent_date.strftime("%B %d, %Y")

    if email_type == "Follow-up":
        body = f"""Dear {recipient},

With reference to {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date_str}.

Please note that the document is currently {status}.

We kindly request your update on the review status at your earliest convenience to avoid any delay in project progress.
"""

    elif email_type == "Pending Reminder":
        body = f"""Dear {recipient},

This is a reminder regarding {doc_type} No. {doc_number} concerning "{title}" ({revision}), submitted on {sent_date_str}.

Please note that the document is still pending.

Kindly provide your response or update at your earliest convenience.
"""

    elif email_type == "Approved Notice":
        body = f"""Dear {recipient},

We are pleased to inform you that {doc_type} No. {doc_number} regarding "{title}" ({revision}) has been reviewed and approved.

Please proceed accordingly.
"""

    elif email_type == "Approved with Comments Notice":
        body = f"""Dear {recipient},

Please be informed that {doc_type} No. {doc_number} regarding "{title}" ({revision}) has been reviewed and marked as Approved with Comments.

Kindly address all comments and resubmit the revised document for further review and approval.
"""

    elif email_type == "Rejected Notice":
        body = f"""Dear {recipient},

Please be informed that {doc_type} No. {doc_number} regarding "{title}" ({revision}) has been reviewed and marked as Rejected.

Kindly review the comments and submit the revised document after addressing all remarks.
"""

    elif email_type == "Resubmit Request":
        body = f"""Dear {recipient},

Regarding {doc_type} No. {doc_number} concerning "{title}" ({revision}), the document has been reviewed with comments.

Kindly address all comments and resubmit the revised document for further review and approval.
"""

    elif email_type == "Submission":
        body = f"""Dear {recipient},

Please find attached {doc_type} No. {doc_number} regarding "{title}" ({revision}) submitted for your review and record.

Please proceed accordingly.
"""

    else:
        body = f"""Dear {recipient},

Please proceed regarding {doc_type} No. {doc_number} concerning "{title}" ({revision}).

Thank you.
"""

    if urgent:
        body += "\nThis matter is considered urgent and requires your immediate attention.\n"

    signature_name = sender_name if sender_name else "[Sender Name]"
    signature_company = company_name if company_name else "[Company Name]"

    body += f"""
Best regards,
{signature_name}
Document Control Department
{signature_company}
"""
    return body


def render_email_card(subject: str, recipient: str, body: str):
    st.markdown(
        f"""
        <div style="
            background-color:#ffffff;
            padding:22px;
            border-radius:14px;
            border:1px solid #e5e7eb;
            box-shadow:0 4px 14px rgba(0,0,0,0.08);
            font-family:Arial, sans-serif;
            color:#111827;
            margin-bottom:18px;
        ">
            <div style="font-size:18px; font-weight:700; margin-bottom:8px;">
                {subject}
            </div>

            <div style="font-size:13px; color:#6b7280; margin-bottom:18px;">
                To: {recipient}
            </div>

            <div style="
                font-size:15px;
                line-height:1.8;
                white-space:pre-wrap;
            ">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------- App ----------
st.title("📧 Engineering Document Control Email Automation Tool")

tab1, tab2 = st.tabs(["Single Email", "Bulk Emails (Excel)"])

# =======================
# Single Email Tab
# =======================
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        project_code = st.text_input("Project Code")
        doc_number = st.text_input("Document Number")
        doc_type = st.selectbox("Document Type", ["Submittal", "RFI", "Drawing", "Report"])
        title = st.text_input("Title")
        revision = st.text_input("Revision (e.g. Rev.00)")
        sent_date = st.date_input("Sent Date", value=date.today())

    with col2:
        status = st.selectbox(
            "Status",
            ["Under Review", "Pending", "Approved", "Approved with Comments", "Rejected", "Submitted"],
        )
        email_type = st.selectbox(
            "Email Type",
            [
                "Follow-up",
                "Pending Reminder",
                "Approved Notice",
                "Approved with Comments Notice",
                "Rejected Notice",
                "Resubmit Request",
                "Submission",
            ],
        )
        recipient_name = st.text_input("Recipient Name")
        sender_name = st.text_input("Sender Name")
        company_name = st.text_input("Company Name")
        urgent = st.checkbox("Mark as Urgent 🚨")

    if st.button("Generate Email", type="primary"):
        if not project_code or not doc_number or not title or not revision:
            st.error("❌ Please fill all required fields")
        else:
            recipient = recipient_name if recipient_name else get_default_recipient(status)
            subject = build_subject(project_code, doc_type, doc_number, title, email_type, urgent)
            body = build_email_body(
                recipient=recipient,
                doc_type=doc_type,
                doc_number=doc_number,
                title=title,
                revision=revision,
                sent_date=sent_date,
                status=status,
                email_type=email_type,
                sender_name=sender_name,
                company_name=company_name,
                urgent=urgent,
            )
            full_email = f"Subject: {subject}\n\n{body}"

            st.success("✅ Email Generated Successfully")
            st.markdown("### 📧 Email Preview")
            render_email_card(subject, recipient, body)

            copy_col1, copy_col2 = st.columns(2)
            with copy_col1:
                st.markdown("#### Copy Subject")
                st.code(subject, language="text")
            with copy_col2:
                st.markdown("#### Copy Email Body")
                st.code(body, language="text")

            with st.expander("Show Full Email"):
                st.code(full_email, language="text")

# =======================
# Bulk Email Tab
# =======================
with tab2:
    st.subheader("📂 Bulk Email Generator (Excel)")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        st.markdown("### Preview")
        st.dataframe(df, use_container_width=True)

        status_options = ["All"]
        if "Status" in df.columns:
            status_options += [x for x in df["Status"].dropna().astype(str).unique().tolist()]

        status_filter = st.selectbox("Filter by Status", status_options)

        filtered_df = df.copy()
        if status_filter != "All":
            filtered_df = filtered_df[filtered_df["Status"].astype(str) == status_filter]

        if st.button("Generate Emails from Excel", type="primary"):
            required_columns = [
                "Project Code",
                "Document Number",
                "Document Type",
                "Title",
                "Revision",
                "Status",
                "Action Required",
            ]

            missing_cols = [c for c in required_columns if c not in filtered_df.columns]
            if missing_cols:
                st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
            else:
                results = []

                for _, row in filtered_df.iterrows():
                    row_project_code = str(row["Project Code"])
                    row_doc_number = str(row["Document Number"])
                    row_doc_type = str(row["Document Type"])
                    row_title = str(row["Title"])
                    row_revision = str(row["Revision"])
                    row_status = str(row["Status"])
                    row_action = str(row["Action Required"])

                    row_recipient = get_default_recipient(row_status)

                    # map Action Required -> email type
                    if row_action == "Follow-up":
                        row_email_type = "Follow-up"
                    elif row_action == "Resubmit":
                        row_email_type = "Resubmit Request"
                    elif row_action == "None" and row_status == "Submitted":
                        row_email_type = "Submission"
                    elif row_status == "Approved":
                        row_email_type = "Approved Notice"
                    elif row_status == "Approved with Comments":
                        row_email_type = "Approved with Comments Notice"
                    elif row_status == "Rejected":
                        row_email_type = "Rejected Notice"
                    elif row_status == "Pending":
                        row_email_type = "Pending Reminder"
                    else:
                        row_email_type = row_action

                    row_subject = build_subject(
                        row_project_code,
                        row_doc_type,
                        row_doc_number,
                        row_title,
                        row_email_type,
                        urgent=False,
                    )

                    row_body = build_email_body(
                        recipient=row_recipient,
                        doc_type=row_doc_type,
                        doc_number=row_doc_number,
                        title=row_title,
                        revision=row_revision,
                        sent_date=date.today(),
                        status=row_status,
                        email_type=row_email_type,
                        sender_name=sender_name,
                        company_name=company_name,
                        urgent=False,
                    )

                    row_full_email = f"Subject: {row_subject}\n\n{row_body}"

                    results.append(
                        {
                            "Document Number": row_doc_number,
                            "Status": row_status,
                            "Email": row_full_email,
                        }
                    )

                result_df = pd.DataFrame(results)

                st.success("✅ Emails Generated Successfully")
                st.dataframe(result_df, use_container_width=True)

                st.markdown("### 📧 Bulk Email Preview")

                for item in results:
                    subject_line = item["Email"].split("\n")[0].replace("Subject: ", "")
                    body_part = "\n".join(item["Email"].split("\n")[2:])
                    render_email_card(subject_line, "Team", body_part)

                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.markdown("**Copy Subject**")
                        st.code(subject_line, language="text")
                    with col_b:
                        st.markdown("**Copy Body**")
                        st.code(body_part, language="text")

                csv = result_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Emails as CSV",
                    data=csv,
                    file_name="generated_emails.csv",
                    mime="text/csv",
                )
