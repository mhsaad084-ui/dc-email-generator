import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="DC Email Generator", layout="wide")

st.title("📧 Engineering Document Control Email Automation Tool")

# ---------- Helpers ----------
def get_default_recipient(status):
    if status in ["Under Review", "Pending"]:
        return "Consultant"
    elif status in ["Approved with Comments", "Rejected", "Approved"]:
        return "Contractor"
    else:
        return "Client"

def build_subject(project_code, doc_type, doc_number, title, email_type, urgent):
    subject = f"{project_code} - {doc_type} - {doc_number} - {title} - {email_type}"
    if urgent:
        subject = "[URGENT] " + subject
    return subject

def build_body(recipient, doc_type, doc_number, title, revision, sent_date, status, email_type, sender_name, company_name, urgent):
    
    sent_date = sent_date.strftime("%B %d, %Y")

    if email_type in ["Follow-up", "Pending Reminder"]:
        body = f"""Dear {recipient},

With reference to {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date}.

Please note that the document is currently {status}.

We kindly request your update at your earliest convenience.
"""

    elif email_type == "Approved Notice":
        body = f"""Dear {recipient},

We are pleased to inform you that {doc_type} No. {doc_number} "{title}" ({revision}) has been approved.

Please proceed accordingly.
"""

    elif email_type == "Resubmit Request":
        body = f"""Dear {recipient},

Regarding {doc_type} No. {doc_number} "{title}" ({revision}), the document has been reviewed with comments.

Kindly address all comments and resubmit.
"""

    else:
        body = "Dear Team,\n\nPlease proceed accordingly."

    if urgent:
        body += "\n\nThis matter is urgent and requires immediate attention."

    body += f"""

Best regards,
{sender_name}
Document Control Department
{company_name}
"""

    return body

def render_email(subject, recipient, body):
    st.markdown(f"""
    <div style="
        background:#ffffff;
        padding:20px;
        border-radius:12px;
        border:1px solid #e0e0e0;
        box-shadow:0 4px 10px rgba(0,0,0,0.1);
        color:#000;
        font-family:Arial;
    ">
        <div style="font-weight:bold;font-size:18px;">{subject}</div>
        <div style="color:gray;margin-bottom:10px;">To: {recipient}</div>
        <div style="white-space:pre-wrap;line-height:1.8;">{body}</div>
    </div>
    """, unsafe_allow_html=True)

# ---------- TABS ----------
tab1, tab2 = st.tabs(["Single Email", "Bulk Emails (Excel)"])

# ================= SINGLE =================
with tab1:

    col1, col2 = st.columns(2)

    with col1:
        project_code = st.text_input("Project Code")
        doc_number = st.text_input("Document Number")
        doc_type = st.selectbox("Document Type", ["Submittal", "RFI", "Drawing", "Report"])
        title = st.text_input("Title")
        revision = st.text_input("Revision")

    with col2:
        status = st.selectbox("Status", ["Under Review", "Pending", "Approved", "Approved with Comments", "Rejected", "Submitted"])
        email_type = st.selectbox("Email Type", ["Follow-up", "Pending Reminder", "Approved Notice", "Resubmit Request"])
        recipient_name = st.text_input("Recipient Name")
        sender_name = st.text_input("Sender Name")
        company_name = st.text_input("Company Name")
        urgent = st.checkbox("Mark as Urgent 🚨")

    sent_date = st.date_input("Sent Date", value=date.today())

    # Status Color
    status_colors = {
        "Pending": "orange",
        "Under Review": "blue",
        "Approved": "green",
        "Approved with Comments": "purple",
        "Rejected": "red",
    }

    st.markdown(
        f"<span style='color:{status_colors.get(status,'white')};font-weight:bold;'>Status: {status}</span>",
        unsafe_allow_html=True
    )

    # Preview Subject
    if project_code and doc_number and title:
        preview_subject = f"{project_code} - {doc_type} - {doc_number} - {title}"
        st.info(f"Preview Subject: {preview_subject}")

    # Generate
    if st.button("Generate Email"):

        if not project_code:
            st.warning("Project Code required")
        elif not doc_number:
            st.warning("Document Number required")
        elif not title:
            st.warning("Title required")
        elif not revision:
            st.warning("Revision required")
        elif not sender_name:
            st.warning("Sender Name required")

        else:
            recipient = recipient_name if recipient_name else get_default_recipient(status)

            subject = build_subject(project_code, doc_type, doc_number, title, email_type, urgent)
            body = build_body(recipient, doc_type, doc_number, title, revision, sent_date, status, email_type, sender_name, company_name, urgent)

            full_email = f"Subject: {subject}\n\n{body}"

            st.success("Email Generated Successfully")

            render_email(subject, recipient, body)

            colA, colB = st.columns(2)

            with colA:
                st.code(subject)
                st.caption("Copy Subject")

            with colB:
                st.code(body)
                st.caption("Copy Body")

    if st.button("Reset Form"):
        st.rerun()

# ================= BULK =================
with tab2:

    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if uploaded_file:

        df = pd.read_excel(uploaded_file)
        st.dataframe(df)

        status_filter = st.selectbox("Filter by Status", ["All"] + list(df["Status"].unique()))

        if status_filter != "All":
            df = df[df["Status"] == status_filter]

        if st.button("Generate Emails from Excel"):

            results = []

            for _, row in df.iterrows():

                subject = f"{row['Project Code']} - {row['Document Type']} - {row['Document Number']} - {row['Title']} - {row['Action Required']}"

                email = f"""Subject: {subject}

Dear Team,

Regarding {row['Document Type']} No. {row['Document Number']} "{row['Title']}" ({row['Revision']}),

Status: {row['Status']}

Action Required: {row['Action Required']}

Best regards,
{sender_name if sender_name else '[Sender Name]'}"""

                results.append(email)

            st.success("Emails Generated Successfully")

            result_df = pd.DataFrame(results, columns=["Email"])
            st.dataframe(result_df)

            for e in results:
                subject_line = e.split("\n")[0].replace("Subject: ", "")
                body_part = "\n".join(e.split("\n")[2:])

                render_email(subject_line, "Team", body_part)

                col1, col2 = st.columns(2)
                with col1:
                    st.code(subject_line)
                with col2:
                    st.code(body_part)

            csv = result_df.to_csv(index=False).encode('utf-8')

            st.download_button(
                "Download Emails as CSV",
                csv,
                "emails.csv",
                "text/csv"
            )
