import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="DC Email Generator", layout="centered")

st.title("📧 Engineering Email Automation Tool")

# ---------------- INPUTS ----------------
project_code = st.text_input("Project Code")
doc_number = st.text_input("Document Number")
doc_type = st.selectbox("Document Type", ["Submittal", "RFI", "Drawing", "Report"])
title = st.text_input("Title")
revision = st.text_input("Revision (e.g. Rev.00)")

status = st.selectbox("Status", [
    "Under Review", "Pending", "Approved", "Approved with Comments", "Rejected", "Submitted"
])

email_type = st.selectbox("Email Type", [
    "Follow-up", "Pending Reminder", "Approved Notice", "Resubmit Request"
])

sent_date = st.date_input("Sent Date", value=date.today())

recipient_name = st.text_input("Recipient Name")
sender_name = st.text_input("Sender Name")
company_name = st.text_input("Company Name")

urgent = st.checkbox("Mark as Urgent 🚨")

# ---------------- GENERATE EMAIL ----------------
if st.button("Generate Email"):

    if not project_code or not doc_number or not title or not revision or not sender_name:
        st.error("❌ Please fill all required fields")

    else:

        recipient = recipient_name if recipient_name else "Consultant"

        subject = f"{project_code} - {doc_type} - {doc_number} - {title} - {email_type}"

        if urgent:
            subject = "[URGENT] " + subject

        # ---------------- TEMPLATE LOGIC ----------------
        if email_type in ["Follow-up", "Pending Reminder"]:
            body = f"""Dear {recipient},

With reference to {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date.strftime("%B %d, %Y")}.

Please note that the document is still under review.

We kindly request your update on the review status at your earliest convenience to avoid any delay in project progress.
"""

        elif email_type == "Approved Notice":
            body = f"""Dear {recipient},

We are pleased to inform you that {doc_type} No. {doc_number} regarding "{title}" ({revision}) has been reviewed and approved.

Please proceed accordingly.
"""

        elif email_type == "Resubmit Request":
            body = f"""Dear {recipient},

Regarding {doc_type} No. {doc_number} "{title}" ({revision}), the document has been reviewed with comments.

Kindly address all comments and resubmit the revised document for further review and approval.
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

        full_email = f"Subject: {subject}\n\n{body}"

        # ---------------- UI ----------------
        st.success("✅ Email Generated Successfully")

        st.markdown("### 📧 Email Preview")

        st.markdown(
            f"""
            <div style="
                background-color:#ffffff;
                padding:20px;
                border-radius:12px;
                border:1px solid #e0e0e0;
                box-shadow:0 4px 12px rgba(0,0,0,0.08);
                font-family:Arial;
                color:#000;
            ">
                <div style="font-size:17px; font-weight:bold; margin-bottom:10px;">
                    {subject}
                </div>

                <div style="font-size:13px; color:#666; margin-bottom:15px;">
                    To: {recipient}
                </div>

                <div style="font-size:15px; line-height:1.8; white-space:pre-wrap;">
                    {body}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # ---------------- COPY BUTTONS ----------------
        st.markdown("### 📋 Copy Options")

        col1, col2 = st.columns(2)

        with col1:
            st.code(subject, language="text")
            st.caption("Copy Subject")

        with col2:
            st.code(body, language="text")
            st.caption("Copy Email Body")

# ---------------- BULK EMAIL ----------------
st.divider()
st.subheader("📂 Bulk Email Generator (Excel)")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    st.dataframe(df)

    if st.button("Generate Bulk Emails"):

        results = []

        for index, row in df.iterrows():

            subject = f"{row['Project Code']} - {row['Document Type']} - {row['Document Number']} - {row['Title']} - {row['Action Required']}"

            email = f"""Subject: {subject}

Dear Team,

Regarding {row['Document Type']} No. {row['Document Number']} "{row['Title']}" ({row['Revision']}),

Status: {row['Status']}

Action Required: {row['Action Required']}

Best regards,
{sender_name if sender_name else "[Sender Name]"}"""

            results.append(email)

        st.success("✅ Bulk Emails Generated")

        for e in results:

            subject_line = e.split("\n")[0].replace("Subject: ", "")
            body_part = "\n".join(e.split("\n")[2:])

            html_block = f"""
            <div style="
                background-color:#ffffff;
                padding:20px;
                border-radius:12px;
                border:1px solid #e0e0e0;
                box-shadow:0 4px 12px rgba(0,0,0,0.08);
                font-family:Arial;
                color:#000;
                margin-bottom:25px;
            ">
                <div style="font-size:17px; font-weight:bold; margin-bottom:8px;">
                    {subject_line}
                </div>

                <div style="font-size:13px; color:#666; margin-bottom:15px;">
                    To: Team
                </div>

                <div style="font-size:15px; line-height:1.8; white-space:pre-wrap;">
                    {body_part}
                </div>
            </div>
            """

            st.markdown(html_block, unsafe_allow_html=True)

            st.code(e)

        # Download
        csv = pd.DataFrame(results, columns=["Emails"]).to_csv(index=False).encode('utf-8')

        st.download_button(
            label="📥 Download All Emails",
            data=csv,
            file_name="emails.csv",
            mime="text/csv"
        )
