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
status = st.selectbox("Status", ["Under Review", "Pending", "Approved with Comments", "Rejected", "Submitted"])
sent_date = st.date_input("Sent Date", value=date.today())
action_required = st.selectbox("Action Required", ["Follow-up", "Resubmit", "None"])
language = st.selectbox("Language", ["English", "Arabic"])

# ---------------- GENERATE ----------------
if st.button("Generate Email"):

    if not project_code or not doc_number or not title or not revision:
        st.error("❌ Please fill all required fields")

    else:

        # Recipient Logic
        if status in ["Under Review", "Pending"]:
            recipient = "Consultant"
        elif status in ["Approved with Comments", "Rejected"]:
            recipient = "Contractor"
        else:
            recipient = "Client"

        subject_action = action_required if action_required != "None" else status
        subject = f"{project_code} - {doc_type} - {doc_number} - {title} - {subject_action}"

        # ---------------- EMAIL BODY ----------------
        email = f"""Subject: {subject}

Dear {recipient},

With reference to {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date.strftime("%B %d, %Y")}.

Please note that the document is currently {status}.

We kindly request your update or necessary action at your earliest convenience to avoid any delay in project progress.

Best regards,
[Sender Name]
Document Control Department
[Company Name]"""

        # ---------------- GMAIL STYLE UI ----------------
        st.success("✅ Email Generated Successfully")

        st.markdown("### 📧 Email Preview")

        st.markdown(
            f"""
            <div style="
                background-color:#ffffff;
                padding:20px;
                border-radius:10px;
                border:1px solid #ddd;
                box-shadow:0 2px 8px rgba(0,0,0,0.1);
                font-family:Arial;
                color:#000;
            ">

                <div style="font-size:18px; font-weight:bold; margin-bottom:10px;">
                    {subject}
                </div>

                <div style="font-size:14px; color:#555; margin-bottom:20px;">
                    To: {recipient}
                </div>

                <div style="
                    font-size:15px;
                    line-height:1.8;
                    white-space:pre-wrap;
                ">
                    Dear {recipient},

                    With reference to {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date.strftime("%B %d, %Y")}.

                    Please note that the document is currently {status}.

                    We kindly request your update or necessary action at your earliest convenience to avoid any delay in project progress.

                    Best regards,
                    [Sender Name]
                    Document Control Department
                    [Company Name]
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

# ---------------- EXCEL ----------------
st.divider()
st.subheader("📂 Upload Excel File")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.dataframe(df)

    if st.button("Generate Emails from Excel"):

        results = []

        for index, row in df.iterrows():

            subject = f"{row['Project Code']} - {row['Document Type']} - {row['Document Number']} - {row['Title']} - {row['Action Required']}"

            email = f"""Subject: {subject}

Dear Team,

Regarding {row['Document Type']} No. {row['Document Number']} "{row['Title']}" ({row['Revision']}),

Status: {row['Status']}

Action Required: {row['Action Required']}

Best regards,
[Sender Name]"""

            results.append(email)

        st.success("✅ Bulk Emails Generated")

        for e in results:
            st.markdown("---")
            st.text(e)
