import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="DC Email Generator", layout="centered")

st.title("📧 Document Control Email Generator")

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

# ---------------- GENERATE SINGLE EMAIL ----------------
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

        # ---------------- ENGLISH ----------------
        if language == "English":

            if status in ["Under Review", "Pending"]:
                email = f"""Subject: {subject}

Dear {recipient},

We are following up on {doc_type} No. {doc_number} regarding "{title}" ({revision}), submitted on {sent_date.strftime("%B %d, %Y")}.

Kindly provide an update on the review status at your earliest convenience.

Best regards,
[Sender Name]"""

            elif status in ["Approved with Comments", "Rejected"]:
                email = f"""Subject: {subject}

Dear {recipient},

Please be informed that {doc_type} No. {doc_number} regarding "{title}" ({revision}) has been {status}.

Kindly address all comments provided and submit the revised documents for further review.

Best regards,
[Sender Name]"""

            else:
                email = f"""Subject: {subject}

Dear {recipient},

Please find the document submitted for your review and record.

Best regards,
[Sender Name]"""

        # ---------------- ARABIC ----------------
        else:

            if status in ["Under Review", "Pending"]:
                email = f"""Subject: {subject}

السادة/ {recipient}،

تحية طيبة وبعد،

يرجى التكرم بالمتابعة بخصوص {doc_type} رقم {doc_number} المتعلق بـ "{title}" ({revision})، والذي تم إرساله بتاريخ {sent_date.strftime("%B %d, %Y")}، ولا يزال قيد المراجعة.

نرجو التكرم بتزويدنا بتحديث في أقرب وقت ممكن.

وتفضلوا بقبول فائق الاحترام،
[Sender Name]"""

            elif status in ["Approved with Comments", "Rejected"]:
                email = f"""Subject: {subject}

السادة/ {recipient}،

تحية طيبة وبعد،

نحيطكم علمًا بأن {doc_type} رقم {doc_number} المتعلق بـ "{title}" ({revision}) قد تم اعتماده بالحالة: {status}.

يرجى التكرم بمراجعة الملاحظات وإعادة تقديم المستند بعد التعديل.

وتفضلوا بقبول فائق الاحترام،
[Sender Name]"""

            else:
                email = f"""Subject: {subject}

السادة/ {recipient}،

تحية طيبة وبعد،

يرجى العلم بأنه تم تقديم المستند الخاص بـ "{title}" ({revision}) للمراجعة والاعتماد.

وتفضلوا بقبول فائق الاحترام،
[Sender Name]"""

        # ---------------- OUTPUT ----------------
        st.success("✅ Email Generated Successfully")

        st.markdown("### 📧 Generated Email")

        st.markdown(
            f"""
            <div style="
                background-color:#111;
                padding:25px;
                border-radius:12px;
                font-size:18px;
                font-family:Arial;
                line-height:1.8;
                white-space:pre-wrap;
                border:1px solid #555;
            ">
            {email}
            </div>
            """,
            unsafe_allow_html=True
        )

# ---------------- EXCEL SECTION ----------------
st.divider()
st.subheader("📂 Upload Excel File")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.write("Preview:")
    st.dataframe(df)

    status_filter = st.selectbox(
        "Filter by Status",
        ["All"] + list(df["Status"].unique())
    )

    if status_filter != "All":
        df = df[df["Status"] == status_filter]

    if st.button("Generate Emails from Excel"):

        results = []

        for index, row in df.iterrows():

            project_code = row["Project Code"]
            doc_number = row["Document Number"]
            doc_type = row["Document Type"]
            title = row["Title"]
            revision = row["Revision"]
            status = row["Status"]
            action_required = row["Action Required"]

            if status in ["Under Review", "Pending"]:
                recipient = "Consultant"
            elif status in ["Approved with Comments", "Rejected"]:
                recipient = "Contractor"
            else:
                recipient = "Client"

            subject = f"{project_code} - {doc_type} - {doc_number} - {title} - {action_required}"

            email = f"""Subject: {subject}

Dear {recipient},

Regarding {doc_type} No. {doc_number} "{title}" ({revision}),

Status: {status}

Action: {action_required}

Best regards,
[Sender Name]"""

            results.append({
                "Document Number": doc_number,
                "Status": status,
                "Email": email
            })

        result_df = pd.DataFrame(results)

        st.success("✅ Emails Generated Successfully")

        st.dataframe(result_df)

        csv = result_df.to_csv(index=False).encode('utf-8')

        st.download_button(
            label="📥 Download Emails as CSV",
            data=csv,
            file_name="generated_emails.csv",
            mime="text/csv"
        )