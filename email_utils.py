import smtplib
from email.message import EmailMessage
import os
import streamlit as st

def send_email_with_attachment(to_email: str, subject: str, body: str, attachment_bytes: bytes, filename: str) -> bool:
    # Read SMTP config from Streamlit secrets or environment variables
    smtp_server = None
    smtp_port = None
    user = None
    password = None
    # Prefer credentials from Streamlit session state if set via UI, then fall back to secrets/env
    smtp_server = st.session_state.get("smtp_server") or None
    smtp_port = st.session_state.get("smtp_port") or None
    user = st.session_state.get("smtp_user") or None
    password = st.session_state.get("smtp_password") or None
    if not smtp_server:
        try:
            smtp_server = st.secrets["smtp"]["server"]
            smtp_port = int(st.secrets["smtp"]["port"])
            user = st.secrets["smtp"]["user"]
            password = st.secrets["smtp"]["password"]
        except Exception:
            smtp_server = os.environ.get("SMTP_SERVER")
            smtp_port = int(os.environ.get("SMTP_PORT", 587))
            user = os.environ.get("SMTP_USER")
            password = os.environ.get("SMTP_PASSWORD")

    if not smtp_server or not user or not password:
        st.error("SMTP credentials not configured. Set Streamlit secrets, environment variables, or provide them in the SMTP Configuration panel.")
        return False

    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(attachment_bytes, maintype="application", subtype="pdf", filename=filename)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        st.error(f"SMTP send failed: {e}")
        return False
