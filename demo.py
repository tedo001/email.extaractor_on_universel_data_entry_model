import imapclient
import email
from email.policy import default
import os
log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
import logging
log_file = os.path.join(log_dir, 'app.log')
logging.basicConfig(filename=log_file, level=logging.INFO)
def save_emails_to_excel(email_data, output_path="outputs/emails.xlsx"):
    import pandas as pd
    df = pd.DataFrame(email_data)
    df.to_excel(output_path, index=False)
def get_email_data(email_account, password, search_criteria="UNSEEN", folder="INBOX"):
    try:
        with imapclient.IMAPClient("imap.gmail.com", ssl=True) as server:
            server.login(email_account, password)
            server.select_folder(folder)
            messages = server.search(search_criteria)
            email_data = []
            for msg_id, data in server.fetch(messages, ["RFC822"]).items():
                email_msg = email.message_from_bytes(data[b"RFC822"], policy=default)
                email_data.append({
                    "subject": str(email_msg["Subject"]),
                    "from": str(email_msg["From"]),
                    "body": _get_email_body(email_msg)
                })
            return email_data
    except Exception as e:
        raise Exception(f"Email Processing Failed: {str(e)}")
def _get_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode()
    else:
        return msg.get_payload(decode=True).decode()

