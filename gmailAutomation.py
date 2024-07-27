import imaplib
import email
from email.header import decode_header
from plyer import notification
import pyttsx3
import time
from dotenv import load_dotenv
import os

# load .env file
load_dotenv()

username = os.getenv('EMAIL')
password = os.getenv('PASSWORD')

if username is None or password is None:
    raise ValueError("Please set your EMAIL and PASSWORD environment variables in the .env file")

print("Username:", username)  # print username
print("Password:", "*" * len(password))  # masked password

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

last_seen_id = None

def initialize_last_seen_id():
    global last_seen_id
    try:

        # connect Server
        mail = imaplib.IMAP4_SSL("imap.gmail.com")

        # login
        mail.login(username, password)

        # what to check
        mail.select("inbox")

        # all mails
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        
        if email_ids:
            last_seen_id = int(email_ids[-1])  # initialise new email ID

        # Close and logOut
        mail.close()
        mail.logout()
    except imaplib.IMAP4.error as e:
        print("Failed to authenticate:", e)

def check_email():
    global last_seen_id
    try:
        # connect server
        mail = imaplib.IMAP4_SSL("imap.gmail.com")

        # Login
        mail.login(username, password)

        # mailbox 
        mail.select("inbox")

        # mails in inbox
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()

        new_emails = []
        for email_id in email_ids:
            if last_seen_id is None or int(email_id) > last_seen_id:
                new_emails.append(email_id)

        if new_emails:
            last_seen_id = int(new_emails[-1])  # last seen ID = new email ID

        for email_id in new_emails:

            # get mail by ID
            status, msg = mail.fetch(email_id, "(RFC822)")
            for response_part in msg:
                if isinstance(response_part, tuple):

                    # content
                    msg = email.message_from_bytes(response_part[1])

                    # subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")

                    # sender
                    from_, encoding = decode_header(msg.get("From"))[0]
                    if isinstance(from_, bytes):
                        from_ = from_.decode(encoding if encoding else "utf-8")

                    # print details
                    print("From:", from_)
                    print("Subject:", subject)
                    

                    # short if msg is long
                    truncated_subject = (subject[:61] + '...') if len(subject) > 64 else subject
                    truncated_from = (from_[:61] + '...') if len(from_) > 64 else from_

                    # desktop notify
                    notification.notify(
                        title=f"New Email from {truncated_from}",
                        message=truncated_subject,
                        timeout=10  # Notify timeout
                    )

                    # Read and speak
                    speak(f" hey buddy !!! You have a new email from {from_}. The subject is {subject}.")

        # Close and logout
        mail.close()
        mail.logout()
    except imaplib.IMAP4.error as e:
        print("Failed to authenticate:", e)

if __name__ == "__main__":
    initialize_last_seen_id()
    while True:
        check_email()
        # new email check
        time.sleep(60)
