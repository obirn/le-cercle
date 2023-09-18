import imaplib
from time import time
from email.message import Message

def main():
    print("Loading IMAP client...")
    M = imaplib.IMAP4_SSL(host="mail.gandi.net",port="993")
    M.login("serviceclient@lecercledesparfumeurscreateurs.com", os.getenv("IMAP_CLIENT_PASS"))
    print("IMAP client OK")
    
    M.select("Sent")

    new_message = Message()
    new_message["From"] = "hello@itsme.com"
    new_message["Subject"] = "My new mail."
    new_message.set_payload("This is my message.")

    M.append('Sent', '', imaplib.Time2Internaldate(time()), str(new_message).encode('utf-8'))

    M.close()
    M.logout()

if __name__ == "__main__":
    main()