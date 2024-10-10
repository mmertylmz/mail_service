import win32serviceutil
import win32service
import win32event
import time
import logging
import os
import imaplib
import email
from email.header import decode_header
from configparser import ConfigParser

class EmailAttachmentService(win32serviceutil.ServiceFramework):
    _svc_name_ = "EmailAttachmentService"
    _svc_display_name_ = "Email Attachment Fetcher Service"
    _svc_description_ = "A service that fetches email attachments every minute."
    
    def __init__(self, args):
        super().__init__(args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logger = self._setup_logger()
        self.running = True
        self.config = ConfigParser()
        self.config.read(os.path.join(self.base_dir, 'config.ini'))
        self.username = self.config.get('GMAIL', 'username')
        self.password = self.config.get('GMAIL', 'password')
        self.imap_server = self.config.get('GMAIL', 'imap_server')
        self.attachment_dir = os.path.join(self.base_dir, "attachments")

    def _setup_logger(self):
        logger = logging.getLogger('EmailAttachmentService')
        log_file_path = os.path.join(self.base_dir, 'log', 'service.log')
        handler = logging.FileHandler(log_file_path , encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        return logger

    def SvcStop(self):
        self.logger.info('Service is stopping...')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.logger.info('Service is starting...')
        while self.running:
            try:
                self.fetch_and_save_attachments()
            except Exception as e:
                self.logger.error(f"Error: {str(e)}")
            time.sleep(60) 

    def fetch_and_save_attachments(self):
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.username, self.password)
        except imaplib.IMAP4.error as e:
            self.logger.error(f"IMAP login failed: {e}")
            return

        mail.select("INBOX")
        status, messages = mail.search(None, 'UNSEEN')
        if status != 'OK':
            self.logger.error("Could not find email.")
            return

        email_ids = messages[0].split()
        for email_id in email_ids[-1:]:
            result, message_data = mail.fetch(email_id, '(RFC822)')
            if result != 'OK':
                self.logger.error(f"{email_id} could not fetch email.")
                continue
            
            raw_email = message_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            # Get Sender email
            from_header = email_message["From"]
            from_decoded = decode_header(from_header)[0][0]
            if isinstance(from_decoded, bytes):
                try:
                    from_decoded = from_decoded.decode()
                except UnicodeDecodeError:
                    from_decoded = from_decoded.decode('utf-8', errors='replace')
            self.logger.info(f"Sender: {from_decoded}")

            self.save_attachments(email_message)
        
        mail.logout()

    def save_attachments(self, msg):
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if filename:
                    decoded_header = decode_header(filename)
                    filename, charset = decoded_header[0]

                    if isinstance(filename, bytes):
                        charset = charset or 'utf-8'
                        filename = filename.decode(charset, errors='replace')

                    filename = os.path.basename(filename)
                    filepath = os.path.join(self.attachment_dir, filename)

                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    self.logger.info(f"{filename} saved at {filepath}.")

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(EmailAttachmentService)
