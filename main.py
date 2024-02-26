import logging
import os
from exchangelib import Credentials, Account, Configuration, DELEGATE, Q
from attachment_handler import AttachmentHandler
import sys

#syslog
log_file = os.path.join(os.getcwd(), 'email_processor.log')
logging.basicConfig(filename=log_file, level=logging.INFO)
logger = logging.getLogger()

#ouputs console
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)


class EmailProcessor:
    def __init__(self, host, email_address, password):
        self.host = host
        self.email_address = email_address
        self.password = password
        self.account = None

    def connect_to_exchange_server(self):
        try:
            if self.host == '':
                credentials = Credentials(self.email_address, self.password)
                self.account = Account(self.email_address, credentials=credentials, autodiscover=True, access_type=DELEGATE)
            else:
                creds = Credentials(username=self.email_address, password=self.password)
                config = Configuration(server=self.host, credentials=creds)
                self.account = Account(primary_smtp_address=self.email_address, autodiscover=False, config=config, access_type=DELEGATE)
              
        except Exception as e:
            logger.error(f"Failed to connect to Exchange server: {e}")

    def process_emails(self, is_read, subject, count):
        try:
            if not self.account:
                raise ValueError("Must connect to Exchange server first")

            if is_read:
                logger.info('onlyRead')
                if count != 0:
                    mail = self.account.inbox.filter(Q(subject__icontains=subject) & Q(is_read=False)).order_by('-datetime_received')[:count]
                else:
                    mail = self.account.inbox.filter(Q(subject__icontains=subject) & Q(is_read=False))
            else:
                if count !=0:
                    mail = self.account.inbox.filter(subject__icontains=subject).order_by('-datetime_received')[:count]
                else:
                    mail = self.account.inbox.filter(subject__icontains=subject)

            for item in mail:
                if item.attachments:
                    attachment_handler = AttachmentHandler()
                    attachment_handler.process_attachments(item.attachments)

        except Exception as e:
            logger.error(f"Error processing emails: {e}")

# 使用示例
if __name__ == "__main__":
    host = '' # ''不指定host
    address = '@dynasafe.com.tw'
    password = ''

    processor = EmailProcessor(host, address, password)
    processor.connect_to_exchange_server()
    processor.process_emails(False, '電子', 0) #未讀郵件,'subject'