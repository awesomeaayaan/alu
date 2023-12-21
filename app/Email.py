import email
import imaplib

from robot.libraries.BuiltIn import BuiltIn

from qrlib.QRComponent import QRComponent
from qrlib.QREnv import QREnv
from RPA.Email.ImapSmtp import ImapSmtp
from email.header import decode_header

import Constants

class EmailReader(QRComponent):
    def __init__(self):
        super().__init__()
        self.subject_to_search = ''
        self.download_dir = Constants.ATTACHMENT 

    def _load_emailreader_vault(self):
        self.imap_creds = QREnv.VAULTS['imap']

    def imap_login(self):
        logger = self.run_item.logger
        logger.info(f"Logging to IMAP")
        imap = imaplib.IMAP4(self.imap_creds['host'], int(self.imap_creds['port']))
        imap.login(self.imap_creds['account'], self.imap_creds['password'])
        logger.info(f"Logging to IMAP connect successfull")
        self.imap = imap
        
    def download_attachment(self):
        logger = self.run_item.logger
        imap = self.imap

        self.subject_to_search = self.imap_creds["subject"]
        search_criteria = f'UNSEEN SUBJECT "{self.subject_to_search}"'
        imap.select('INBOX')
        result, data = imap.uid('search', None, search_criteria)

        if result == 'OK':
            byte_mail_list = [num for num in data[0].split()]
            if not byte_mail_list:
                return False
            logger.info(byte_mail_list)

            _, fetched_data = imap.uid('fetch', byte_mail_list[0], '(RFC822)')
            msg = email.message_from_bytes(fetched_data[0][1])

            sender, _ = decode_header(msg['From'])[0]
            BuiltIn().log_to_console(sender)
            BuiltIn().log_to_console(_)

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue

                # Check if the part has a filename (attachment)
                if part.get('Content-Disposition') is None:
                    continue
                # Extract the attachment filename
                filename = part.get_filename()
                
                # Download the attachment
                if filename:
                    filepath = f"{self.download_dir}\\{filename}"
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
            return filename, sender                

    
class EmailSender(QRComponent):
    def __init__(self):
        super().__init__()
        self.subject = 'From RPA bot'
        self.body = ''
        self.recipients = None
        self.smtp_creds = None
        self.cc = None

    def _load_email_sender_vault(self):
        self.smtp_creds = QREnv.VAULTS['smtp']

    def auth_and_send(self, ticket_no, valid, FTP_path='', error_in: list=[], cost=0):
        self._load_email_sender_vault()
        self.cc = self.smtp_creds['cc']
        self.recipients = self.smtp_creds['recipients']
        mail = ImapSmtp(smtp_server=self.smtp_creds['server'], smtp_port=int(self.smtp_creds['port']))
        # mail.authorize(account=self.smtp_creds['account'], password=self.smtp_creds['password'], smtp_server=self.smtp_creds['server'], smtp_port=self.smtp_creds['port'])
        if valid and error_in:
            self.body = f'''Dear all,
The completed reports generated for ticket no {ticket_no} can be found in FTP path {FTP_path}

The list of names where error occured can be found below.
{error_in}

The total cost of generation of report is {cost}.

Thank you.

This mail is auto generated by the system. Please do not reply.
        '''
        elif valid and not error_in:
            self.body = f'''Dear all,

The required reports are generated for ticket no {ticket_no}.

The FTP path is {FTP_path}.

The total cost of generation of report is {cost}.

Thank you.

This mail is auto generated by the system. Please do not reply.
        ''' 
        
        else:
            self.body = f'''Dear all,

The data in attached file is not valid with ticket no {ticket_no}. Please fill all the data before sending.

Thank you.

This mail is auto generated by the system. Please do not reply.
        '''

        mail.send_message(
            sender=self.smtp_creds['account'],
            recipients=self.recipients,
            cc=self.cc,
            subject=self.subject,
            body=self.body,
        )

    def already_completed(self, ticket_no):
        self._load_email_sender_vault()
        mail = ImapSmtp(smtp_server=self.smtp_creds['server'], smtp_port=int(self.smtp_creds['port']))
        # mail.authorize(account=self.smtp_creds['account'], password=self.smtp_creds['password'], smtp_server=self.smtp_creds['server'], smtp_port=self.smtp_creds['port'])
        self.body = f'''Dear all,

The excel file with ticket no {ticket_no} has already been completed. Please check the ticket again.

Thank you.

This mail is auto generated by the system. Please do not reply.
        '''

        mail.send_message(
            sender=self.smtp_creds['account'],
            recipients=self.recipients,
            cc=self.cc,
            subject=self.subject,
            body=self.body,
        )