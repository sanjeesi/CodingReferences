import email
import imaplib
import os
import logging
import json

class ReadEmail():

    connection = None
    error = None

    def __init__(self):
        self.load_config()
        # self.createLogger()

    def load_config(self):
        try:
            ConfigFile = "config/PARAMS.config"
            global configdata
            global sftpuser
            global sftpip
            global mailboxHost
            global mailboxUser
            global mailboxPass
            global FilePath
            global Inbox
            global Closed

            with open(ConfigFile) as JF:
                configdata=json.load(JF)
            mailboxHost=configdata['MAILBOX_HOST']
            mailboxUser=configdata['MAILBOX_USER']
            mailboxPass=configdata['MAILBOX_PASS']
            FilePath=configdata['UC14_IN_PATH']
            Inbox = configdata['Inbox']
            Closed = configdata['Closed']
            sftpuser=configdata['sftpuser']
            sftpip=configdata['sftpip']
        except Exception as e:
            print('CONFIG load error: Cannot proceed'+str(e))
            logging.info('CONFIG load error: Cannot proceed'+str(e))
            logging.error('CONFIG load error: Cannot proceed')
            exit(1)	

    def createLogger(self):
        dateToday = date.today()
        ddmmyy = dateToday.strftime("%m%d%Y")
        LOG_FILENAME = configdata['HOME']+'log/FetchEmail_' + str(ddmmyy) +'.log'
        logging.basicConfig(filename=LOG_FILENAME, format='%(asctime)s:%(levelname)s: %(message)s', level = logging.INFO)
        logging.getLogger("paramiko").setLevel(logging.WARNING)

    def connect(self):
        imap = imaplib.IMAP4_SSL(mailboxHost)
        imap.login(mailboxUser, mailboxPass)
        imap.select(mailbox = Inbox, readonly=False) # so we can mark mails as read
        return imap

    def save_attachment(self, msg, download_folder):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)

        return: file path to attachment
        """
        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename().upper()
            excelIndex = filename.find('.CSV')
            if excelIndex > -1:
                filename = filename[:excelIndex]+' FRM-EML'+filename[excelIndex:]
                att_path = os.path.join(download_folder, filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path

    def fetch_unread_messages(self):
        """
        Retrieve messages
        """
        imap = self.connect()
        emails = []
        (result, messages) = imap.search(None, 'All')
        if result == "OK":
            for message in str(messages[0], 'utf-8').split(' '):
                try: 
                    ret, data = imap.fetch(message,'(RFC822)')
                except:
                    print ("No new emails to read.")
                    imap.logout()
                    exit()

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)

                #MOVE MESSAGE TO Closed FOLDER
            for idx in range(1, len(emails)+1):
                result = imap.copy(str(idx), Closed)
                if result[0] == 'OK':
                    response, data = imap.store(str(idx), '+FLAGS', '(\Deleted)')
            print(imap.expunge())
            return emails
        imap.close()
        imap.logout()
        self.error = "Failed to retreive emails."
        return emails

    # def parse_email_address(self, email_address):
    #     """
    #     Helper function to parse out the email address from the message

    #     return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
    #     """
    #     return email.utils.parseaddr(email_address)

def main():
    obj = ReadEmail()
    emails = obj.fetch_unread_messages()
    for email in emails:
        print(obj.save_attachment(email, FilePath))
    #print(obj.indexNoList, obj.TADList, obj.invoiceAmt, obj.paymentAmt, obj.paymentDate)


if __name__== "__main__":
    main()
