import Constants
import os
from Constants import FTP_PATH_URL
from robot.libraries.BuiltIn import BuiltIn
from ftplib import FTP
from qrlib.QRComponent import QRComponent
from qrlib.QREnv import QREnv

class FTPcomp(QRComponent):
    def __init__(self) -> None:
        super().__init__()
        self.download_dir = f"{QREnv.BASE_DIR}\\pdf_files"
        self.ftp = FTP()

    def _load_ftp_vault(self):
        self.ftp_cred = QREnv.VAULTS['ftp_credentials']

    def connect_ftp(self):
        self._load_ftp_vault()
        username, password = self.ftp_cred['username'], self.ftp_cred['password']
        self.ftp.connect(self.ftp_cred['server'], int(self.ftp_cred['port']))
        self.ftp.login(username, password)

    def upload_file(self, ticket):
        self.connect_ftp()
        list_dir = os.listdir('./pdf_files')

        file = ''.join(a for a in list_dir if a.lower().endswith('.pdf'))
        self.ftp.cwd('CIC_Report')
        try:
            self.ftp.mkd(str(ticket))
            self.ftp.cwd(str(ticket))
        except:
            self.ftp.cwd(str(ticket))
        try:
            fh = open(os.curdir+'/pdf_files/'+file, 'rb') # reading in binary
            self.ftp.storbinary('STOR %s' %file, fh)
            #del this later
            # list_dirr = os.listdir('./output')
            # filee = ''.join(a for a in list_dirr if a.lower().endswith('.xlsx'))
            # fh = open(os.curdir+'/output/'+filee, 'rb') # reading in binary
            # self.ftp.storbinary('STOR %s' %filee, fh)
        except Exception as e:
            # SEND EMAIL WITH ATTACHMENT
            print(e)
            raise e
        finally:
            self.ftp.quit()

    def path_returner(self, ticket):
        self.connect_ftp()
        self.ftp.cwd('CIC_Report')
        self.ftp.cwd(str(ticket))
        PATH_FILES = self.ftp.pwd()
        PATH_FILES = Constants.FTP_PATH_URL + str(PATH_FILES) 
        self.ftp.quit()
        return PATH_FILES

    def test_upload(self, ticket):
        self.connect_ftp()
        list_dir = os.listdir('./output')
        file = ''.join(a for a in list_dir if a.lower().endswith('error_in.txt'))
        self.ftp.cwd('CIC_Report')
        try:
            self.ftp.mkd(str(ticket))
            self.ftp.cwd(str(ticket))
        except:
            self.ftp.cwd(str(ticket))
        try:
            fh = open(os.curdir+'/output/'+file, 'rb') # reading in binary
            self.ftp.storbinary('STOR %s' %file, fh)
        except Exception as e:
            # SEND EMAIL WITH ATTACHMENT
            print(e)
            raise e
        finally:
            self.ftp.quit()

    def check_data_exists(self, ticket_no):
        self.connect_ftp()
        try:
            list_local = os.listdir(f'C:/CIC/{str(ticket_no)}')
        except Exception as e:
            return
        try:
            self.ftp.cwd('CIC_Report')
            try:
                self.ftp.cwd(str(ticket_no))
            except:
                self.ftp.mkd(str(ticket_no))
                self.ftp.cwd(str(ticket_no))
            list_ftp = self.ftp.nlst()
            if len(list_local) != len(list_ftp):
                for file in list_local:
                    fh = open(f'C:/CIC/{str(ticket_no)}/{file}', 'rb') # reading in binary
                    self.ftp.storbinary('STOR %s' %file, fh)
        except Exception as e:
            raise e
        self.ftp.quit()