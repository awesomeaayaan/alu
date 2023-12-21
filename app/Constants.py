import os

TIMEOUT = 30
SAVE_TO_LOCAL = True
MATCH_LIMIT = 2
NO_HIT_REPORT_COST = 250
HIT_REPORT_COST = 550

# Database part
base = os.path.join(os.environ['USERPROFILE'], "Documents", "RPA", "CIC_v2")
chromedriver = os.path.join(base, 'ChromeDriver', 'chromedriver.exe')

DB_LOCATION = f"{base}\\cic_database"
DB_NAME = "cic.db"

FILES_LOCATION = f"{base}\\files"
ATTACHMENT = f"{FILES_LOCATION}\\downloads"
# READ_FOLDER = f"{base}\\read_folder"
# TIME_TO_STOP = '18:00:00'

SC_SOURCE_PATH = f"{base}\\files\\screenshots\\"

FTP_PATH_URL = 'ftp://10.20.101.11/'
FTP_HOST = '10.20.101.11'
FTP_PORT = 21
