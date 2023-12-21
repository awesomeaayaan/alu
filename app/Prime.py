from qrlib.QRComponent import QRComponent
from RPA.Browser.Selenium import Selenium, ChromeOptions
import Constants
from qrlib.QREnv import QREnv
from robot.libraries.BuiltIn import BuiltIn
import Utils
import time
import datetime
from pathlib import Path
from FTP import FTP
import Extracting_data_excel
import os

class PrimeDash(QRComponent):

    def __init__(self, headless=False):
        super().__init__()
        self.browser = Selenium()
        # self.browser.auto_close = False
        self.headless = headless
        self.options = ChromeOptions()
        self.prefs = {
            "download.default_directory": f"{Constants.ATTACHMENT}"
        }
        self.options.add_experimental_option("prefs", self.prefs)
    
    def _load_primedash_vault(self):
        self.prime_cred = QREnv.VAULTS['prime_credentials']
    
    def wait_and_click(self, xpath, timeout):
        self.browser.wait_until_element_is_visible(xpath, timeout)
        self.browser.click_element(xpath)

    def wait_and_input_text(self, xpath,text, timeout):
        self.browser.wait_until_element_is_visible(xpath, timeout)
        if text:
            self.browser.input_text(xpath, text, clear=True)

    def wait_and_input_password(self, xpath,password, timeout):
        self.browser.wait_until_element_is_visible(xpath, timeout)
        self.browser.input_password(xpath, password)
           
    def login(self):
        # Login X-paths
        USERNAME_INPUT = '//input[@id="UserName"]'
        PASSWORD_INPUT = '//input[@id="Password"]'
        LOGIN_BTN = '//input[@value="Login"]'
        self.browser.auto_close = False
        
        # options = ChromeOptions()
        # options.add_argument('start-maximized')
        # executable_path = 'C:\\Users\\RPA\\Documents\\chromedriver\\chromedriver.exe'
        # self.browser.open_browser('http://10.20.101.27:1098/PCBLPortal', browser='chrome', executable_path=executable_path, options=options)

        self.browser.open_available_browser(
            url=self.prime_cred['url'], browser_selection='Chrome', preferences=self.prefs, download=False, maximized=True, alias='CIC_website')
        # get credentials from vault
        self.wait_and_input_text(xpath=USERNAME_INPUT,
                            text=self.prime_cred['username'], timeout=Constants.TIMEOUT)
        self.wait_and_input_password(xpath=PASSWORD_INPUT,
                                password=self.prime_cred['password'], timeout=Constants.TIMEOUT)
        self.wait_and_click(xpath=LOGIN_BTN, timeout=Constants.TIMEOUT)
        
        self.browser.wait_until_element_is_visible('//div[@class="row"]', Constants.TIMEOUT)
    
    def goto_dashboard(self, logger):
        PENDING_TASK_VIEW = '//h6[contains(text(), "Pending Task")]/parent::*/parent::*//label'
        # PENDING_TASK_VIEW = '//h6[contains(text(), "My Request")]/parent::*/parent::*//label'
        # PENDING_TASK_VIEW = '//h6[contains(text(), "My Involved Request")]/parent::*/parent::*//label'
        logger.info(f"Go to the portal dashboard")
        self.browser.go_to(self.prime_cred['dashboard'])
        self.wait_and_click(xpath=PENDING_TASK_VIEW, timeout=Constants.TIMEOUT)
        logger.info(f"Go to the portal dashboard successfull")

    def check_if_picked(self):
        PICKED_STATUS = '//td[@class="sorting_1" and (text() = "1")]/parent::*/td[text() = "Picked"]'
        self.browser.wait_until_element_is_visible('//th', timeout=Constants.TIMEOUT)
        list = self.browser.get_webelements(PICKED_STATUS)
        if len(list) > 0:
            return True
        return False

    def picking_process(self, logger):
        try:
            # Getting first pick
            VIEW_ARROW = f'//td[@class="sorting_1" and contains(text(), "1")]/parent::*/td/span'
            logger.info(f'Scrolling element till viewed')
            self.browser.scroll_element_into_view(VIEW_ARROW)
            self.wait_and_click(xpath=VIEW_ARROW, timeout=Constants.TIMEOUT)
            logger.info(f'Element: {VIEW_ARROW} is visible')
        except:
            time.sleep(30)
            return {}

        # Add picking later
        PICK_LOADING = '//img[@alt="Please wait Loading!!!"]/parent::*[@class="loading"]'
        PICK_BTN = '//button[contains(text(), "Pick")]'
        self.browser.wait_until_element_is_not_visible(PICK_LOADING, timeout=Constants.TIMEOUT)
        self.wait_and_click(xpath=PICK_BTN, timeout=Constants.TIMEOUT)

        LOADING = '//div[@class="loading"]'
        self.browser.wait_until_element_is_not_visible(LOADING, timeout=Constants.TIMEOUT)
        PREVIEW_DETAILS_BTN = '//a[@title="Preview Details"]'
        self.wait_and_click(xpath=PREVIEW_DETAILS_BTN, timeout=Constants.TIMEOUT)
       
        time.sleep(1)
        self.browser.switch_window("NEW")

        # Remove if wait and click is used
        self.browser.wait_until_element_is_visible(
            '//div[@class="card"]/div[@class="card-body"]/table[2]//b', Constants.TIMEOUT)
        EXPORT_BTN = '//button[@id="exporttable"]'
        self.wait_and_click(xpath=EXPORT_BTN, timeout=Constants.TIMEOUT)
        # check if excel file downloaded
        end_time = time.time() + 60
        while time.time() <= end_time:
            list_dir = os.listdir(f"{Constants.ATTACHMENT}")
            file_name = [
                file for file in list_dir if file.lower().endswith(".xlsx")]
            if file_name:
                break
            time.sleep(1)
        else:
            raise Exception("Excel File not downloaded.")

        ind_guar_count = self.get_individual_guarantor_count()

        if Extracting_data_excel.institution_checker():
            ins_guar_count = self.get_institution_guarantor_count()
        else:
            ins_guar_count = 0

        credit_info = Extracting_data_excel.read_credit_info_report()
        credit_type = credit_info['client_type']
        

        if credit_type == "Individual":
            ind_count = self.get_individual_count()
            self.browser.close_window()
            self.browser.switch_window('MAIN')
            return {
                'credit_info':credit_info,
                'client_detail': Extracting_data_excel.read_individual_detail(int(ind_count)),
                'individual_guar_detail': Extracting_data_excel.read_individual_guarantor_detail(int(ind_guar_count)),
                'institution_guar_detail': Extracting_data_excel.read_institution_guarantor_detail(int(ins_guar_count))
            }
        else:
            ins_count = self.get_institution_count()
            self.browser.close_window()
            self.browser.switch_window('MAIN')
            return {
                'credit_info':credit_info,
                'client_detail': Extracting_data_excel.read_institution_detail(int(ins_count)),
                'individual_guar_detail': Extracting_data_excel.read_individual_guarantor_detail(int(ind_guar_count)),
                'institution_guar_detail': Extracting_data_excel.read_institution_guarantor_detail(int(ins_guar_count))
            }

    def transfer_ticket(self):
        # Replacing user
        REPLACE_USER = '//nav//a[contains(text(), "Replace User")]'
        REPLACE_USER_ARROW = '//td[contains(text(), "CAD RPA")]/parent::*//span'
        REPLACE_USER_TEXT = '//h3[contains(text(), "Replace User")]'
        SELECT_USER = '//select[@id="ReplaceWithUserId"]'
        REPLACE_TO = 'Sumi Pudasani'
        COMMENT_AREA = '//textarea[@id="Comment"]'
        SAVE_BTN = '//button[@id="saveReplaceUser"]'
        ERROR_MSG = 'ERROR from RPA, check FTP if any reports are generated.'

        self.wait_and_click(xpath=REPLACE_USER, timeout=Constants.TIMEOUT)
        try:
            self.wait_and_click(xpath=REPLACE_USER_ARROW, timeout=Constants .TIMEOUT)
        except:
            self.browser.scroll_element_into_view('//a[@id="replaceUser_previous"]') # previous button
            self.wait_and_click(xpath=REPLACE_USER_ARROW, timeout=Constants.TIMEOUT)
        self.browser.wait_until_element_is_visible(
            REPLACE_USER_TEXT, timeout=Constants.TIMEOUT)
        self.browser.press_keys(None, 'SPACE')
        self.browser.select_from_list_by_label(SELECT_USER, REPLACE_TO)
        self.wait_and_input_text(xpath=COMMENT_AREA, text=ERROR_MSG, timeout=Constants.TIMEOUT)
        self.wait_and_click(xpath=SAVE_BTN, timeout=Constants.TIMEOUT)
        self.browser.wait_until_element_is_visible('//th')


    def get_individual_count(self):
        INDIVIDUAL_COUNT = '//table[not(@hidden="hidden")]//b[contains(text(), "Individual Details")]'
        individual_num = self.browser.get_element_count(INDIVIDUAL_COUNT)
        return individual_num

    def get_institution_count(self):
        INSTITUTION_COUNT = '//table[not(@hidden="hidden")]//b[contains(text(), "Institution Details")]'
        institution_num = self.browser.get_element_count(INSTITUTION_COUNT)
        return institution_num

    def get_individual_guarantor_count(self):
        ind_guar_num = self.browser.get_element_count(
            '//table[not(@hidden="hidden")]//b[contains(text(), "Individual Guarantor")]')
        return ind_guar_num

    def get_institution_guarantor_count(self):
        ins_guar_num = self.browser.get_element_count(
            '//table[not(@hidden="hidden")]//b[contains(text(), "Institution Guarantor")]')
        return ins_guar_num

    def after_completion(self, condition, FTP_PATH='', amt_charge=0):
        ACTION = '//select[@name="actonId"]'
        CHARGE_AMT = '//label[contains(text(), "CIC Charge Amount")]/parent::*/input'
        COMMENT = '//textarea[@id="appComment"]'
        FILE_CHOOSE = '//input[@id="upload"]'
        if condition:
            ACTION_VALUE = 'Completed'
            self.browser.select_from_list_by_label(ACTION, ACTION_VALUE)
            self.wait_and_input_text(xpath=CHARGE_AMT, text=amt_charge, timeout=Constants.TIMEOUT)
            # FTP_PATH = path_returner(ticket_no)    call from process
            with open('./output/filler.txt', 'w') as f:
                f.write(FTP_PATH)
            self.browser.choose_file(FILE_CHOOSE, f'{os.getcwd()}\\output\\filler.txt')
            self.wait_and_input_text(xpath=COMMENT, text=FTP_PATH,
                                timeout=Constants.TIMEOUT)
            SAVE = '//button[@id="saveAllClient" and not(@ng-disabled)]'
            self.browser.scroll_element_into_view('//h3[contains(text(), "Approval Logs")]')
            self.wait_and_click(xpath=SAVE, timeout=Constants.TIMEOUT)
            self.browser.wait_until_element_is_visible('//h6[contains(text(), "Pending Task")]', Constants.TIMEOUT)
        else:
            ACTION_VALUE = 'Reject' # Reject
            self.browser.select_from_list_by_label(ACTION, ACTION_VALUE)
            # wait_and_input_text(xpath=CHARGE_AMT, text=amt_charge, timeout=Constants.TIMEOUT)
            self.wait_and_input_text(
                xpath=COMMENT, text='The data in ticket has already been completed please check it again.', timeout=Constants.TIMEOUT)
            SAVE = '//button[@id="saveAllClient" and (@ng-disabled)]'
            self.browser.scroll_element_into_view('//h3[contains(text(), "Approval Logs")]')
            self.wait_and_click(xpath=SAVE, timeout=Constants.TIMEOUT)
            self.browser.wait_until_element_is_visible('//h6[contains(text(), "Pending Task")]', Constants.TIMEOUT)

    def logout(self):
        self.browser.close_all_browsers()