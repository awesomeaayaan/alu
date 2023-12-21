from RPA.Browser.Selenium import Selenium
from qrlib.QRComponent import QRComponent
from qrlib.QRRunItem import QRRunItem
from qrlib.QREnv import QREnv
import Constants

from Loan import Loan
from Individual import Individual
from Institution import Institution
import os
import time
from robot.libraries.BuiltIn import BuiltIn

class CIC(QRComponent):

    
    def __init__(self, loan: Loan, unique_id=None, headless=False) -> None:
        super().__init__()
        self.browser = Selenium()
        # self.browser.auto_close = False
        self.headless = headless
        self.loan = loan
        self.unique_id = unique_id
        self.is_success = False
        self.download_dir = f"{QREnv.BASE_DIR}\\pdf_files"
        self.cic_cred = QREnv.VAULTS['cic']
        
    def login(self):
        USERNAME_INPUT = '//input[@id="txtUsrID"]'
        PASSWORD_INPUT = '//input[@id="txtPwd"]'
        LOGIN_BTN = '//input[@name="btnLogin"]'
        LOGIN_SUCCESS = '//a[@class="drop"]'
        LOGIN_ERROR_MSG_XPATH = '//div[@id="ErrorMessage"]'
        ERROR_TEXT = [
            'User Id does not exists',
            'Incorrect password.',
            'User is already logged in or User does not logout properly. Please try to login after 1 minutes'
        ]

        try:
            # options = ChromeOptions()
            prefs = {
                "download.default_directory": self.download_dir
            }
            # options.add_experimental_option("prefs", prefs)
            if self.headless:
                self.browser.open_available_browser(
                    url=self.cic_cred["url"], preferences=prefs, download=False, headless=True, alias='CIC_website')
                self.browser.set_window_size(1920,1080)
            else:
                self.browser.open_available_browser(
                    url=self.cic_cred["url"], preferences=prefs, download=False, maximized=True, alias='CIC_website')
        except Exception as e:
            raise e

        try:
            self.browser.wait_until_element_is_visible(
                locator=USERNAME_INPUT, timeout=Constants.TIMEOUT)
            self.browser.input_text(locator=USERNAME_INPUT, text=self.cic_cred["username"])

            self.browser.wait_until_element_is_visible(
                locator=PASSWORD_INPUT, timeout=Constants.TIMEOUT)
            self.browser.input_password(locator=PASSWORD_INPUT, password=self.cic_cred["password"])

            self.browser.click_element(locator=LOGIN_BTN)

            self.browser.wait_until_element_is_visible(
                locator=LOGIN_SUCCESS, timeout=Constants.TIMEOUT)
        except Exception as e:
            error_text = self.browser.get_text(LOGIN_ERROR_MSG_XPATH)
            for msg in ERROR_TEXT:
                if msg in error_text or msg == error_text:
                    raise Exception(msg)
            raise e

    def logout(self):
        SESSION_DROPDOWN = '//a[@class="drop"]'
        LOGOUT_BTN = '//a[contains(text(), "Logout")]'

        try:
            self.browser.click_element(locator=SESSION_DROPDOWN)
            self.browser.wait_until_element_is_visible(
                locator=LOGOUT_BTN, timeout=Constants.TIMEOUT)
            self.browser.click_element(locator=LOGOUT_BTN)
            self.browser.close_all_browsers()
        except Exception as e:
            raise e

    def navigate_to(self, type):
        if (type == 'Consumer'):
            REPORT = '//a[@title="Consumer Report"]'
            SOLUTION = '//span[contains(text(), "Consumer Solution")]'
            SEARCH = '//a[contains(text(), "KSKL Conventional Consumer Scoring Product")]'
        elif (type == 'Commercial'):
            REPORT = '//a[@title="Commercial Report"]'
            SOLUTION = '//span[contains(text(), "Commercial Solution")]'
            SEARCH = '//a[contains(text(), "Commercial Comprehensive Hit Report")]'
        else:
            raise Exception(
                f"Navigation should be to only Consumer or Commercial. Found {type}")

        try:
            self.browser.wait_until_element_is_visible(
                locator=REPORT, timeout=Constants.TIMEOUT)
            self.browser.click_element_when_visible(locator=REPORT)
            self.browser.click_element_when_visible(locator=SOLUTION)
            self.browser.click_element_when_visible(locator=SEARCH)
        except Exception as e:
            raise e

    def fill_loan_data(self):
        try:
            CONSUMER_SELECT_CREDIT_TYPE = '//select[@id="ctl00_MasterContentPlaceHolder_gvUsageBand_ctl03_ddlNewCFType"]'
            CONSUMER_APP_AMOUNT = '//input[@id="ctl00_MasterContentPlaceHolder_gvUsageBand_ctl03_txtNewApplAmt"]'
            CONSUMER_APP_ADD_BTN = '//input[@id="ctl00_MasterContentPlaceHolder_gvUsageBand_ctl03_btnAdd"]'

            self.browser.select_from_list_by_label(
                CONSUMER_SELECT_CREDIT_TYPE, self.loan.type_of_credit_facility)
            self.browser.input_text(
                locator=CONSUMER_APP_AMOUNT, text=self.loan.loan_amount)
            self.browser.click_element_when_visible(
                locator=CONSUMER_APP_ADD_BTN)
            
        except Exception as e:
            raise e

    def fill_individual_data(self, individual: Individual, id_type="citizenship"):
        # Consumer Search Details
        CONSUMER_NAME = '//input[@id="ctl00_MasterContentPlaceHolder_txtName"]'
        CONSUMER_FATHER_NAME = '//input[@id="ctl00_MasterContentPlaceHolder_txtFathersName"]'
        CONSUMER_GENDER = '//select[@id="ctl00_MasterContentPlaceHolder_ddlGender"]'
        CONSUMER_IDENTIFIER_TYPE = '//select[@id="ddlIdSrcTyp"]'
        CONSUMER_IDENTIFIER_NUMBER = '//input[@id="ctl00_MasterContentPlaceHolder_txtIdNo"]'
        CONSUMER_IDENTIFIER_DATE = '//input[@id="ctl00_MasterContentPlaceHolder_txtIdIssueDateBS"]'
        CONSUMER_IDENTIFIER_DISTRICT = '//select[@id="ctl00_MasterContentPlaceHolder_ddlIdIssueDistrict"]'

        try:
            identifier_district = str(individual.identifier_district).capitalize()
            self.browser.wait_until_element_is_visible(
                locator=CONSUMER_NAME, timeout=Constants.TIMEOUT)
            
            self.browser.input_text(locator=CONSUMER_NAME, text=individual.name)
            self.browser.input_text(
                locator=CONSUMER_FATHER_NAME, text=individual.father_name)
            self.browser.select_from_list_by_label(
                CONSUMER_GENDER, individual.gender)
            
            if (id_type == "citizenship"):
                self.browser.select_from_list_by_label(
                    CONSUMER_IDENTIFIER_TYPE, individual.identifier_type)
                self.browser.input_text(
                    locator=CONSUMER_IDENTIFIER_NUMBER, text=individual.identifier_number)
            else:
                raise Exception(
                    f"Invalid id type. Expected citizenship. Found #{id_type}")
                
            # Filling these data is not important.
            # self.browser.input_text(CONSUMER_IDENTIFIER_DATE, individual.identifier_date_bs)
            # try:
            #     self.browser.select_from_list_by_label(CONSUMER_IDENTIFIER_DISTRICT, identifier_district)
            # except:
            #     pass
        except Exception as e:
            raise e

    def fill_institution_data(self, institution: Institution, id_type="pan"):
        # Commercial search details
        COMMERCIAL_NAME = '//input[@id="ctl00_MasterContentPlaceHolder_txtName"]'
        COMMERCIAL_IDENTIFIER_TYPE = '//select[@id="ctl00_MasterContentPlaceHolder_ddlIdSrcTyp"]'
        COMMERCIAL_IDENTIFICATION_NUMBER = '//input[@id="ctl00_MasterContentPlaceHolder_txtIdNo"]'
        COMMERCIAL_ISSUE_DATE = '//input[@id="ctl00_MasterContentPlaceHolder_txtIdIssueDateBS"]'
        COMMERCIAL_ISSUE_DISTRICT = '//select[@id="ddlIdIssueDistrict"]'
        
        try:
            if id_type:
                identification_num = institution.pan_number
                select_label = 'Pan Number'
            else:
                identification_num = institution.company_registration_number
                select_label = 'Company Registration Number'
            identifier_issue_district = str(institution.pan_registration_district).strip().capitalize()
            
            self.browser.wait_until_element_is_visible(
                locator=COMMERCIAL_NAME, timeout=Constants.TIMEOUT)
            self.browser.input_text(locator=COMMERCIAL_NAME, text=institution.name)

            self.browser.select_from_list_by_label(
                COMMERCIAL_IDENTIFIER_TYPE, select_label)
            self.browser.input_text(
                locator=COMMERCIAL_IDENTIFICATION_NUMBER, text=identification_num)
            
            # self.browser.input_text(locator=COMMERCIAL_ISSUE_DATE, text=institution.pan_registration_bs)
            # try:
            #     self.browser.select_from_list_by_label(COMMERCIAL_ISSUE_DISTRICT, identifier_issue_district)
            # except:
            #     pass
            
        except Exception as e:
            raise e

    def search(self):
        SEARCH_BTN = '//input[@id="ctl00_MasterContentPlaceHolder_btnSearch"]'
        # TODO: check loading logo
        
        # Although login button is visible initially,
        # it's not clickable due to loading mask.
        # Retry click every 1 second
        for i in range(Constants.TIMEOUT):
            try:
                self.browser.wait_until_element_is_visible(
                    locator=SEARCH_BTN, timeout=Constants.TIMEOUT)
                self.browser.click_element(locator=SEARCH_BTN)
                break
            except Exception as e:
                time.sleep(1)
                if i >= Constants.TIMEOUT-1:
                    raise e

    # def has_search_results(self):
    #     DATA_TABLE_GENERAL = '//ul[@id="ulGridResults"]'
    #     try:
    #         self.browser.wait_until_element_is_visible(DATA_TABLE_GENERAL, 60)
    #         return True
    #     except Exception as e:
    #         return False
        
    def has_search_results(self, type):
        if type == 'individual':
            DATA_TABLE_GENERAL = '//ul[@id="ulGridResults"]'
        else:
            DATA_TABLE_GENERAL = '//ul[@id="ulListView"]/ul[@class="grid"]'
        for i in range(24):
            try:
                self.browser.wait_until_element_is_visible(DATA_TABLE_GENERAL, 5)
                return True
            except Exception as e:
                window_handles = self.browser.get_window_handles()
                if len(window_handles) == 2:
                    return False
        raise Exception("Search result not visible.")

    def matches_for_individual(self, individual: Individual):
        DATA_TABLE_GENERAL = '//ul[@id="ulGridResults"]'
        self.browser.wait_until_element_is_visible(DATA_TABLE_GENERAL, 60)
        number_of_tables = self.browser.get_element_count(DATA_TABLE_GENERAL)
        
        for table in range(1, number_of_tables+1):
            cib = {}
            DATA_TABLE = f'//ul[@id="ulGridResults"][{table}]'
            # DATA_NAME = f'//ul[@id="ulGridResults"][{table}]//span[text()="Name:"]/following-sibling::b/span'
            # DATA_FATHER_NAME = f'//ul[@id="ulGridResults"][{table}]//span[text()="Father\'s Name:"]/following-sibling::span'
            # DATA_DOB_BS = f'//ul[@id="ulGridResults"][{table}]//span[text()="Date of Birth (BS) :"]/following-sibling::span'
            # DATA_GENDER = f'//ul[@id="ulGridResults"][{table}]//span[text()="Gender:"]/following-sibling::span'
            # DATA_NATIONALITY = f'//ul[@id="ulGridResults"][{table}]//span[text()="Nationality:"]/following-sibling::span'
            # DATA_IDENTIFIER_TYPE = f'//ul[@id="ulGridResults"][{table}]//span[text()="Identifier Type:"]/following-sibling::span'
            # DATA_IDENTIFIER_NUMBER = f'//ul[@id="ulGridResults"][{table}]//span[text()="Identification Number:"]/following-sibling::span'
            # DATA_IDENTIFIER_DATE_BS = f'//ul[@id="ulGridResults"][{table}]//span[text()="Identifier Issue / Expiry Date (BS):"]/following-sibling::span'
            # DATA_IDENTIFIER_DISTRICT = f'//ul[@id="ulGridResults"][{table}]//span[text()="Identifier Issue District:"]/following-sibling::span'
            # TODO: Confidence level
            cib["index"] = table
            data = self.browser.get_text(DATA_TABLE).split("\n")
            
            cib["name"] = data[0].split(':')[1].strip()
            cib["father_name"] = data[2].split(':')[1].strip()
            cib["dob_bs"] = data[4].split(':')[1].strip()
            cib["gender"] = data[5].split(':')[1].strip()
            cib["nationality"] = data[6].split(':')[1].strip()
            cib["identifier_type"] = data[7].split(':')[1].strip()
            cib["identifier_number"] = data[8].split(':')[1].strip()
            cib["identifier_date_bs"] = data[10].split(':')[1].strip()
            cib["identifier_district"] = data[11].split(':')[1].strip()
            cib["matches"] = []
            BuiltIn().log_to_console(f'{cib["name"]}')
            # cib["name"] = self.browser.get_text(DATA_NAME)
            # cib["father_name"] = self.browser.get_text(DATA_FATHER_NAME)
            # cib["dob_bs"] = self.browser.get_text(DATA_DOB_BS)
            # cib["gender"] = self.browser.get_text(DATA_GENDER)
            # cib["nationality"] = self.browser.get_text(DATA_NATIONALITY)
            # cib["identifier_type"] = self.browser.get_text(DATA_IDENTIFIER_TYPE)
            # cib["identifier_number"] = self.browser.get_text(DATA_IDENTIFIER_NUMBER)
            # cib["identifier_date_bs"] = self.browser.get_text(DATA_IDENTIFIER_DATE_BS)
            # cib["identifier_district"] = self.browser.get_text(DATA_IDENTIFIER_DISTRICT)
            # cib["matches"] = []
            # BuiltIn().log_to_console(cib)

            if (individual.is_match(cib)):
                individual.match_data.append(cib)
                if (len(individual.match_data) == Constants.MATCH_LIMIT):
                    break
                
        return individual.match_data

    def matches_for_institution(self, institution: Institution, type='pan'):
        DATA_TABLE_GENERAL = '//ul[@id="ulListView"]/ul[@class="grid"]'
        self.browser.wait_until_element_is_visible(DATA_TABLE_GENERAL, 60)
        number_of_tables = self.browser.get_element_count(DATA_TABLE_GENERAL)
        
        for table in range(1, number_of_tables+1):
            cib = {}
            DATA_TABLE = f'//ul[@class="grid"][{table}]'
            
            # DATA_NAME = f'//ul[@class="grid"][{table}]//span[text()="Name:"]/following-sibling::b/span'
            # DATA_REGISTRATION_DATE_BS = f'//ul[@class="grid"][{table}]//span[text()="Company Registration Date (BS) :"]/following-sibling::span'
            # DATA_IDENTIFIER_TYPE = f'//ul[@class="grid"][{table}]//span[text()="Identifier Type:"]/following-sibling::span'
            # DATA_IDENTIFIER_NUMBER = f'//ul[@class="grid"][{table}]//span[text()="Identification Number:"]/following-sibling::span'
            # DATA_IDENTIFIER_ISSUE_DATE_BS = f'//ul[@class="grid"][{table}]//span[text()="Identifier Issue Date (BS) :"]/following-sibling::span'
            # DATA_IDENTIFIER_ISSUE_DISTRICT = f'//ul[@class="grid"][{table}]//span[text()="Identifier Issue District / Authority:"]/following-sibling::span'
            # get values
            
            cib["index"] = table
            data = self.browser.get_text(DATA_TABLE).split("\n")
            
            cib["name"] = data[0].split(':')[1].strip()
            cib["reg_date_bs"] = data[3].split(':')[1].strip()
            cib["identifier_type"] = data[4].split(':')[1].strip()
            cib["identifier_number"] = data[5].split(':')[1].strip()
            cib["identifier_issue_date_bs"] = data[7].split(':')[1].strip()
            cib["identifier_issue_district"] = data[8].split(':')[1].strip()
            cib["matches"] = []
            BuiltIn().log_to_console(cib['name'])
            # cib["index"] = table
            # cib["name"] = self.browser.get_text(DATA_NAME)
            # cib["reg_date_bs"] = self.browser.get_text(DATA_REGISTRATION_DATE_BS)
            # cib["identifier_type"] = self.browser.get_text(DATA_IDENTIFIER_TYPE)
            # cib["identifier_number"] = self.browser.get_text(DATA_IDENTIFIER_NUMBER)
            # cib["identifier_issue_date_bs"] = self.browser.get_text(DATA_IDENTIFIER_ISSUE_DATE_BS)
            # cib["identifier_issue_district"] = self.browser.get_text( DATA_IDENTIFIER_ISSUE_DISTRICT) 
            # cib["matches"] = []

            if type == 'pan':    
                if (institution.is_match(cib)):
                    institution.match_data.append(cib)
                    # if (len(institution.match_data) == Constants.MATCH_LIMIT):
                    #     break
            else:
                if (institution.is_match(cib, type='not pan')):
                    institution.match_data.append(cib)
                    # if (len(institution.match_data) == Constants.MATCH_LIMIT):
                    #     break

        return institution.match_data

    def generate_report(self, has_search_result, index=-1):
        # index = -1 for no hit report
        try:
            if has_search_result:
                if (index == -1):
                    SUB_NOT_FOUND = '//span[@id="lblNoSubjectMatch"]'
                    self.browser.click_element(locator=SUB_NOT_FOUND)
                else:
                    GENERATE_REPORT = f'//ul[@class="grid"][{index}]//span[contains(text(), "Generate Report")]'
                    self.browser.click_element(GENERATE_REPORT)
        
            PDF = '//img[@id="pdf"]'
            # self.browser.is_alert_present(None, 'ACCEPT')
            BuiltIn().sleep(1)
            self.switch_to_popup_window()

            self.browser.wait_until_element_is_visible(locator=PDF, timeout=Constants.TIMEOUT)
            self.browser.click_element(locator=PDF)
        except Exception as e:
            raise e

        # check if excel file downloaded
        end_time = time.time() + 120
        while time.time() <= end_time:
            list_dir = os.listdir(self.download_dir)
            file_name = [
                file for file in list_dir if file.lower().endswith(".pdf")]
            if len(file_name) > 1:
                raise Exception("More than 1 PDF files found.")
            elif len(file_name) == 1:
                break
            time.sleep(1)
        else:
            raise Exception("PDF File not downloaded.")

        try:
            self.browser.close_window()
            self.browser.switch_window('MAIN')
        except Exception as e:
            raise e

    def switch_to_popup_window(self):
        for i in range(Constants.TIMEOUT):
            window_handles = self.browser.get_window_handles()
            if len(window_handles) == 2:
                self.browser.switch_window(window_handles[1])
                break
        else:
            raise Exception("Popup window not found")

    def change_parameter(self):
        CHANGE_PARAMETER = '//span[@id="LblChngSearchParams"]'
        self.browser.click_element(locator=CHANGE_PARAMETER)
        
    def del_pdf(self):
        try:
            file_list = os.listdir(self.download_dir)
            PDF_FILE = ''.join(
                file for file in file_list if file.lower().endswith(".pdf"))
            os.remove(f"./pdf_files/{PDF_FILE}")
        except:
            raise Exception('Could not remove PDF file.')
