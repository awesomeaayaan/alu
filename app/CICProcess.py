import time
import uuid
import pathlib
import traceback
import os

from robot.libraries.BuiltIn import BuiltIn

import Utils
import Constants
import Extracting_data_excel
from CIC import CIC
from Loan import Loan
from FTP import FTPcomp
from DB import SQLite_db
from Email import EmailReader
from Email import EmailSender
from app.Prime import PrimeDash
from Individual import Individual
from Institution import Institution
import Extracting_data_excel_email
import datetime

# Importing this library to remove warnings
import warnings


from qrlib.QREnv import QREnv
from qrlib.QRProcess import QRProcess
from qrlib.QRDecorators import run_item
from qrlib.QRRunItem import QRRunItem


class CICProcess(QRProcess):
    def __init__(self):
        super().__init__()

        # This removes warnings.
        warnings.filterwarnings('ignore')
        
        self.prime = PrimeDash()
        self.ftp = FTPcomp()
        self.db = SQLite_db()
        self.email = EmailReader()
        self.email_sender = EmailSender()
        
        self.register(self.prime)
        self.register(self.ftp)
        self.register(self.db)
        self.register(self.email)
        self.register(self.email_sender)
        self.run_item = None
        self.data = []

    @run_item(is_ticket=False)
    def before_run(self, *args, **kwargs):
        self.run_item = QRRunItem(logger_name='Prime Cic logger')

        self.notify(self.run_item)

        logger = self.run_item.logger
        try:
            logger.info('removing excel file if there is any')
            Utils.del_excel()
            logger.info('removed excel file successfully.')
        except:
            pass


        # self.ftp.upload_file('24090')
        # raise Exception('ball')
        try:
            logger.info("Creating required directory")
            pathlib.Path(Constants.base).mkdir(parents=True, exist_ok=True)
            pathlib.Path(Constants.FILES_LOCATION).mkdir(parents=True, exist_ok=True)
            logger.info("Creating attachment downloading folder")
            pathlib.Path(Constants.ATTACHMENT).mkdir(parents=True, exist_ok=True)
            # pathlib.Path(Constants.READ_FOLDER).mkdir(parents=True, exist_ok=True)
            pathlib.Path(Constants.DB_LOCATION).mkdir(parents=True, exist_ok=True)
            pathlib.Path('pdf_files').mkdir(parents=True, exist_ok=True)
            pathlib.Path(f'{Constants.FILES_LOCATION}\\read_folder').mkdir(parents=True, exist_ok=True)
            pathlib.Path(f'{Constants.FILES_LOCATION}\\error_folder').mkdir(parents=True, exist_ok=True)
            logger.info(f"Completed creating required directory")
        except Exception as e:
            logger.error(e)
            run_item.report_data = {
                "Remarks": "Error creating required folder"
            }
            raise e

        try:
            logger.info(f"Creating database connection")
            self.db.create_connection()
            logger.info(f"Creating database table if not exists")
            self.db.create_cic_table()
            logger.info(f'Created table succesful')
        except Exception as e:
            logger.error(e)
            run_item.report_data = {
                "Remarks": "Initializing database error"
            }
            raise e

        try:
            logger.info(f"Loading required vaults")
            self.prime._load_primedash_vault()
            self.ftp._load_ftp_vault()
            self.email._load_emailreader_vault()
            self.email_sender._load_email_sender_vault()
            self.flag = QREnv.VAULTS['flag']['flag']
            self.check_duration = QREnv.VAULTS['flag']['check_duration']
            logger.info(f"Getting flag from vault whether to process CIC from portal or using excel file")
        except Exception as e:
            logger.error(e)
            run_item.report_data = {
                "Remarks": "Loading necessary vault error"
            }
            raise e
        
        try:
            if self.flag != "email":
                logger.info("Bot will read ticket from pcbl portal")
                self.prime.login()
        except Exception as e:
            logger.error(e)
            run_item.report_data = {
                'Remarks': "Opening PCBL portal error"
            }
            self.prime.browser.screenshot(None, './output/error.png')
            raise e
        
        try:
            if self.flag ==  "email":
                logger.info(f"Connect to the IMAP email")
                self.email.imap_login()
        except Exception as e:
            logger.error(e) 
            run_item.report_data = {
                "Remarks": "Login to imap failure"
            }
            run_item.set_error()
            run_item.post()
            raise e

    @run_item(is_ticket=False, post_success=False)
    def before_run_item(self, *args, **kwargs):
        # Get run item created by decorator. Then notify to all components about new run item.
        run_item:  QRRunItem = kwargs['run_item']
        self.notify(run_item)

        logger = self.run_item.logger
        is_email = kwargs['is_email']

        if is_email:
            try:
                logger.info(f'Downloading atachement')
                self.file_name, sender = self.email.download_attachment()
                index = str(sender).index('<')
                self.sender_email = str(sender)[index+1:-1]
                BuiltIn().log_to_console(self.file_name)
                BuiltIn().log_to_console(self.sender_email)
                self.email_sender.recipients = self.sender_email
                self.email_sender.already_completed('asdf')
                BuiltIn().sleep(5000)
                if self.file_name:
                    pass
                else:
                    raise Exception('File not downloaded.')
            except Exception as e:
                logger.error(e)
                run_item.report_data = {
                    "Remarks": "Download attachment error"
                }
                raise e

    @run_item(is_ticket=True)
    def execute_run_item(self, *args, **kwargs):
        # Get run item created by decorator. Then notify to all components about new run item.
        run_item: QRRunItem() = kwargs["run_item"]
        self.notify(run_item)
        pass

    @run_item(is_ticket=False, post_success=False)
    def after_run_item(self, *args, **kwargs):
        run_item: QRRunItem() = kwargs["run_item"]
        self.notify(run_item)
        # Get run item created by decorator. Then notify to all components about new run item.
        pass

    @run_item(is_ticket=False, post_success=False)
    def after_run(self, *args, **kwargs):
        # Get run item created by decorator. Then notify to all components about new run item.
        self.run_item = QRRunItem(is_ticket=True, logger_name='Prime_CIC')
        self.notify(self.run_item)
        logger = self.run_item.logger
        self.prime.logout()
        
    def execute_run(self):
        # raise Exception('apple')
        
        asdf = 0  
        while True:
            run_item = QRRunItem(is_ticket=True, logger_name='Prime_CIC')
            self.notify(self.run_item)
            logger = self.run_item.logger

            runitem_data = {
                'ticket_no': ''
            }

            if self.flag == 'email':
                self.before_run_item(is_email=True)
                try:
                    logger.info(f"Read credit info report from file: {self.file_name}")
                    credit_info = Extracting_data_excel_email.read_credit_info_report(file_name=self.file_name)
                    credit_type = credit_info['client_type']
                    logger.info(f"Credit info report read successfully")
                except Exception as e:
                    logger.error(f'Error reading the data from file {self.file_name}')
                    logger.error(e)
                    run_item.report_data= {
                        "File Name": self.file_name,
                        "Remarks": "Could not read the excel file"
                    }
                    run_item.set_error()
                    run_item.post()
                    logger.info('Copying excel file to error folder.')
                    Utils.copy_at_error(f'{Constants.ATTACHMENT}\\{self.file_name}')
                    logger.info('Copied excel file to error folder.')
                    continue

                if credit_type == 'Individual':
                    try:
                        logger.info(f"Getting individual and institution guranteer details for individual")
                        dict = {
                            'credit_info': credit_info,
                            'client_detail': Extracting_data_excel_email.read_individual_detail(file_name=self.file_name, logger=logger),
                            'individual_guar_detail': Extracting_data_excel_email.read_individual_guarantor_detail(file_name=self.file_name, logger=logger),
                            'institution_guar_detail': Extracting_data_excel_email.read_institution_guarantor_detail(file_name=self.file_name, logger=logger)
                        }
                        logger.info(f"Getting individual and institution guranteer details for individual successfull")
                    except Exception as e:
                        logger.error(e)
                        run_item.report_data= {
                            "File Name": self.file_name,
                            "Remarks": "Could not read the individual data"
                        }
                        run_item.set_error()
                        run_item.post()
                        logger.info('Copying excel file to error folder.')
                        Utils.copy_at_error(f'{Constants.ATTACHMENT}\\{self.file_name}')
                        logger.info('Copied excel file to error folder.')
                        continue
                else:
                    try:
                        logger.info(f"Getting individual and institution guranteer details for institution")
                        dict = {
                            'credit_info': credit_info,
                            'client_detail': Extracting_data_excel_email.read_institution_detail(file_name=self.file_name, logger=logger),
                            'individual_guar_detail': Extracting_data_excel_email.read_individual_guarantor_detail(file_name=self.file_name, logger=logger),
                            'institution_guar_detail': Extracting_data_excel_email.read_institution_guarantor_detail(file_name=self.file_name, logger=logger)
                        }
                        logger.info(f"Getting individual and institution guranteer details for individual successfully")
                    except Exception as e:
                        logger.error(e)
                        run_item.report_data= {
                            "File Name": self.file_name,
                            "Remarks": "Could not read corporate data"
                        }
                        run_item.set_error()
                        run_item.post()
                        logger.info('Copying excel file to error folder.')
                        Utils.copy_at_error(f'{Constants.ATTACHMENT}\\{self.file_name}')
                        logger.info('Copied excel file to error folder.')
                        continue
                BuiltIn().log_to_console(dict)
                try:
                    logger.info('Copying excel file to read folder.')
                    Utils.copy_at_read(f'{self.email.download_dir}\\{self.file_name}')
                    logger.info('Copied excel file to read folder.')
                except Exception as e:
                    logger.error(e)
                    logger.error(f"Could not copy file to read folder")
                    run_item.report_data= {
                        "File Name": self.file_name,
                        "Remarks": "Could move file to read folder"
                    }
                    run_item.set_error()
                    run_item.post()
                    # continue

                try:
                    logger.info(f"Reading ticket number")
                    ticket_no = Extracting_data_excel_email.ticket_num(file_name=self.file_name)
                    logger.info(f"Reading ticket number successfully")
                except Exception as e:
                    logger.error(e)
                    run_item.report_data= {
                        "File Name": self.file_name,
                        "Remarks": "Could not read corporate data"
                    }
                    run_item.set_error()
                    run_item.post()
                    self.email_sender.auth_and_send(valid=False)
                    logger.info('Copying excel file to error folder.')
                    Utils.copy_at_error(f'{Constants.ATTACHMENT}\\{self.file_name}')
                    logger.info('Copied excel file to error folder.')
                    continue
                
                if not ticket_no:
                    ticket_no = ''.join([a for a in str(datetime.datetime.now()) if a.isalnum()])
                logger.info(f"Ticke number is: {ticket_no}")
            else:
                logger.info(f"Ticket is read from portal")
                try:
                    logger.info(f"Going to the dashboard in ticket portal")
                    error_type = 'portal'
                    self.prime.goto_dashboard(logger=logger)
                    logger.info(f"Starting picking process")
                    dict = self.prime.picking_process(logger=logger)
                    BuiltIn().log_to_console(dict)
                    logger.info(f"Starting picking process")
                    logger.info(dict)
                    if not dict:
                        logger.info('No ticket found.')
                        BuiltIn().log_to_console('No ticket found.')
                        return
                    logger.info(f"Extracting ticket number")
                    error_type = 'excel'
                    ticket_no = Extracting_data_excel.ticket_num()
                    logger.info(f"Extractied ticket number")
                    self.file_name = Extracting_data_excel.excel_path()
                    BuiltIn().log_to_console(ticket_no)
                except Exception as e:
                    logger.error(e)
                    run_item.report_data= {
                        "File Name": self.file_name,
                        "Remarks": "Could not read get data from PCBL portal"
                    }
                    run_item.set_error()
                    run_item.post()
                    if error_type == 'portal':
                        self.prime.browser.screenshot(None, './output/error.png')
                        pass

                    raise e
            
            try:
                logger.info('Checking the validity of the data in excel.')
                if Utils.validity_checker(dict):
                    logger.info('The data provided is valid.')
                    try:
                        logger.info('Checking whether the task is already completed or not.')
                        # status = self.db.check_status_table1(ticket_no=ticket_no)
                        check_date = self.db.check_whether_ticket_is_completed(type=dict['credit_info']['client_type'], dict=dict['client_detail'][0], check_duration=self.check_duration)
                        # check_date = False
                        # if status == 'complete': # duplicate
                        if check_date:
                            Utils.copy_at_error(source=f'{self.email.download_dir}\\{self.file_name}')
                            if self.flag == 'email':
                                logger.info(f'sending already completed status email')
                                self.email_sender.already_completed(ticket_no=ticket_no)
                                logger.info('sent email successfully.')
                            else:
                                if self.db.fetch_flag_data(ticket_no=ticket_no):
                                    self.prime.after_completion(condition=False)
                                else:
                                    self.prime.transfer_ticket()
                                    # self.email_sender.already_completed(ticket_no=ticket_no)

                            continue

                        logger.info('Checking whether the data exists in database or not.')
                        if not self.db.data_exist_checker(ticket_no=ticket_no, file_name=str(self.file_name)):
                            

                            logger.info('Inserting data into the database as the data does not exist.')
                            test_data = dict['client_detail'][0]
                            if str(dict['credit_info']['client_type']).lower() == 'individual':
                                insert_dict = {'first_name': test_data['first_name'], 'issue_date': test_data['issue_date'], 'issue_district': test_data['issue_district'], 'identification_number': test_data['identification_number'], 'father_name': test_data['father_name'], 'company_name': '', 'company_registration_name': '', 'company_registration_date': '', 'PAN_number': ''}
                            else:
                                insert_dict = {'first_name': '', 'issue_date': '', 'issue_district': '', 'identification_number': '', 'father_name': '', 'company_name': test_data['company_name'], 'company_registration_name': test_data['company_registration_name'], 'company_registration_date': test_data['company_registration_date'], 'PAN_number': test_data['PAN_number']}

                            self.db.insert_data_table1(ticket_no=ticket_no, file_name=str(self.file_name), dict=insert_dict)
                            client_data = dict['client_detail'][0]

                            if str(dict['credit_info']['client_type']).lower() == 'individual':
                                logger.info('Inserting all individual data and making institution data as an empty string.')
                                data = Utils.create_entity_tuple_data('individual',ticket_no, client_data)
                                # check the name in db and status to whether continue or skip.
                                self.db.insert_data_table2(data=data)
                                for ind in range(len(dict['individual_guar_detail'])):
                                    data = Utils.create_entity_tuple_data('individual', ticket_no, dict['individual_guar_detail'][ind])
                                    self.db.insert_data_table2(data=data)
                                for ins in range(len(dict['institution_guar_detail'])):
                                    data = Utils.create_entity_tuple_data('institution',ticket_no , dict['institution_guar_detail'][ins])
                                    self.db.insert_data_table2(data=tuple(data))
                            else:
                                logger.info('Inserting all institution data and making individual data as an empy string.')
                                data = Utils.create_entity_tuple_data('institution',ticket_no, client_data)
                                self.db.insert_data_table2(data=tuple(data))
                                for ins in range(len(dict['institution_guar_detail'])):
                                    data = Utils.create_entity_tuple_data('institution', ticket_no, dict['institution_guar_detail'][ins])
                                    self.db.insert_data_table2(data=tuple(data))
                                for ind in range(len(dict['individual_guar_detail'])):
                                    data = Utils.create_entity_tuple_data('individual', ticket_no, dict['individual_guar_detail'][ind])
                                    self.db.insert_data_table2(data=data)
                                
                            logger.info('Inserted data successfully into the database.')
                         
                    except Exception as e:
                        logger.error(e)
                        run_item.report_data= {
                            "File Name": self.file_name,
                            "Remarks": "Could not insert data into the database."
                        }
                        run_item.set_error()
                        run_item.post()
                        raise e
                    
                    loan_amount = dict['credit_info']['loan_amount']
                    loan_amount = ''.join([a for a in loan_amount if str(a).isnumeric()])
                    logger.info('Initializing loan object to pass details into cic.')
                    loan=Loan(loan_amount=int(loan_amount))
                    logger.info('Initialized loan object successfully.')
                    logger.info('Initializing cic object with loan object.')
                    cic = CIC(loan=loan)
                    logger.info('Initialized cic object successfully.')
                    cic.login()
                    report_cost = 0
                    if str(dict['credit_info']['client_type']).lower() == 'individual': # client type is individual

                        for ind in range(len(dict['individual_guar_detail'])+1):
                            BuiltIn().log_to_console('-------------------------->')
                            try:
                                if ind == len(dict['individual_guar_detail']) or len(dict['individual_guar_detail'])==0: 
                                    name = str(dict['client_detail'][0]['first_name'])
                                    if str(self.db.check_status_table2(type='individual', ticket_no=ticket_no, name=name)) != 'pending':
                                        continue # either it is completed or there is error
                                    individual = Individual(dict['client_detail'][0])
                                else:
                                    name = str(dict['individual_guar_detail'][ind]['first_name'])
                                    if str(self.db.check_status_table2(type='individual', ticket_no=ticket_no, name=name)) != 'pending':
                                        continue # either it is completed or there is error
                                    individual = Individual(dict['individual_guar_detail'][ind])
                                cic.navigate_to('Consumer')
                                cic.fill_individual_data(individual=individual)
                                time.sleep(1)
                                for i in range(30):
                                    try:
                                        cic.fill_loan_data()
                                        break
                                    except Exception:
                                        time.sleep(2)
                                        continue
                                cic.search()
                            except Exception as e:
                                raise e

                            if cic.has_search_results(type=str(dict['credit_info']['client_type']).lower()):
                                match_data = cic.matches_for_individual(individual=individual)
                                BuiltIn().log_to_console('match_data')
                                BuiltIn().log_to_console(match_data)
                                BuiltIn().log_to_console('match_data')

                                try: # runs when there is search result
                                    if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                        BuiltIn().log_to_console(len(match_data))
                                        for inde in range(len(match_data)):
                                            BuiltIn().log_to_console('random----------------------------------------random')
                                            index = match_data[inde]['index']
                                            BuiltIn().log_to_console(f'Generating {index}')
                                            cic.generate_report(has_search_result=True, index = index)
                                            BuiltIn().log_to_console(f'Generated {index}')
                                            pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                            BuiltIn().log_to_console(f'Copying on local in loop {index}')
                                            Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                            BuiltIn().log_to_console(f'FTP upload on loop {index}')
                                            self.ftp.upload_file(ticket=ticket_no)
                                            report_cost += 550
                                            BuiltIn().log_to_console(f'deleting on loop {index}')
                                            cic.del_pdf()
                                    else:
                                        index = -1
                                        cic.generate_report(has_search_result=True)
                                        pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'

                                        Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                        self.ftp.upload_file(ticket=ticket_no)
                                        report_cost += 250
                                        cic.del_pdf()

                                    self.db.update_status_table_individual(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_individual(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                            else: # runs when there is no search result 
                                try:
                                    cic.switch_to_popup_window()
                                    cic.generate_report(has_search_result=False)
                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                    self.ftp.upload_file(ticket=ticket_no)
                                    report_cost += 250
                                    self.db.update_status_table_individual(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_individual(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                        for ins in range(len(dict['institution_guar_detail'])): # institution guarantor
                            name = str(dict['institution_guar_detail'][ins]['company_name'])
                            if str(self.db.check_status_table2(type='institution', ticket_no=ticket_no, name=name)) != 'pending':
                                continue

                            institution = Institution(dict['institution_guar_detail'][ins])
                            cic.navigate_to('Commercial')
                            cic.fill_institution_data(institution=institution)
                            time.sleep(1)
                            for i in range(30):
                                try:
                                    cic.fill_loan_data()
                                    break
                                except Exception:
                                    time.sleep(2)
                                    continue
                            cic.search()

                            if cic.has_search_results(type='institution'.lower()):
                                match_data = cic.matches_for_institution(institution=institution)
                                try:
                                    if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                        for inde in range(len(match_data)):
                                            index = match_data[inde]['index']
                                            cic.generate_report(has_search_result=True, index = index)
                                            pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                            Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                            self.ftp.upload_file(ticket=ticket_no)
                                            report_cost += 550
                                            cic.del_pdf()
                                    else:
                                        cic.navigate_to('Commercial')
                                        cic.fill_institution_data(institution=institution, id_type='not pan')
                                        time.sleep(1)
                                        for i in range(30):
                                            try:
                                                cic.fill_loan_data()
                                                break
                                            except Exception:
                                                time.sleep(2)
                                                continue
                                        cic.search()

                                        if cic.has_search_results(type='institution'.lower()):
                                            match_data = cic.matches_for_institution(institution=institution, type='not pan')
                                            if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                                for inde in range(len(match_data)):
                                                    index = match_data[inde]['index']
                                                    cic.generate_report(has_search_result=True, index = index)
                                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                                    self.ftp.upload_file(ticket=ticket_no)
                                                    report_cost += 550
                                                    cic.del_pdf()
                                            else:
                                                index = -1
                                                cic.generate_report(has_search_result=True)
                                                pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                                Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                                self.ftp.upload_file(ticket=ticket_no)
                                                report_cost += 250
                                                cic.del_pdf()

                                    self.db.update_status_table_institution(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_institution(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                            else:
                                try:
                                    cic.switch_to_popup_window()
                                    cic.generate_report(has_search_result=False)
                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                    self.ftp.upload_file(ticket=ticket_no)
                                    report_cost += 250
                                    self.db.update_status_table_institution(status='complete', ticket_no=ticket_no, name=name)
                                    cic.del_pdf()

                                except Exception as e:
                                    self.db.update_status_table_institution(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                        # self.db.update_cir_cost(ticket_no=ticket_no, cost=report_cost)
                    else: # this part runs when the client is institution
                        for ins in range(len(dict['institution_guar_detail'])+1):
                            if ins == len(dict['institution_guar_detail']) or len(dict['institution_guar_detail'])==0:
                                name = str(dict['client_detail'][0]['company_name'])
                                BuiltIn().log_to_console(name)
                                if str(self.db.check_status_table2(type='institution', ticket_no=ticket_no, name=name)) != 'pending':
                                    continue
                                BuiltIn().log_to_console('making ins')
                                institution = Institution(dict['client_detail'][0])
                            else:
                                name = str(dict['institution_guar_detail'][ins]['company_name'])
                                if str(self.db.check_status_table2(type='institution', ticket_no=ticket_no, name=name)) != 'pending':
                                    continue
                                institution = Institution(dict['institution_guar_detail'][ins])
                            cic.navigate_to('Commercial')
                            cic.fill_institution_data(institution=institution)
                            time.sleep(1)
                            for i in range(30):
                                try:
                                    cic.fill_loan_data()
                                    break
                                except Exception:
                                    time.sleep(2)
                                    continue
                            cic.search()

                            if cic.has_search_results(type=str(dict['credit_info']['client_type']).lower()):
                                match_data = cic.matches_for_institution(institution=institution)
                                try:
                                    if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                        for inde in range(len(match_data)):
                                            index = match_data[inde]['index']
                                            cic.generate_report(has_search_result=True, index = index)
                                            pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                            Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                            self.ftp.upload_file(ticket=ticket_no)
                                            report_cost += 550
                                            cic.del_pdf()
                                    else:
                                        cic.navigate_to('Commercial')
                                        cic.fill_institution_data(institution=institution, id_type='not pan')
                                        time.sleep(1)
                                        for i in range(30):
                                            try:
                                                cic.fill_loan_data()
                                                break
                                            except Exception:
                                                time.sleep(2)
                                                continue
                                        cic.search()

                                        if cic.has_search_results(type='institution'.lower()):
                                            match_data = cic.matches_for_institution(institution=institution, type='not pan')
                                            if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                                for inde in range(len(match_data)):
                                                    index = match_data[inde]['index']
                                                    cic.generate_report(has_search_result=True, index = index)
                                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                                    self.ftp.upload_file(ticket=ticket_no)
                                                    report_cost += 550
                                                    cic.del_pdf()
                                            else:
                                                index = -1
                                                cic.generate_report(has_search_result=True)
                                                pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                                Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                                self.ftp.upload_file(ticket=ticket_no)
                                                report_cost += 250
                                                cic.del_pdf()

                                    self.db.update_status_table_institution(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_institution(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                            else:
                                try:
                                    cic.switch_to_popup_window()
                                    cic.generate_report(has_search_result=False)
                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                    self.ftp.upload_file(ticket=ticket_no)
                                    report_cost += 250
                                    cic.del_pdf()

                                    self.db.update_status_table_institution(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_institution(status='error', ticket_no=ticket_no, name=name)
                                    raise e

                        for ind in range(len(dict['individual_guar_detail'])):
                            name = str(dict['individual_guar_detail'][ind]['first_name'])
                            # if str(self.db.check_status_table2(type='individual', ticket_no=ticket_no, name=name)) != 'pending':
                            #     continue
                            individual = Individual(dict['individual_guar_detail'][ind])
                            cic.navigate_to('Consumer')
                            cic.fill_individual_data(individual=individual)
                            time.sleep(1)
                            for i in range(30):
                                try:
                                    cic.fill_loan_data()
                                    break
                                except Exception:
                                    time.sleep(2)
                                    continue
                            cic.search()

                            if cic.has_search_results(type=str(dict['credit_info']['client_type']).lower()):
                                match_data = cic.matches_for_individual(individual=individual)
                                try:
                                    if len(match_data)>=1 and len(match_data) <= 10: # 10 is limit
                                        for inde in range(len(match_data)):
                                            index = match_data[inde]['index']
                                            cic.generate_report(has_search_result=True, index = index)
                                            pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                            Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                            self.ftp.upload_file(ticket=ticket_no)
                                            report_cost += 550
                                            cic.del_pdf()
                                    else:
                                        index = -1
                                        cic.generate_report(has_search_result=True)
                                        pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                        Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                        self.ftp.upload_file(ticket=ticket_no)
                                        report_cost += 250
                                        cic.del_pdf()

                                    self.db.update_status_table_individual(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_individual(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                            else:
                                try:
                                    cic.switch_to_popup_window()
                                    cic.generate_report(has_search_result=False)
                                    pdf_filepath = f'{cic.download_dir}\\{"".join(os.listdir(cic.download_dir))}'
                                    Utils.copy_at_local(ticket_no=ticket_no, source=pdf_filepath)
                                    self.ftp.upload_file(ticket=ticket_no)
                                    report_cost += 250  
                                    cic.del_pdf()
                                    self.db.update_status_table_individual(status='complete', ticket_no=ticket_no, name=name)
                                except Exception as e:
                                    self.db.update_status_table_institution(status='error', ticket_no=ticket_no, name=name)
                                    raise e
                    cic.logout()
                    BuiltIn().log_to_console(f'{report_cost} ---------------------------->')
                    total_cost = self.db.update_cir_cost(ticket_no=ticket_no, cost=report_cost)

                    self.db.update_status_table1(status='complete', ticket_no=ticket_no)
                    FTP_path = self.ftp.path_returner(ticket=ticket_no)

                    error_list = self.db.fetch_error_in_table2(ticket_no=ticket_no)
                    if self.flag != 'email':
                        if not error_list:
                            self.prime.after_completion(condition=True, FTP_PATH=FTP_path, amt_charge=total_cost)
                        else:
                            self.email_sender.recipients = 'noreply.cic@pcbl.com.np'
                            self.email_sender.auth_and_send(ticket_no=ticket_no, valid=True, FTP_path=FTP_path, error_in=error_list)
                            self.prime.transfer_ticket()
                        self.db.update_save_flag(ticket_no)
                    else:
                        self.email_sender.auth_and_send(ticket_no=ticket_no, valid=True, FTP_path=FTP_path, error_in=error_list, cost=total_cost)

                    runitem_data['ticket_no'] = ticket_no
                    self.run_item.report_data = runitem_data
                    self.run_item.set_success()
                    # self.run_item.post()
                else:
                    logger.info('The data provided in invalid.')
                    # FTP_path = self.ftp.path_returner(ticket=ticket_no)
                    self.prime.after_completion(condition=False)
                    self.email_sender.auth_and_send(ticket_no=ticket_no, valid=False)
            except Exception as e:
                logger.error(e)
                run_item.report_data= {
                    "File Name": self.file_name,
                    "Remarks": "Error while generating loan report."
                }
                Utils.copy_at_error(source=f'{self.email.download_dir}\\{self.file_name}')
                try:
                    self.db.update_cir_cost(ticket_no=ticket_no, cost=report_cost)
                except:
                    pass
                run_item.set_error()
                # run_item.post()
                try:
                    cic.logout()
                except:
                    # Error was occured in PCBLPortal, CIC is not open.
                    pass
                raise e
            finally:
                Utils.del_excel()
            continue
