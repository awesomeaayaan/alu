import os
import shutil
from pathlib import Path
import Constants
from robot.libraries.BuiltIn import BuiltIn
import Constants


def del_excel():
    try:
        file_list = os.listdir(Constants.ATTACHMENT)
        EXCEL_FILE = ''.join(file for file in file_list if file.lower().endswith('.xlsx'))
        os.remove(f'{Constants.ATTACHMENT}\\{EXCEL_FILE}')
    
    except:
        raise Exception('Could not remove excel file.')

def path_returner_pdf(file_name = False):
    list_dir = os.listdir("./pdf_files")
    if file_name:
        return list_dir[0]
    else:  
        pdf_filepath = str(Path().absolute()) + "\\pdf_files\\" + f"{''.join(list_dir)}"
        return pdf_filepath

def local_file_path(cic_ref_no):
    try:
        list_local = os.listdir(f'{Constants.FILES_LOCATION}\\{str(cic_ref_no)}')
        if list_local:
            file_name = list_local[0]
            file_path = f'{Constants.FILES_LOCATION}\\{str(cic_ref_no)}\\{file_name}'
            return file_path
        else:
            return False
    except:
        return False

def file_name_returner(cic_ref_no):
    list_local = os.listdir(f'{Constants.FILES_LOCATION}\\{str(cic_ref_no)}')
    return list_local[0]
    

def cir_no_extractor(file_name: str):
    second_index = file_name.find('-', file_name.find('-') + 1)
    file_name = file_name[:second_index] + file_name[second_index:].replace('-', '/', 1)
    file_name = file_name[:-4]
    return file_name

def string_checker(str1:str, str2:str):
    str1 = ''.join(a for a in str1 if a.isalnum())
    str2 = ''.join(b for b in str2 if b.isalnum())
    if str1.lower() == str2.lower():
        return True
    return False

def number_checker(num1:str, num2:str):
    num1 = ''.join(a for a in num1 if a.isnumeric())
    num2 = ''.join(b for b in num2 if b.isnumeric())
    if num1 == num2:
        return True
    if num1 and num2:
        num1 = int(num1)
        num2 = int(num2)
        if num1 == num2:
            return True
    return False

def create_entity_tuple_data(type, ticket_no, data):
    if type == 'individual':
        entity = [
            ticket_no,
            str(data['first_name']),
            str(data['gender']),
            str(data['nationality']),
            str(data['identifier_type']),
            str(data['issue_date']),
            str(data['identification_number']),
            str(data['issue_district']),
            str(data['date_of_birth']),
            str(data['father_name']),
            str(data['mother_name']),
            str(data['grandfather_name']),
            str(data['spouse_name']),
            '',
            '',
            '',
            '',
            '',
            '',
            ''
        ]
    else:
        entity = [
            ticket_no,
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            str(data['company_name']),
            str(data['company_registration_name']),
            str(data['company_registration_date']),
            str(data['company_registration_organization']),
            str(data['PAN_number']),
            str(data['PAN_registration_date']),
            str(data['PAN_registration_district'])
        ]
    return entity

def copy_at_local(ticket_no, source):
    ticket_no = str(ticket_no)
    try:
        os.mkdir(Constants.FILES_LOCATION)
    except:
        pass
    try:
        os.mkdir(f'{Constants.FILES_LOCATION}\\{ticket_no}')
    except:
        pass
    try:
        destination = f'{Constants.FILES_LOCATION}\\{ticket_no}'
        shutil.copy2(src=source, dst=destination)
    except Exception as e:
        raise e
    
def copy_at_error(source):
    try:
        destination = f'{Constants.FILES_LOCATION}\\error_folder'
        shutil.copy2(src=source, dst=destination)
    except Exception as e:
        raise e
    
def copy_at_read(source):
    try:
        destination = f'{Constants.FILES_LOCATION}\\read_folder'
        shutil.copy2(src=source, dst=destination)
    except Exception as e:
        raise e
    
def validity_checker(dict):
    loan_data = dict['credit_info']
    client_data = dict['client_detail']
    ind_g_data = dict['individual_guar_detail']
    ins_g_data = dict['institution_guar_detail']
    if client_data != {}:
        for data in range(len(client_data)):
            dict = client_data[data]
            if str(loan_data['client_type']).strip().lower() == 'institution':
                if not institute_validity(dict):
                    print('institute invalid')
                    return False
            else:
                if not individual_validity(dict):
                    print('invalid individual')
                    return False
    else:
        return False
    if ind_g_data:
        for i in range(len(ind_g_data)):
            dict = ind_g_data[i]
            if not individual_validity(dict):
                print('individual invalid')
                return False
            
    if ins_g_data:
        for i in range(len(ins_g_data)):
            dict = ins_g_data[i]
            if not institute_validity(dict):
                print('institute invalid')
                return False
    return True

def institute_validity(dict):
    if dict['company_name'] and dict['PAN_number'] and dict['PAN_registration_date'] and dict['PAN_registration_district'] and dict['company_registration_date'] and dict['company_registration_name']:
        return True
    else:
        return False
    
def individual_validity(dict):
    if dict['first_name'] and dict['gender'] and dict['identifier_type'] and dict['identification_number'] and dict['father_name'] and dict['issue_date'] and dict['date_of_birth'] and dict['issue_district']:
        return True
    else:
        return False