import pandas as pd
import os
import re
# dataframe for Credit info report
import Constants

LOAN_DETAIL_LIST = [
    'branch_name',
    'tel_fax_no',
    'loan_purpose',
    'loan_type',
    'loan_amount',
    'request_type',
    'self_bank_credit_report',
    'client_type',
    'account_type',
    'has_guarantor'
]
INDIVIDUAL_DETAIL_LIST = [
    'first_name',
    'gender',
    'nationality',
    'identifier_type',
    'issue_date',
    'identification_number',
    'issue_district',
    'date_of_birth',
    'father_name',
    'mother_name',
    'grandfather_name',
    'spouse_name'
]
INSTITUTION_DETAIL_LIST = [
    'company_name',
    'company_registration_name',
    'company_registration_date',
    'company_registration_organization',
    'PAN_number',
    'PAN_registration_date',
    'PAN_registration_district'
]
INDIVIDUAL_GUARANTOR_DETAIL_LIST = [
    'first_name',
    'gender',
    'nationality',
    'identifier_type',
    'issue_date',
    'identification_number',
    'issue_district',
    'date_of_birth',
    'father_name',
    'mother_name',
    'grandfather_name',
    'spouse_name'
]
INSTITUTION_GUARANTOR_DETAIL_LIST = [
    'company_name',
    'company_registration_name',
    'company_registration_date',
    'company_registration_organization',
    'PAN_number',
    'PAN_registration_date',
    'PAN_registration_district'
]

def excel_path():
    # .. or .
    file_list = os.listdir(Constants.ATTACHMENT)
    EXCEL_FILE = ''.join(
        file for file in file_list if file.lower().endswith(".xlsx"))
    return EXCEL_FILE 

def ticket_num():
    FILE_NAME = excel_path()
    df_credit = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    ticket = df_credit.loc[0, 'Ticket No']
    return ticket

def read_credit_info_report():
    dict = {}
    FILE_NAME = excel_path()
    df_credit = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    cic_branch = df_credit.loc[0, 'Branch']
    if pd.isna(cic_branch):
        cic_branch = ""
    dict[f'{LOAN_DETAIL_LIST[0]}'] = str(cic_branch)
    cic_tel_fax_no = df_credit.loc[0, 'Tel/Fax no']
    if pd.isna(cic_tel_fax_no):
        cic_tel_fax_no = ""
    dict[f'{LOAN_DETAIL_LIST[1]}'] = str(cic_tel_fax_no)
    cic_loan_purpose = df_credit.loc[0, 'Loan Purpose']
    if pd.isna(cic_loan_purpose):
        cic_loan_purpose = ""
    dict[f'{LOAN_DETAIL_LIST[2]}'] = str(cic_loan_purpose)
    cic_loan_type = df_credit.loc[0, 'Loan Type']
    if pd.isna(cic_loan_type):
        cic_loan_type = ""
    dict[f'{LOAN_DETAIL_LIST[3]}'] = str(cic_loan_type)
    cic_loan_amount = (df_credit.loc[0, 'Loan Amount'])[3:-3]
    if pd.isna(cic_loan_amount):
        cic_loan_amount = ""
    dict[f'{LOAN_DETAIL_LIST[4]}'] = str(cic_loan_amount)
    cic_type = df_credit.loc[0, 'Type']
    if pd.isna(cic_type):
        cic_type = ""
    dict[f'{LOAN_DETAIL_LIST[5]}'] = str(cic_type)
    cic_self_bank_report = df_credit.loc[0, 'Self Bank Credit Report']
    if pd.isna(cic_self_bank_report):
        cic_self_bank_report = ""
    dict[f'{LOAN_DETAIL_LIST[6]}'] = str(cic_self_bank_report)
    cic_client_type = df_credit.loc[0, 'Client Type']
    if pd.isna(cic_client_type):
        cic_client_type = ""
    dict[f'{LOAN_DETAIL_LIST[7]}'] = str(cic_client_type)
    cic_account_type = df_credit.loc[0, 'Account Type']
    if pd.isna(cic_account_type):
        cic_account_type = ""
    dict[f'{LOAN_DETAIL_LIST[8]}'] = str(cic_account_type)
    cic_has_guarantor = df_credit.loc[0, 'Has Guarantor']
    if pd.isna(cic_has_guarantor):
        cic_has_guarantor = ""
    dict[f'{LOAN_DETAIL_LIST[9]}'] = str(cic_has_guarantor)
    return dict

def read_individual_detail(data_num):
    FILE_NAME = excel_path()
    df = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    header = df.index[df['Ticket No'] == 'Individual Details - 1'].tolist()[0]
    df_individual = pd.read_excel(
        f"{Constants.ATTACHMENT}\\{FILE_NAME}", header=header+2, index_col=False)
    row = 0
    dict = {}
    dictlist = []
    for num in range(data_num):
        first_name = df_individual.loc[row, 'First Name']
        if pd.isna(first_name):
            first_name = ""
        mid_name = df_individual.loc[row, 'Middle Name']
        if pd.isna(mid_name):
            mid_name = ""
        last_name = df_individual.loc[row, 'Last Name']
        if pd.isna(last_name):
            last_name = ""
        cic_individual_name = f"{first_name} {mid_name} {last_name}"
        cic_individual_name = re.sub(' +', ' ', cic_individual_name)
        dict[f'{INDIVIDUAL_DETAIL_LIST[0]}'] = str(cic_individual_name)
        cic_gender = df_individual.loc[row, 'Gender']
        if pd.isna(cic_gender):
            cic_gender = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[1]}'] = str(cic_gender)
        cic_nationality = df_individual.loc[row, 'Nationality']
        if pd.isna(cic_nationality):
            cic_nationality = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[2]}'] = str(cic_nationality)
        cic_identifier_name = df_individual.loc[row, 'Identifier Type']
        if pd.isna(cic_identifier_name):
            cic_identifier_name = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[3]}'] = str(cic_identifier_name)
        cic_issue_date = str(df_individual.loc[row, 'Issue Date']).split()[0]
        if pd.isna(cic_issue_date):
            cic_issue_date = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[4]}'] = str(cic_issue_date)[:10]
        cic_identification_number = df_individual.loc[row,
                                                      'Identification Number']
        if pd.isna(cic_identification_number):
            cic_identification_number = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[5]}'] = str(cic_identification_number)
        cic_issue_district = df_individual.loc[row, 'Issued District']
        if pd.isna(cic_issue_district):
            cic_issue_district = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[6]}'] = str(cic_issue_district)
        cic_dob_bs = str(df_individual.loc[row, 'Date of Birth']).split()[0]
        if pd.isna(cic_dob_bs):
            cic_dob_bs = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[7]}'] = str(cic_dob_bs)[:10]
        cic_father_name = df_individual.loc[row, "Father's Name"]
        if pd.isna(cic_father_name):
            cic_father_name = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[8]}'] = str(cic_father_name)
        cic_mother_name = df_individual.loc[row, "Mother's Name"]
        if pd.isna(cic_mother_name):
            cic_mother_name = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[9]}'] = str(cic_mother_name)
        cic_grandfather_name = df_individual.loc[row, "Grand Father's Name"]
        if pd.isna(cic_grandfather_name):
            cic_grandfather_name = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[10]}'] = str(cic_grandfather_name)
        cic_spouse_name = df_individual.loc[row, "Spouse's Name"]
        if pd.isna(cic_spouse_name):
            cic_spouse_name = ""
        dict[f'{INDIVIDUAL_DETAIL_LIST[11]}'] = str(cic_spouse_name)
        row += 3
        newdict = dict.copy()
        dictlist.append(newdict)
    return dictlist

def read_individual_guarantor_detail(row_num):
    FILE_NAME = excel_path()
    df = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    header = df.index[df['Ticket No'] ==
                      'Individual Guarantor Details'].tolist()
    if header == []:
        return header
    else:
        header = df.index[df['Ticket No'] == 'Individual Guarantor Details'].tolist()[
            0]
    df_individual_guarantor = pd.read_excel(
        f"{Constants.ATTACHMENT}\\{FILE_NAME}", header=header+2, index_col=False)
    row = 0
    dict = {}
    dictlist = []
    for num in range(row_num):
        cic_individual_name = df_individual_guarantor.loc[row, 'First Name']
        if pd.isna(cic_individual_name):
            cic_individual_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[0]}'] = str(
            cic_individual_name)
        cic_gender = df_individual_guarantor.loc[row, 'Gender']
        if pd.isna(cic_gender):
            cic_gender = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[1]}'] = str(cic_gender)
        cic_nationality = df_individual_guarantor.loc[row, 'Nationality']
        if pd.isna(cic_nationality):
            cic_nationality = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[2]}'] = str(cic_nationality)
        cic_identifier_name = df_individual_guarantor.loc[row,
                                                          'Identifier Type']
        if pd.isna(cic_identifier_name):
            cic_identifier_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[3]}'] = str(
            cic_identifier_name)
        cic_issue_date = str(df_individual_guarantor.loc[row, 'Issue Date']).split()[0]
        if pd.isna(cic_issue_date):
            cic_issue_date = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[4]}'] = str(cic_issue_date)[
            :10]
        cic_identification_number = df_individual_guarantor.loc[row,
                                                                'Identification Number']
        if pd.isna(cic_identification_number):
            cic_identification_number = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[5]}'] = str(
            cic_identification_number)
        cic_issue_district = df_individual_guarantor.loc[row,
                                                         'Issued District']
        if pd.isna(cic_issue_district):
            cic_issue_district = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[6]}'] = str(
            cic_issue_district)
        cic_dob_bs = str(df_individual_guarantor.loc[row, 'Date of Birth']).split()[0]
        if pd.isna(cic_dob_bs):
            cic_dob_bs = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[7]}'] = str(cic_dob_bs)[:10]
        cic_father_name = df_individual_guarantor.loc[row, "Father's Name"]
        if pd.isna(cic_father_name):
            cic_father_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[8]}'] = str(cic_father_name)
        cic_mother_name = df_individual_guarantor.loc[row, "Mother's Name"]
        if pd.isna(cic_mother_name):
            cic_mother_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[9]}'] = str(cic_mother_name)
        cic_grandfather_name = df_individual_guarantor.loc[row,
                                                           "Grand Father's Name"]
        if pd.isna(cic_grandfather_name):
            cic_grandfather_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[10]}'] = str(
            cic_grandfather_name)
        cic_spouse_name = df_individual_guarantor.loc[row, "Spouse's Name"]
        if pd.isna(cic_spouse_name):
            cic_spouse_name = ""
        dict[f'{INDIVIDUAL_GUARANTOR_DETAIL_LIST[11]}'] = str(cic_spouse_name)
        row += 2
        newdict = dict.copy()
        dictlist.append(newdict)
    return dictlist

def read_institution_detail(row_num):
    FILE_NAME = excel_path()
    df = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    header = df.index[df['Ticket No'] == 'Institution Details - 1'].tolist()[0]
    df_institution = pd.read_excel(
        f"{Constants.ATTACHMENT}\\{FILE_NAME}", header=header+2, index_col=False)
    row = 0
    dict = {}
    dictlist = []
    for num in range(row_num):
        cic_name_of_company = df_institution.loc[row, 'Name of Company']
        if pd.isna(cic_name_of_company):
            cic_name_of_company = ""
        dict[f'{INSTITUTION_DETAIL_LIST[0]}'] = str(cic_name_of_company)
        cic_registration_number = df_institution.loc[row,
                                                     'Company Registration Name']
        if pd.isna(cic_registration_number):
            cic_registration_number = ""
        dict[f'{INSTITUTION_DETAIL_LIST[1]}'] = str(cic_registration_number)
        cic_registration_date = str(df_institution.loc[row,
                                                   'Company Registration Date']).split()[0]
        if pd.isna(cic_registration_date):
            cic_registration_number = ""
        dict[f'{INSTITUTION_DETAIL_LIST[2]}'] = str(cic_registration_date)
        cic_registration_organization = df_institution.loc[row,
                                                           'Company Registration Organization']
        if pd.isna(cic_registration_organization):
            cic_registration_number = ""
        dict[f'{INSTITUTION_DETAIL_LIST[3]}'] = str(
            cic_registration_organization)
        cic_pan_number = df_institution.loc[row, 'PAN Number']
        if pd.isna(cic_pan_number):
            cic_pan_number = ""
        dict[f'{INSTITUTION_DETAIL_LIST[4]}'] = str(cic_pan_number)
        cic_pan_registration_date = str(df_institution.loc[row,
                                                       'PAN Registration Date']).split()[0]
        if pd.isna(cic_pan_registration_date):
            cic_pan_registration_date = ""
        dict[f'{INSTITUTION_DETAIL_LIST[5]}'] = str(cic_pan_registration_date)
        cic_pan_registration_district = df_institution.loc[row,
                                                           'PAN Registration District']
        if pd.isna(cic_pan_registration_district):
            cic_pan_registration_district = ""
        dict[f'{INSTITUTION_DETAIL_LIST[6]}'] = str(
            cic_pan_registration_district)
        row += 3
        newdict = dict.copy()
        dictlist.append(newdict)
    return dictlist

def institution_checker():
    # returns true if there is ins guarantor
    FILE_NAME = excel_path()
    df = pd.read_excel(f'{Constants.ATTACHMENT}\\{FILE_NAME}')
    inst_check = df.index[df['Ticket No'] ==
                            "Institution Guarantor Details"].tolist()
    if inst_check != []:
        return True
    else:
        return False

def read_institution_guarantor_detail(row_num):
    FILE_NAME = excel_path()
    df = pd.read_excel(f"{Constants.ATTACHMENT}\\{FILE_NAME}", index_col=False)
    header = df.index[df['Ticket No'] ==
                      'Institution Guarantor Details'].tolist()
    if header == []:
        return header
    else:
        header = df.index[df['Ticket No'] == 'Institution Guarantor Details'].tolist()[
            0]
    df_institution_guarantor = pd.read_excel(
        f"{Constants.ATTACHMENT}\\{FILE_NAME}", header=header+2, index_col=False)
    row = 0
    dict = {}
    dictlist = []
    for num in range(row_num):
        cic_name_of_company = df_institution_guarantor.loc[row,
                                                           'Name of Company:']
        if pd.isna(cic_name_of_company):
            cic_name_of_company = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[0]}'] = str(
            cic_name_of_company)
        cic_registration_number = df_institution_guarantor.loc[row,
                                                               'Company Registration Name:']
        if pd.isna(cic_registration_number):
            cic_registration_number = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[1]}'] = str(
            cic_registration_number)
        cic_registration_date = str(df_institution_guarantor.loc[row,
                                                             'Company Registration Date:']).split()[0]
        if pd.isna(cic_registration_date):
            cic_registration_number = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[2]}'] = str(
            cic_registration_date)
        cic_registration_organization = df_institution_guarantor.loc[row,
                                                                     'Company Registration Organization:']
        if pd.isna(cic_registration_organization):
            cic_registration_number = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[3]}'] = str(
            cic_registration_organization)
        cic_pan_number = df_institution_guarantor.loc[row, 'PAN Number:']
        if pd.isna(cic_pan_number):
            cic_pan_number = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[4]}'] = str(cic_pan_number)
        cic_pan_registration_date = str(df_institution_guarantor.loc[row,
                                                                 'PAN Registration Date:']).split()[0]
        if pd.isna(cic_pan_registration_date):
            cic_pan_registration_date = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[5]}'] = str(
            cic_pan_registration_date)
        cic_pan_registration_district = df_institution_guarantor.loc[row,
                                                                     'PAN Registration District:']
        if pd.isna(cic_pan_registration_district):
            cic_pan_registration_district = ""
        dict[f'{INSTITUTION_GUARANTOR_DETAIL_LIST[6]}'] = str(
            cic_pan_registration_district)
        row += 2
        newdict = dict.copy()
        dictlist.append(newdict)
    return dictlist
