import Utils

class Institution:

    def __init__(self, institution_data, is_client=False):
        self.is_client = is_client

        self.name = institution_data['company_name']
        self.pan_number = institution_data['PAN_number']
        self.pan_registration_bs = institution_data['PAN_registration_date']
        self.pan_registration_district = institution_data['PAN_registration_district']
        self.company_registration_number = institution_data['company_registration_name']
        self.company_registration_bs = institution_data['company_registration_date']

        self.match_data = []
        self.cost = 0

    def is_match(self, cib, type='pan'):
        # comparison
        if type=='pan':
            if (Utils.number_checker(self.pan_number, cib["identifier_number"])):
                cib["matches"].append("pan_number")
                return True
            else:
                return False
        else:
            if (Utils.number_checker(self.company_registration_number, cib["identifier_number"])):
                cib["matches"].append("com_reg_number")
                return True
            else:
                return False