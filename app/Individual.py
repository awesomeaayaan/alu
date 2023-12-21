import Utils
from robot.libraries.BuiltIn import BuiltIn

class Individual:
    def __init__(self, individual_data, is_client=False):
        self.is_client = is_client

        self.name = individual_data['first_name']
        self.father_name = individual_data['father_name']
        self.identifier_number = individual_data['identification_number']
        self.identifier_date_bs = individual_data['issue_date']

        self.dob_bs = ''
        self.gender = individual_data['gender']
        self.nationality = individual_data['nationality']
        self.identifier_type = 'Citizenship Number' # individual_data['identifier_type']
        self.identifier_district = individual_data['issue_district']

        self.match_data = []
        self.cost = 0
        
    def is_match(self, cib):
        # comparison
        high_count = 0
        low_count = 0

        if (Utils.string_checker(self.name, cib['name'])):
            high_count += 1
            cib['matches'].append("name")
        
        if (Utils.string_checker(self.father_name, cib['father_name'])):
            high_count += 1
            cib['matches'].append("father_name")

        if (Utils.string_checker(self.identifier_number, cib['identifier_number'])):
            high_count += 1
            cib['matches'].append("identifier_number")

        if (Utils.number_checker(self.identifier_date_bs, cib['identifier_date_bs'])):
            high_count += 1
            cib['matches'].append("identifier_date_bs")

        if high_count >= 3:
            return True
        
        if high_count < 2:
            return False
   
        if Utils.number_checker(self.dob_bs, cib['dob_bs']):
            low_count += 1
            cib['matches'].append("dob_bs")

        if Utils.string_checker(self.gender, cib['gender']):
            low_count += 1
            cib['matches'].append("gender")

        if Utils.string_checker(self.nationality, cib['nationality']):
            low_count += 1
            cib['matches'].append("nationality")

        if Utils.string_checker(self.identifier_type, cib['identifier_type']):
            low_count += 1
            cib['matches'].append("identifier_type")

        if Utils.string_checker(self.identifier_district, cib['identifier_district']):
            low_count += 1
            cib['matches'].append("identifier_district")

        if (high_count == 2 and low_count >= 4):
            cib['matches'].append('name')
            return True
        else:
            return False