class Loan:

    CLIENT_TYPE_INDIVIDUAL = "Individual"
    CLIENT_TYPE_INSTITUTE = "Insititute"

    def __init__(self, loan_amount=0):
        # self.client_type = client_type
        self.loan_amount = loan_amount
        self.type_of_credit_facility = 'Other Product'