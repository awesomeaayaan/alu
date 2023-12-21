import sqlite3
import traceback
import Constants

from robot.libraries.BuiltIn import BuiltIn
from datetime import datetime, timedelta

from qrlib.QRComponent import QRComponent
from qrlib.QREnv import QREnv


class SQLite_db(QRComponent):
    def __init__(self) -> None:
        super().__init__()
        self.db_name = Constants.DB_LOCATION + '\\' + Constants.DB_NAME 
        self.conn = ''
        self.cur = ''

    def create_connection(self):
        conn = sqlite3.connect(self.db_name)
        self.conn = conn
        self.cur = conn.cursor()

    def close_connection(self, conn):
        return conn.close() 

    def create_cic_table(self):
        query = """CREATE TABLE IF NOT EXISTS CIC (
                Ticket_no VARCHAR(15) PRIMARY KEY,
                first_name VARCHAR(255),
                issue_date VARCHAR(25),
                identification_number VARCHAR(255),
                issue_district VARCHAR(255),
                father_name VARCHAR(255),
                company_name VARCHAR(255),
                company_registration_name VARCHAR(255),
                company_registration_date VARCHAR(255),
                PAN_number VARCHAR(255),
                Cost INTEGER,
                Status VARCHAR(15),
                file_name VARCHAR(255),
                save VARCHAR(10),
                Created_At DATETIME DEFAULT (datetime('now', 'localtime')),
                Updated_At DATETIME DEFAULT (datetime('now', 'localtime'))
            )"""

        query2 = f"""
        CREATE TRIGGER IF NOT EXISTS update_timestamp AFTER UPDATE ON CIC FOR EACH ROW
        BEGIN UPDATE CIC SET Updated_At = (datetime('now', 'localtime')) WHERE Ticket_no = NEW.Ticket_no; END;"""

        query3 = f"""
        CREATE TABLE IF NOT EXISTS CIC_Detail(
        Ticket_no VARCHAR(15),
        first_name VARCHAR(255),
        gender VARCHAR(15),
        nationality VARCHAR(15),
        identifier_type VARCHAR(15),
        issue_date VARCHAR(25),
        identification_number VARCHAR(255),
        issue_district VARCHAR(255),
        date_of_birth VARCHAR(255),
        father_name VARCHAR(255),
        mother_name VARCHAR(255),
        grandfather_name VARCHAR(255),
        spouse_name VARCHAR(255),
        company_name VARCHAR(255),
        company_registration_name VARCHAR(255),
        company_registration_date VARCHAR(255),
        company_registration_organization VARCHAR(255),
        PAN_number VARCHAR(255),
        PAN_registration_date VARCHAR(255),
        PAN_registration_district VARCHAR(255),
        status VARCHAR(255),
        Created_At DATE DEFAULT (date('now', 'localtime')),
        Updated_At DATE DEFAULT (date('now', 'localtime'))
        )
        """
        query4 = f"""
        CREATE TRIGGER IF NOT EXISTS update_timestamp AFTER UPDATE ON CIC_Detail FOR EACH ROW
        BEGIN UPDATE CIC_Detail SET Updated_At = (date('now', 'localtime')) WHERE Ticket_no = NEW.Ticket_no; END;"""
        with self.conn:
            self.conn.execute(query)
            self.conn.execute(query2)
            self.conn.execute(query3)
            self.conn.execute(query4)
    
    def data_exist_checker(self, ticket_no, file_name):
        with self.conn:
            data = self.cur.execute(f"select * from CIC WHERE Ticket_no='{ticket_no}' AND file_name='{file_name}'").fetchone()
        return data if data else False
    
    def check_whether_ticket_is_completed(self, type, dict: dict, check_duration = 0):
        print(dict)
        if check_duration:
            today = datetime.now()
            date_t = (today - timedelta(days=int(check_duration))).date()
            try:
                if str(type).lower() == 'individual':
                    if dict['first_name']:
                        check_query = f'''SELECT Updated_At, Status FROM CIC WHERE first_name="{dict["first_name"]}" and issue_date="{dict["issue_date"]}" and identification_number="{dict["identification_number"]}" and issue_district="{dict["issue_district"]}" and father_name="{dict["father_name"]}"'''
                else:
                        check_query = f'''SELECT Updated_At, Status FROM CIC WHERE company_name="{dict["company_name"]}" and company_registration_name="{dict["company_registration_name"]}" and company_registration_date="{dict["company_registration_date"]}" and PAN_number="{dict["PAN_number"]}"'''

                with self.conn:
                    returned_data = self.cur.execute(check_query).fetchone()
                    print(returned_data)
                    if returned_data:
                        updated_data = returned_data[0]
                        status = returned_data[1]
                        if updated_data >= str(date_t) and str(status).lower() == 'complete':
                            return True
                    return False
            except Exception as e:
                self.run_item.logger.error(traceback.format_exc())
                raise e
            


    def insert_data_table1(self, ticket_no, file_name, dict):
        """CIC data insertion"""
        insert_query = f'''INSERT INTO CIC
                (
                Ticket_no,
                Status,
                file_name,
                first_name,
                issue_date,
                identification_number,
                issue_district,
                father_name,
                company_name,
                company_registration_name,
                company_registration_date,
                PAN_number,
                save
                )
                VALUES ("{ticket_no}", "pending", "{file_name}", "{dict["first_name"]}", "{dict["issue_date"]}", "{dict["identification_number"]}", "{dict["issue_district"]}", "{dict["father_name"]}", "{dict["company_name"]}", "{dict["company_registration_name"]}", "{dict["company_registration_date"]}", "{dict["PAN_number"]}", "False");'''

        with self.conn:
            self.cur.execute(insert_query)

    def insert_data_table2(self, data):
        insert_query1 = """
            INSERT INTO CIC_Detail (
            Ticket_no,
            first_name,
            gender,
            nationality,
            identifier_type,
            issue_date,
            identification_number,
            issue_district,
            date_of_birth,
            father_name,
            mother_name,
            grandfather_name,
            spouse_name,
            company_name,
            company_registration_name,
            company_registration_date,
            company_registration_organization,
            PAN_number,
            PAN_registration_date,
            PAN_registration_district,
            status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending');
        """
        with self.conn:
            self.cur.execute(insert_query1, data)

    def update_status_table1(self, status, ticket_no):
        update_query = "UPDATE CIC SET Status=? WHERE Ticket_no=?"
        with self.conn:
            self.cur.execute(update_query, (status, ticket_no))

    def update_status_table_individual(self, status, ticket_no, name):
        update_query = "UPDATE CIC_Detail SET status=? WHERE Ticket_no=? AND first_name=?"
        with self.conn:
            self.cur.execute(update_query, (status, ticket_no, name))

    def update_status_table_institution(self, status, ticket_no, name):
        update_query = "UPDATE CIC_Detail SET status=? WHERE Ticket_no=? AND company_name=?"
        with self.conn:
            self.cur.execute(update_query, (status, ticket_no, name))

    def update_cir_cost(self, ticket_no, cost):
        cost_query = f"SELECT Cost FROM CIC WHERE Ticket_no='{ticket_no}'"
        with self.conn:
            i_cost = self.cur.execute(cost_query).fetchone()[0]
            BuiltIn().log_to_console(i_cost)
            BuiltIn().log_to_console(type(i_cost))
            if not i_cost:
                i_cost = 0
        BuiltIn().log_to_console(i_cost)
        total_cost = cost + i_cost
        update_query = "UPDATE CIC SET Ticket_no=?, Cost=? WHERE Ticket_no = ?"
        with self.conn:
            self.cur.execute(update_query, (ticket_no, total_cost, ticket_no))
        return total_cost
        
    def check_status_table1(self, ticket_no):
        check_query = f"SELECT Status FROM CIC WHERE Ticket_no='{ticket_no}'"
        try:
            with self.conn:
                status = self.cur.execute(check_query).fetchone()
                if status:
                    status = ''.join(status)
                    return status
                return ''
            
        except:
            raise Exception('No Ticket number')
        
    def check_status_table2(self, type, ticket_no, name):
        if type == 'individual':
            check_query = f"SELECT status FROM CIC_Detail WHERE Ticket_no=? AND first_name=?"
        else:
            check_query = f"SELECT status FROM CIC_Detail WHERE Ticket_no=? AND company_name=?"
            
        with self.conn:
            status = self.cur.execute(check_query, (ticket_no, name)).fetchone()
            if status:
                status = ''.join(status)
                return status
            return ''
    
    def fetch_report_cost(self, ticket_no):
        fetch_query = f"SELECT Cost FROM CIC WHERE Ticket_no='{ticket_no}'"
        with self.conn:
            cost = self.cur.execute(fetch_query).fetchone()
            return str(cost[0])
        
    def fetch_save_status(self, ticket_no):
        fetch_query = f"SELECT save FROM CIC WHERE Ticket_no='{ticket_no}'"
        with self.conn:
            cost = self.cur.execute(fetch_query).fetchone()
            return str(cost[0])
        
    def fetch_error_in_table2(self, ticket_no):
        fetch_query = f"SELECT first_name, company_name FROM CIC_Detail WHERE status='error' and Ticket_no='{ticket_no}'"
        with self.conn:
            data = self.cur.execute(fetch_query).fetchall()
            return data

    def update_save_flag(self, ticket_no):
        update = f"UPDATE CIC SET save='True' WHERE Ticket_no='{ticket_no}'"
        with self.conn:
            self.cur.execute(update)

    def fetch_flag_data(self, ticket_no):
        query = f"SELECT save FROM CIC WHERE Ticket_no='{ticket_no}'"
        with self.conn:
            flag = self.cur.execute(query).fetchone()
            if flag:
                return flag[0]
            else:
                return False

# b = {'company_name': 'MA RAJDEVI SURGICAL', 'company_registration_name': '17348/2079/080', 'company_registration_date': '2080-02-29', 'company_registration_organization': 'Office of Cottage & Small Industry lahan, Siraha', 'PAN_number': '620041951', 'PAN_registration_date': '2080-02-29', 'PAN_registration_district': 'Siraha'}
# a = SQLite_db()
# a.create_connection()
# print(a.fetch_flag_data('25092'))
# print(a.check_whether_ticket_is_completed(type='Institution', dict=b, check_duration=2))
# print(a.fetch_error_in_table2(24944))
# print(a.db_name)
# print(a.check_whether_ticket_is_completed(dict=b, check_duration=3))

        