from qrlib.QRBot import QRBot
from app.CICProcess import CICProcess
import traceback
from robot.libraries.BuiltIn import BuiltIn


class Bot(QRBot):

    def __init__(self):
        super().__init__()
        self.process = CICProcess()

    def start(self):
        try:
            self.setup_platform_components()
            self.process.before_run()
            self.process.execute_run()
        except Exception as e:
            BuiltIn().log_to_console(e)
            BuiltIn().log_to_console(traceback.format_exc())
            raise Exception(e)

    def teardown(self):
        self.process.after_run()
