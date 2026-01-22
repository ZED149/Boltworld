

# This is the main of the PDF Automation program

# loading enviornemtnal variables into our scope
from dotenv import load_dotenv
import os

load_dotenv(dotenv_path=".env")
START_INDEX = int(os.getenv("START_INDEX"))
EXCEL_FILENAME = os.getenv("EXCEL_FILENAME")
LOGGER_NAME = os.getenv("LOGGER_NAME")
LOGGER_DIR = os.getenv("LOGGER_DIR")

# IMPORTS!!!
from PDF_Automation import PDFAutomation, Logging
from PDF_Automation.handlers import ExcelHandler

# MAIN
base_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(base_dir, ".env")
png_path = os.path.join(base_dir, "zms_logo.png")
icon_path = os.path.join(base_dir, "zms_logo.ico")
logger_path = os.path.join(base_dir, LOGGER_DIR)

# Logger
logger = Logging(logger_name="salman", logger_directory=logger_path)
logger.verbose = True
logger.log_starting_details_to_file()

# Excel Handler
excel_handler = ExcelHandler(logger=logger, filename=EXCEL_FILENAME)
wb = excel_handler.open_file(headers=["ORDER_DETAILS", "DATE", "TIME", "PRINTED_BY"], create_file=True)
indexes = excel_handler.indexing(workbook=wb, start_index=START_INDEX)
ws = wb.active
excel_handler.write(worksheet=ws, data=indexes)
excel_handler.save(workbook=wb)
print("Everything OK")