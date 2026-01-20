
from PDF_Automation import GUI, ExcelHandler
from PDF_Automation import PDFAutomation
import os
from dotenv import load_dotenv



if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base_dir, ".env")
    png_path = os.path.join(base_dir, "zms_logo.png")
    icon_path = os.path.join(base_dir, "zms_logo.ico")

    load_dotenv(dotenv_path=env_path)
    excel_filename = os.getenv("EXCEL_FILE")

    # Excel Handler
    excel_handler = ExcelHandler(filename="boltworld.xlsx")
    # GUI component
    gui_handler = GUI(png=png_path, ico=icon_path, excel_handler=excel_handler)
    
    # PDF Automation object
    pdf_automation = PDFAutomation()
    pdf_automation.run(gui_handler=gui_handler)
