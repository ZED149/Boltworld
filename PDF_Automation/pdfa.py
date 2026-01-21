

# This is the core Python file for PDF Automation class

# IMPORTS !!!
from .handlers.pdf_handler import PDFHandler
from .handlers.excel_handler import ExcelHandler
from .handlers.gui_handler import GUI

class PDFAutomation:
    """This class is reponsible for handling and organizing PDF Automation tasks of any type.
    """

    # private data members
    __pypdf = None                                  # Object to contain pdf file, delete it afterwards

    # constructor
    def __init__(sedf):
        pass

    # initialize
    def initialize(self, pdf_handler: PDFHandler, excel_handler: ExcelHandler):
        """Initializes the automation task for this instance.

        Args:
            filename (str): Name of the PDF file to work on. It needs to be orders file not any other file.

        Returns:
            _type_: _description_
        """

        # Perforimg Validations!

        # validating pdf_handler
        assert type(pdf_handler) == PDFHandler, "pdf_handler needs to be PDFHandler"
        assert pdf_handler != None, "pdf_handler cannot be none"

        # validating excel_handler
        assert type(excel_handler) == ExcelHandler, "excel_handler needs to be ExcelHandler"
        assert excel_handler != None, "excel_handler cannot be none"

        # opening the pdf file
        pdf_handler.open()

        # fetching order details from pdf
        order_details = pdf_handler.fetch_order_details(o_type='web')

        # writing the fetched order details on the Excel file
        wb = excel_handler.open_file(headers=["ORDER_DETAILS", "DATE", "TIME", "USER"], create_file=True)

        # writing data on the excel_file
        excel_handler.write(wb.active, data=order_details)

        # saving the workbook
        code = excel_handler.save(wb)

        return code

    # run
    def run(self, gui_handler: GUI):
        """Takes the GUI Handler of the program and starts the main loop.
        """
        gui_handler.mainloop()