

import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import os
from PIL import Image, ImageTk

from dotenv import load_dotenv
from os import getenv


# loading enviornmental variables into our scope
load_dotenv(dotenv_path=".env")
excel_filename = getenv("EXCEL_FILE")


class GUI(tk.Tk):
    """Contains the frontend components of the PDFAutomation.
    """

    # private data members
    __base_dir = os.path.dirname(os.path.abspath(__file__))
    __png_path = os.path.join(__base_dir, "zms_logo.png")

    # constructor
    def __init__(self):
        super().__init__()
        
        # Absolute path for the icon (CRITICAL)
        icon_path = os.path.join(self.__base_dir, "zms_logo.ico")

        # Windows taskbar + task manager icon
        try:
            self.iconbitmap(icon_path)
        except Exception:
            pass  # fallback below

        # Extra fallback (Windows sometimes needs this)
        icon_img = Image.open(self.__png_path)
        icon_photo = ImageTk.PhotoImage(icon_img)
        self.iconphoto(True, icon_photo)

        # Window configuration
        self.title("PDF Order Extraction System")
        self.geometry("580x420")
        self.resizable(False, False)
        self.configure(bg="#f4f6f8")

         # Styling
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure("TLabel", font=("Segoe UI", 10), background="#f4f6f8")
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))

        self.pdf_path = tk.StringVar()

        self._load_logo()
        self._build_ui()

    # _load_logo
    def _load_logo(self):
        """Load and resize logo"""
        logo_image = Image.open(self.__png_path)
        logo_image = logo_image.resize((140, 140), Image.LANCZOS)
        self.logo = ImageTk.PhotoImage(logo_image)

    # _build_gui
    def _build_ui(self):
        # Logo
        logo_label = ttk.Label(self, image=self.logo, background="#f4f6f8")
        logo_label.pack(pady=(20, 10))

        # Header
        header = ttk.Label(
            self,
            text="PDF Order Automation",
            style="Header.TLabel"
        )
        header.pack(pady=(5, 5))

        # Description
        desc = ttk.Label(
            self,
            text="Extract order numbers, date and time\nfrom PDF files into Excel"
        )
        desc.pack(pady=(0, 20))

        # File selection
        file_frame = ttk.Frame(self)
        file_frame.pack(padx=20, fill="x")

        file_entry = ttk.Entry(
            file_frame,
            textvariable=self.pdf_path,
            state="readonly"
        )
        file_entry.pack(side="left", expand=True, fill="x", padx=(0, 10))

        browse_btn = ttk.Button(
            file_frame,
            text="Browse PDF",
            command=self.browse_file
        )
        browse_btn.pack(side="right")

        # Process button
        process_btn = ttk.Button(
            self,
            text="Process PDF",
            command=self.process_pdf
        )
        process_btn.pack(pady=25)

        # Status
        self.status_label = ttk.Label(
            self,
            text="No file selected",
            foreground="#555555"
        )
        self.status_label.pack()

        # Footer
        footer = ttk.Label(
            self,
            text="Output file: boltworld.xlsx",
            font=("Segoe UI", 9),
            foreground="#777777"
        )
        footer.pack(side="bottom", pady=10)

    # browse_file
    def browse_file(self):
        file = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf")]
        )

        if file:
            self.pdf_path.set(file)
            self.status_label.config(
                text=f"Selected: {os.path.basename(file)}",
                foreground="#0066cc"
            )

    # process_pdf
    def process_pdf(self):
        if not self.pdf_path.get():
            messagebox.showwarning(
                "No File Selected",
                "Please select a PDF file before processing."
            )
            return

        try:
            self.status_label.config(
                text="Processing PDF...",
                foreground="#333333"
            )
            self.update_idletasks()

            from .pdfa import PDFAutomation, ExcelHandler, PDFHandler
            # BAKCEND LINKAGE POINT
            pdf_automation = PDFAutomation()
            # Excel Handler
            excel_handler = ExcelHandler(excel_filename)
            # PDF Handler
            pdf_handler = PDFHandler(filename=self.pdf_path.get())
            status_code = pdf_automation.initialize(pdf_handler=pdf_handler, excel_handler=excel_handler)

            if status_code == 101:
                # it means that the user didn't closed the file, and still clicked OK
                # promts the user that changes havn't been saved in this case,
                self.pdf_path.set("")
                self.status_label.config(
                    text=f"Changes Not Saved! You need to process pdf again after closing Excel file ({excel_filename}) file.",
                    foreground='#821f04'
                )
            else:
                # Reset UI after success
                self.pdf_path.set("")
                self.status_label.config(
                    text="PDF processed successfully ✔",
                    foreground="#1a7f37"
                )

                messagebox.showinfo(
                    "Success",
                    "Order details have been written to boltworld.xlsx"
                )

        except Exception as e:
            self.status_label.config(
                text="An error occurred ✖",
                foreground="#cc0000"
            )
            messagebox.showerror(
                "Processing Error",
                f"Something went wrong:\n\n{str(e)}"
            )

    # prompt_error
    @staticmethod
    def prompt_error(message: dict):
        """Promts an error window to the user with the containing message.

        Args:
            message (str): A dict containting information regarding message.
        """
        root = tk.Tk()
        root.withdraw()  # hide main window
        response = messagebox.showwarning(title=message["title"], message=message["message"], icon=message["icon"])
        root.destroy()