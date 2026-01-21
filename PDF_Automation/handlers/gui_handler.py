import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import os
from PIL import Image, ImageTk
import openpyxl
from datetime import datetime

class GUI(tk.Tk):
    """Contains the frontend components of the PDFAutomation.
    """
    excel_handler = None

    # constructor
    def __init__(self, png: str, ico: str, excel_handler):
        super().__init__()
        
        # Windows taskbar + task manager icon
        try:
            self.iconbitmap(ico)
        except Exception:
            pass  # fallback below
        
        self.excel_handler = excel_handler

        # Extra fallback (Windows sometimes needs this)
        icon_img = Image.open(png)
        icon_photo = ImageTk.PhotoImage(icon_img)
        self.iconphoto(True, icon_photo)

        # Window configuration
        self.title("PDF Order Extraction System")
        self.geometry("650x620")
        self.resizable(False, False)
        self.configure(bg="#f4f6f8")

         # Styling
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure("TLabel", font=("Segoe UI", 10), background="#f4f6f8")
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        self.style.configure("SubHeader.TLabel", font=("Segoe UI", 11, "bold"))

        self.pdf_path = tk.StringVar()

        self._load_logo(logo_path=png)
        self._build_ui()

    # _load_logo
    def _load_logo(self, logo_path: str):
        """Load and resize logo"""
        logo_image = Image.open(logo_path)
        logo_image = logo_image.resize((120, 120), Image.LANCZOS)
        self.logo = ImageTk.PhotoImage(logo_image)

    # _build_gui
    def _build_ui(self):
        # Logo
        logo_label = ttk.Label(self, image=self.logo, background="#f4f6f8")
        logo_label.pack(pady=(15, 8))

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
        desc.pack(pady=(0, 15))

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
        process_btn.pack(pady=15)

        # Status
        self.status_label = ttk.Label(
            self,
            text="No file selected",
            foreground="#555555"
        )
        self.status_label.pack()

        # Separator
        separator = ttk.Separator(self, orient='horizontal')
        separator.pack(fill='x', padx=20, pady=15)

        # Search Section
        self._build_search_section()

        # Footer
        footer = ttk.Label(
            self,
            text="Output file: boltworld.xlsx",
            font=("Segoe UI", 9),
            foreground="#777777"
        )
        footer.pack(side="bottom", pady=10)

    def _build_search_section(self):
        """Build the search functionality section"""
        # Search Header
        search_header = ttk.Label(
            self,
            text="Search Orders",
            style="SubHeader.TLabel"
        )
        search_header.pack(pady=(5, 10))

        # Search Frame
        search_frame = ttk.Frame(self)
        search_frame.pack(padx=20, fill="x")

        # Search Type Selection
        type_frame = ttk.Frame(search_frame)
        type_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(type_frame, text="Search by:").pack(side="left", padx=(0, 10))

        self.search_type = tk.StringVar(value="order")
        
        ttk.Radiobutton(
            type_frame, 
            text="Order Number", 
            variable=self.search_type, 
            value="order"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            type_frame, 
            text="Date", 
            variable=self.search_type, 
            value="date"
        ).pack(side="left", padx=5)
        
        ttk.Radiobutton(
            type_frame, 
            text="User", 
            variable=self.search_type, 
            value="user"
        ).pack(side="left", padx=5)

        # Search Input Frame
        input_frame = ttk.Frame(search_frame)
        input_frame.pack(fill="x")

        self.search_entry = ttk.Entry(input_frame)
        self.search_entry.pack(side="left", expand=True, fill="x", padx=(0, 10))
        self.search_entry.bind('<Return>', lambda e: self.search_orders())

        search_btn = ttk.Button(
            input_frame,
            text="Search",
            command=self.search_orders
        )
        search_btn.pack(side="right")

        # Search hint
        self.search_hint = ttk.Label(
            self,
            text="Enter order number (e.g., 12345)",
            font=("Segoe UI", 8),
            foreground="#888888"
        )
        self.search_hint.pack(pady=(3, 0))

        # Update hint based on search type
        self.search_type.trace('w', self._update_search_hint)

    def _update_search_hint(self, *args):
        """Update search hint based on selected search type"""
        hints = {
            "order": "Enter order number (e.g., 12345)",
            "date": "Enter date in DD-MM-YYYY format (e.g., 19-01-2026)",
            "user": "Enter username"
        }
        self.search_hint.config(text=hints.get(self.search_type.get(), ""))

    def search_orders(self):
        """Search for orders in the Excel file"""
        search_value = self.search_entry.get().strip()
        
        if not search_value:
            messagebox.showwarning(
                "Empty Search",
                "Please enter a search term."
            )
            return

        try:
            from .excel_handler import ExcelHandler
            search_for_order = ExcelHandler(filename="boltworld.xlsx")
            # getting selected search type
            search_type = self.search_type.get()
            # calling the excel_handler search
            results = search_for_order.search(_type=search_type, search_value=search_value, excel_filename="boltworld.xlsx")
            # Display results
            if results[0] == 100:
                self._display_search_results(results, search_value, search_type)
            elif results[0] == 103:     # means that no orders are found
                messagebox.showinfo(
                    "No Results",
                    f"No orders found matching '{search_value}'"
                )

        except Exception as e:
            messagebox.showerror(
                "Search Error",
                f"An error occurred while searching:\n\n{str(e)}"
            )

    def _display_search_results(self, results, search_term, search_type):
        """Display search results in a new window"""
        result_window = tk.Toplevel(self)
        result_window.title("Search Results")
        result_window.geometry("650x450")
        result_window.resizable(True, True)
        result_window.grab_set()

        # Header
        header_text = f"Found {len(results[1])} result(s) for '{search_term}'"
        header = ttk.Label(
            result_window,
            text=header_text,
            font=("Segoe UI", 12, "bold")
        )
        header.pack(pady=(15, 10))

        # Table Frame
        table_frame = ttk.Frame(result_window)
        table_frame.pack(fill="both", expand=True, padx=15, pady=5)

        # Scrollbars
        y_scroll = ttk.Scrollbar(table_frame)
        y_scroll.pack(side="right", fill="y")

        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal")
        x_scroll.pack(side="bottom", fill="x")

        # Treeview for results
        columns = ("Order Number", "Date", "Time", "User")
        tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set,
            height=12
        )

        y_scroll.config(command=tree.yview)
        x_scroll.config(command=tree.xview)

        # Define columns
        tree.heading("Order Number", text="Order Number")
        tree.heading("Date", text="Date")
        tree.heading("Time", text="Time")
        tree.heading("User", text="User")

        tree.column("Order Number", width=150, anchor="center")
        tree.column("Date", width=120, anchor="center")
        tree.column("Time", width=120, anchor="center")
        tree.column("User", width=150, anchor="center")

        # Insert data
        for row in results[1]:
            tree.insert("", "end", values=row)

        tree.pack(fill="both", expand=True)

        # Close button
        action_frame = tk.Frame(result_window)
        action_frame.pack(fill="x", pady=(8, 14))

        close_btn = tk.Button(
            action_frame,
            text="Close",
            command=result_window.destroy,
            font=("Segoe UI", 11, "bold"),
            bg="#2563eb",
            fg="white",
            activebackground="#1e40af",
            activeforeground="white",
            relief="flat",
            padx=50,
            pady=12,
            cursor="hand2"
        )
        close_btn.pack()

        result_window.bind("<Escape>", lambda e: result_window.destroy())

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

            from ..pdfa import PDFAutomation, PDFHandler
            # BAKCEND LINKAGE POINT
            pdf_automation = PDFAutomation()
            # PDF Handler
            pdf_handler = PDFHandler(filename=self.pdf_path.get())
            status_code = pdf_automation.initialize(pdf_handler=pdf_handler, excel_handler=self.excel_handler)

            if status_code == 101:
                # it means that the user didn't closed the file, and still clicked OK
                # promts the user that changes havn't been saved in this case,
                self.pdf_path.set("")
                self.status_label.config(
                    text=f"Changes Not Saved! You need to process pdf again after closing Excel file.",
                    foreground='#821f04'
                )
            else:
                # Reset UI after success
                self.pdf_path.set("")
                self.status_label.config(
                    text="PDF processed successfully ✓",
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
    def prompt_error(code: int , message: dict):
        """Promts an error window to the user with the containing message.

        Args:
            message (str): A dict containting information regarding message.
        """
        if code == 102: # the file doesn't exists
            root = tk.Tk()
            root.withdraw() # hide main window
            response = messagebox.showinfo(
                "No Data",
                "The Excel file doesn't exist yet.\nProcess a PDF first to create the database."
            )
            root.destroy()
        elif code == 101: # the file is already opened by some other process/program
            root = tk.Tk()
            root.withdraw()  # hide main window
            response = messagebox.showwarning(title=message["title"], message=message["message"], icon=message["icon"])
            root.destroy()

    # show_duplicate_orders
    @staticmethod
    def show_duplicate_orders(duplicate_orders: list):
        """Displays duplicate order numbers in a professional UI window"""

        if not duplicate_orders:
            return

        window = tk.Toplevel()
        window.title("Duplicate Orders Detected")
        window.geometry("420x500")
        window.resizable(False, False)
        window.grab_set()  # modal

        # Header
        ttk.Label(
            window,
            text="Duplicate Orders Skipped",
            font=("Segoe UI", 12, "bold")
        ).pack(pady=(15, 5))

        # Description
        ttk.Label(
            window,
            text="The following order numbers already exist\nand were not added again:",
            font=("Segoe UI", 10)
        ).pack(pady=(0, 10))

        # Table frame
        frame = ttk.Frame(window)
        frame.pack(fill="both", expand=True, padx=15, pady=5)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")

        # Listbox (clean & readable)
        listbox = tk.Listbox(
            frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 10),
            height=10
        )
        listbox.pack(side="left", fill="both", expand=True)

        scrollbar.config(command=listbox.yview)

        # Insert duplicates
        for order in duplicate_orders:
            listbox.insert(tk.END, str(order))

        # Close button
        action_frame = tk.Frame(window)
        action_frame.pack(fill="x", pady=(8, 14))

        close_btn = tk.Button(
            action_frame,
            text="Close",
            command=window.destroy,
            font=("Segoe UI", 11, "bold"),
            bg="#2563eb",
            fg="white",
            activebackground="#1e40af",
            activeforeground="white",
            relief="flat",
            padx=50,
            pady=14,
            cursor="hand2"
        )
        close_btn.pack()

        window.after(100, lambda: window.focus_force())
        window.bind("<Escape>", lambda e: window.destroy())

        window.wait_window()  # block until closed