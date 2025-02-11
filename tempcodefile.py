import customtkinter as ctk
from tkinter import messagebox, simpledialog
from tkinter import ttk
import tkinter as tk
import webbrowser
import sys, os
from datetime import datetime


# Import your business logic functions and constants.
from Transport_Challan import generate_challan, load_data_file, save_data_file, COMPANY_DATA_FILE, TRANSPORT_DATA_FILE

# Import Matplotlib modules for graphing
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -------------------- colors ------------------------------
white = '#FFFFFF' # for values of entry eg supplier name, customer name contact no and so on
grey = '#87888C'
cyan = '#A9DFD8'
deep_slate = '#2B2B36'
bluish_gray = '#171821' 


# -------------------- Helper Functions --------------------

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def create_dropdown(parent, row, column, text, names):
    """Creates a label and dropdown (OptionMenu) in the given parent using grid."""
    sorted_names = sorted(names)
    label = ctk.CTkLabel(parent, text=text, font=("Arial", 16))
    label.grid(row=row, column=column, padx=10, pady=5, sticky="w")
    dropdown = ctk.CTkOptionMenu(parent, values=sorted_names, width=400, font=("Arial", 16))
    dropdown.grid(row=row, column=column + 1, padx=10, pady=5)
    return dropdown

def create_label_and_entry(parent, text, row, column, width=600):
    """Creates a label and an entry (Textbox) in the given parent using grid."""
    label = ctk.CTkLabel(parent, text=text, font=("Arial", 16))
    label.grid(row=row, column=column, padx=5, pady=5, sticky="nsew")
    entry = ctk.CTkTextbox(parent, width=width, height=30)
    entry.grid(row=row, column=column + 1, padx=5, pady=5, sticky="nsew")
    entry.bind("<Tab>", focus_next_widget)
    return entry

def focus_next_widget(event):
    """Allows Tab to move focus to the next widget."""
    event.widget.tk_focusNext().focus_set()
    return "break"

def get_monthly_challan_counts(directory="generated_challans"):
    """
    Scans the specified directory for challan files and counts how many were
    generated in each month. Assumes file names are in the format:
    
    transport_challan_{company}_{transport}_{d}_{m}_{y}_{counter}.xlsx
    
    where the date is in the form "%d_%m_%y" (e.g., "12_12_25").
    
    Returns a dictionary with month abbreviations (e.g., "Dec") as keys and counts as values.
    """
    monthly_counts = {}
    if not os.path.isdir(directory):
        return monthly_counts  # Return an empty dict if the directory doesn't exist
    
    for file in os.listdir(directory):
        if file.startswith("transport_challan") and file.endswith(".xlsx"):
            # Split the filename by underscores
            parts = file.split("_")
            # For a file name like:
            # "transport_challan_ABC_Corp_XYZ_Logistics_12_12_25_1.xlsx"
            # the parts will be:
            # parts[0] = "transport"
            # parts[1] = "challan"
            # parts[2...n-5] = company and transport name (could be multiple parts)
            # parts[-4] = day, parts[-3] = month, parts[-2] = year, parts[-1] = counter + ".xlsx"
            if len(parts) >= 7:
                # Combine the three parts for the date
                date_str = f"{parts[-4]}_{parts[-3]}_{parts[-2]}"
                try:
                    # Parse the date using the format "%d_%m_%y"
                    dt = datetime.strptime(date_str, "%d_%m_%y")
                    # Use abbreviated month name (e.g., "Dec")
                    month = dt.strftime("%b")
                    monthly_counts[month] = monthly_counts.get(month, 0) + 1
                except ValueError:
                    # Skip this file if the date format is not as expected
                    continue
    return monthly_counts

def load_names():
    """Load supplier and customer names from data files."""
    company_data = load_data_file(COMPANY_DATA_FILE)
    transport_data = load_data_file(TRANSPORT_DATA_FILE)
    company_names = list(company_data.keys())
    transport_names = list(transport_data.keys())
    return company_names, transport_names

# Load names globally
supplier_names, customer_names = load_names()

# Set appearance and color theme
ctk.set_appearance_mode("system")
# ctk.set_default_color_theme("green")

# -------------------- Main Application with Tab Navigation --------------------

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=deep_slate)  
        self.title("Transport Challan")
        self.geometry("2000x840")
        
        # Configure grid for sidebar and main_frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # Sidebar frame with dark color
        self.sidebar = ctk.CTkFrame(self, width=150, fg_color=deep_slate, bg_color = bluish_gray, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="ns")
        
        # Main content area - set fg_color to match main app background
        self.main_frame = ctk.CTkFrame(self, fg_color=deep_slate, corner_radius=0)  # Match fg_color here
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=2)
        
        # Create pages (each is a CTkFrame)
        self.pages = {
            "home": HomePage(self.main_frame),
            "add_supplier": AddSupplierPage(self.main_frame),
            "add_customer": AddCustomerPage(self.main_frame),
            "statistics": StatisticsPage(self.main_frame),
            "about": AboutPage(self.main_frame)
        }
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
        
        # Create sidebar navigation buttons
        self.create_sidebar_button("🏠 Home", "home", white)
        self.create_sidebar_button("📦 Add Supplier", "add_supplier", '#FCB859')
        self.create_sidebar_button("👤 Add Customer", "add_customer", '#F2C8ED')
        self.create_sidebar_button("📊 Statistics", "statistics", '#A9DFD8')
        self.create_sidebar_button("About", "about", white)
        
        # Show Home page by default
        self.pages["home"].tkraise()
    
    def create_sidebar_button(self, text, page_name, text_color):
        btn = ctk.CTkButton(self.sidebar, text=text,fg_color=deep_slate, hover_color=bluish_gray,text_color=text_color, command=lambda: self.pages[page_name].tkraise())
        btn.pack(pady=10, padx=10, fill="x")

# -------------------- Page Classes --------------------

class HomePage(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=deep_slate)
        # Title
        title = ctk.CTkLabel(self, text="💳 Generate Challan", font=("Arial", 24))
        title.grid(row=0, column=0, columnspan=2, pady=20)
        
        # Dropdowns for supplier and customer names
        self.company_text = create_dropdown(self, 1, 0, "Select Supplier Name:", supplier_names)
        self.transport_text = create_dropdown(self, 2, 0, "Select Customer Name:", customer_names)
        
        # Entry fields for contact, challan number, and date
        self.contact_no_text = create_label_and_entry(self, "Contact No:", 3, 0)
        self.challan_number_text = create_label_and_entry(self, "Challan Number:", 4, 0)
        self.date_text = create_label_and_entry(self, "Date (DD.MM.YY):", 5, 0)
        self.company_text.set("Select Company")
        self.transport_text.set("Select transport")
        
        # Item entry section in its own frame
        self.item_frame = ctk.CTkFrame(self, fg_color=deep_slate)
        self.item_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=10, sticky="nsew")
        
        # Item detail fields
        self.item_name_text = create_label_and_entry(self.item_frame, "Item Name:", 0, 0, width=400)
        self.item_name_text.insert('1.0', 'saree')
        self.hsn_text = create_label_and_entry(self.item_frame, "HSN Code:", 0, 2, width=200)
        self.hsn_text.insert('1.0', '5407')
        self.pieces_text = create_label_and_entry(self.item_frame, "Pieces:", 0, 4, width=100)
        self.amount_text = create_label_and_entry(self.item_frame, "Amount:", 0, 6, width=100)
        self.discount_text = create_label_and_entry(self.item_frame, "Discount:", 5, 0, width=200)
        self.gst_text = create_label_and_entry(self.item_frame, "GST:", 5, 2, width=200)
        
        # Treeview for displaying added items
        self.items = []
        style = ttk.Style()
        style.configure("Treeview", rowheight=20, font=('Arial', 20))
        style.configure("Treeview.Heading", font=('Arial', 20))
        self.tree = ttk.Treeview(self.item_frame, columns=("Name", "HSN", "Pieces", "Amount"), show="headings")
        self.tree.heading("Name", text="Item Name")
        self.tree.heading("HSN", text="HSN Code")
        self.tree.heading("Pieces", text="Pieces")
        self.tree.heading("Amount", text="Amount")
        self.tree.grid(row=3, column=0, columnspan=4, pady=7, sticky="nsew")
        
        # Buttons for item management
        self.add_item_button = ctk.CTkButton(self.item_frame, text="Add Item", command=self.add_item)
        self.add_item_button.grid(row=2, column=0, columnspan=7, pady=10)
        self.clear_button = ctk.CTkButton(self.item_frame, text="Clear Items", command=self.clear_items)
        self.clear_button.grid(row=4, column=0, columnspan=7, pady=10)
        
        # Other party fields
        self.no_of_other_goods = create_label_and_entry(self, "No. of other party goods:", 7, 0)
        self.amount_of_other_goods = create_label_and_entry(self, "Amount of other party goods:", 8, 0)
        
        # Generate Challan Button
        self.generate_button = ctk.CTkButton(self, text="Generate Challan", command=self.submit_data)
        self.generate_button.place(relx=0.5, rely=0.95, anchor='center')
    
    def clear_items(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.items.clear()

    def add_item(self):
        item_name = self.item_name_text.get("1.0", "end-1c").strip()
        hsn = self.hsn_text.get("1.0", "end-1c").strip()
        pieces = self.pieces_text.get("1.0", "end-1c").strip()
        amount = self.amount_text.get("1.0", "end-1c").strip()
        if not item_name or not hsn or not pieces.isdigit() or not amount.isdigit():
            print("Please enter valid item details.")
            return
        self.tree.insert("", "end", values=(item_name, hsn, pieces, amount))
        self.items.append((item_name, hsn, int(pieces), int(amount)))
        # self.item_name_text.delete("1.0", "end")
        # self.hsn_text.delete("1.0", "end")
        self.pieces_text.delete("1.0", "end")
        self.amount_text.delete("1.0", "end")
    
    def submit_data(self):
        company_data = load_data_file(COMPANY_DATA_FILE)
        transport_data = load_data_file(TRANSPORT_DATA_FILE)
        company_name = self.company_text.get().strip().upper()
        transport_name = self.transport_text.get().strip().upper()
        date = self.date_text.get("1.0", "end-1c").strip()
        discount = self.discount_text.get("1.0", "end-1c").strip()
        gst = self.gst_text.get("1.0", "end-1c").strip().upper()
        contact_no = self.contact_no_text.get("1.0", "end-1c").strip()
        
        # If supplier or customer not found, prompt to add (logic can be extended)
        if company_name not in company_data:
            response = messagebox.askyesno("Add Supplier", f"Supplier '{company_name}' not found. Add it?")
            if not response:
                return
        if transport_name not in transport_data:
            response = messagebox.askyesno("Add Customer", f"Customer '{transport_name}' not found. Add it?")
            if not response:
                return
        
        if not company_name or not transport_name or not date or not discount.isdigit() or not gst.isdigit():
            messagebox.showerror("Error", "Please enter valid data in all fields.")
            return
        
        challan_number = self.challan_number_text.get("1.0", "end-1c").strip()
        if not challan_number:
            challan_number = simpledialog.askstring("Challan Number", "Please enter the Challan Number:")
        if not challan_number:
            messagebox.showerror("Error", "Challan Number is required.")
            return
        
        # Call the generate_challan function
        generate_challan(
            company_data=company_data,
            transport_data=transport_data,
            contact_no=contact_no,
            company_name=company_name,
            transport_name=transport_name,
            items_data=self.items,
            discount=int(discount),
            gst=int(gst),
            date=date,
            challan_number=challan_number,
            No_of_Other_Party_Goods=self.no_of_other_goods.get("1.0", "end-1c").strip(),
            Amount_of_Other_Party_Goods=self.amount_of_other_goods.get("1.0", "end-1c").strip()
        )
        messagebox.showinfo("Success", "Challan generated successfully!")
        
        # Clear fields after submission
        self.company_text.set("Select Company")
        self.transport_text.set("Select transport")
        self.challan_number_text.delete("1.0", "end")
        self.date_text.delete("1.0", "end")
        self.discount_text.delete("1.0", "end")
        self.gst_text.delete("1.0", "end")
        self.clear_items()
        self.no_of_other_goods.delete("1.0", "end")
        self.amount_of_other_goods.delete("1.0", "end")

class AddSupplierPage(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=deep_slate)
        title = ctk.CTkLabel(self, text="📦 Add/Modify Supplier", text_color='#FCB859', font=("Arial", 24))
        title.grid(row=0, column=0, columnspan=2, pady=20)
        # Supplier details fields
        self.address1_text = create_label_and_entry(self, "Address(1) of supplier:", 1, 0)
        self.address2_text = create_label_and_entry(self, "Address(2) of supplier:", 2, 0)
        self.gst_text = create_label_and_entry(self, "GST no. of supplier:", 3, 0)
        self.submit_button = ctk.CTkButton(self, text="Submit Supplier", command=self.submit_supplier_data)
        self.submit_button.grid(row=4, column=0, columnspan=2, pady=20)
    
    def submit_supplier_data(self):
        company_name = simpledialog.askstring("Supplier Name", "Enter Supplier Name:")
        if not company_name or not company_name.strip():
            messagebox.showerror("Error", "Supplier Name is required.")
            return
        company_name = company_name.strip().upper()
        company_data = load_data_file(COMPANY_DATA_FILE)
        company_data[company_name] = {
            'address1': self.address1_text.get("1.0", "end-1c").strip().upper(),
            'address2': self.address2_text.get("1.0", "end-1c").strip(),
            'gst': self.gst_text.get("1.0", "end-1c").strip(),
        }
        save_data_file(COMPANY_DATA_FILE, company_data)
        messagebox.showinfo("Success", f"Supplier '{company_name}' added/modified successfully!")
        self.address1_text.delete("1.0", "end")
        self.address2_text.delete("1.0", "end")
        self.gst_text.delete("1.0", "end")
        
        # Update the supplier dropdown in HomePage
        # Traverse the widget hierarchy: self -> main_frame -> MainApp
        home_page = self.master.master.pages.get("home")
        if home_page:
            new_supplier_names, _ = load_names()
            # Update the dropdown values and optionally set a default value
            home_page.company_text.configure(values=sorted(new_supplier_names))

class AddCustomerPage(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent , fg_color=deep_slate)
        title = ctk.CTkLabel(self, text="👤 Add/Modify Customer",text_color='#F2C8ED', font=("Arial", 24))
        title.grid(row=0, column=0, columnspan=2, pady=20)
        # Customer details fields
        self.station_text = create_label_and_entry(self, "Address of customer:", 1, 0)
        self.gst_text = create_label_and_entry(self, "GST no. of customer:", 2, 0)
        self.way_text = create_label_and_entry(self, "Transport Way:", 3, 0)
        self.submit_button = ctk.CTkButton(self, text="Submit Customer", command=self.submit_customer_data)
        self.submit_button.grid(row=4, column=0, columnspan=2, pady=20)
    
    def submit_customer_data(self):
        customer_name = simpledialog.askstring("Customer Name", "Enter Customer Name:")
        if not customer_name or not customer_name.strip():
            messagebox.showerror("Error", "Customer Name is required.")
            return
        customer_name = customer_name.strip().upper()
        transport_data = load_data_file(TRANSPORT_DATA_FILE)
        transport_data[customer_name] = {
            'station': self.station_text.get("1.0", "end-1c").strip().upper(),
            'gst': self.gst_text.get("1.0", "end-1c").strip(),
            'Way': self.way_text.get("1.0", "end-1c").strip().upper(),
        }
        save_data_file(TRANSPORT_DATA_FILE, transport_data)
        messagebox.showinfo("Success", f"Customer '{customer_name}' added/modified successfully!")
        self.gst_text.delete("1.0", "end")
        self.way_text.delete("1.0", "end")
        self.station_text.delete("1.0", "end")
        
        # Update the customer dropdown in HomePage
        home_page = self.master.master.pages.get("home")
        if home_page:
            _, new_customer_names = load_names()
            home_page.transport_text.configure(values=sorted(new_customer_names))

class StatisticsPage(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=deep_slate)
        # Title label
        title = ctk.CTkLabel(self, text="📊 Statistics", font=("Arial", 24))
        title.pack(pady=20)
        
        # Add a refresh button so the user can update the graph manually
        refresh_button = ctk.CTkButton(self, text="Refresh", command=self.update_graph)
        refresh_button.pack(pady=10)
        
        # Placeholder for the canvas so we can update it later.
        self.canvas = None
        
        # Build the initial graph.
        self.update_graph()

    def update_graph(self):
        # If a canvas already exists, destroy it to clear the old graph.
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
        
        # Get the updated monthly counts from the generated files.
        monthly_counts = get_monthly_challan_counts()  # Your helper function that returns a dict like {"Jan": count, ...}
        if not monthly_counts:
            monthly_counts = {}  # Or you could use {"No Data": 0} if you prefer.
        
        # Mapping month abbreviations to their numeric values.
        month_map = {
            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
            "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
            "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
        }
        
        # Get the current month number.
        current_month = datetime.now().month
        
        # Define a ranking function to order months cyclically so that the current month is last.
        def rank_func(m):
            num = month_map[m]
            # If the month number is greater than the current month,
            # subtract 12 to place it before the current month in the cyclic order.
            return num if num <= current_month else num - 12
        
        # Sort the keys (month abbreviations) from monthly_counts using the ranking function.
        ordered_months = sorted(monthly_counts.keys(), key=rank_func)
        counts = [monthly_counts[m] for m in ordered_months]
        
        # Create a new Matplotlib figure.
        fig = Figure(figsize=(6, 4), dpi=100)
        ax = fig.add_subplot(111)
        ax.bar(ordered_months, counts, color='skyblue')
        ax.set_title("Challans Generated Per Month")
        ax.set_xlabel("Month")
        ax.set_ylabel("Number of Challans")
        if counts:
            ax.set_ylim(0, max(counts) + 5)  # Add some headroom
        
        # Embed the figure into the CustomTkinter frame.
        self.canvas = FigureCanvasTkAgg(fig, master=self)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

class AboutPage(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=deep_slate)
        title = ctk.CTkLabel(self, text="About App", font=("Arial", 24))
        title.pack(pady=20)
        about_text = (
            "Developed by: Mayank Dhanuka\n"
            "This application helps generate Transport Challans for suppliers and customers.\n\n"
            "Guidelines:\n"
            " - Please do not leave any field blank.\n"
            " - If a field needs to be blank, type '00'.\n\n"
            "For issues or suggestions, contact us via the email link below."
        )
        text_label = ctk.CTkLabel(self, text=about_text, font=("Arial", 14), justify="left")
        text_label.pack(pady=10)
        gmail_button = ctk.CTkButton(
            self, text="Email: mayankdhanuka899@gmail.com",
            font=("Arial", 14, "underline"),
            fg_color="transparent",
            text_color="blue",
            hover_color="lightblue",
            command=self.open_gmail
        )
        gmail_button.pack(pady=10)
        version_label = ctk.CTkLabel(self, text="Version 1.0 (Beta)", font=("Arial", 14, "italic"))
        version_label.pack(pady=10)
    
    def open_gmail(self):
        gmail_url = "https://mail.google.com/mail/?view=cm&fs=1&to=mayankdhanuka899@gmail.com"
        webbrowser.open(gmail_url)

# -------------------- Main Function --------------------

def main():
    app = MainApp()
    app.mainloop()

if __name__ == '__main__':
    main()
