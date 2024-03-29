import tkinter as tk
from tkinter import ttk, font
import pyodbc
from datetime import datetime
import sys

app = tk.Tk()
app.geometry("700x500")
app.title("INSERT PURCHASE ORDER")

# Label above the frame
instruction_label = ttk.Label(app, text="INSERT THE BELOW FIELDS INTO THE PURCHASE ORDER", foreground="black", font=font.Font(size=11), background="#CCCCCC")
instruction_label.place(relx=0.1, rely=0.1)



# Create a frame to hold the labels and entry fields with a scrollbar
frame = tk.Frame(app, borderwidth=2, relief="groove", border=2, bg="grey")
frame.place(relx=0.1, rely=0.2, relwidth=0.8, relheight=0.6)

canvas = tk.Canvas(frame)
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Function to add labels and entry fields to the scrollable frame
def add_label_and_entry(label_text, row):
    label = ttk.Label(scrollable_frame, text=label_text)
    label.grid(row=row, column=0, sticky='w', padx=10, pady=5)  # Adjusted padx and pady for spacing

    if label_text == "PO STATUS:":
        entry = ttk.Combobox(scrollable_frame)
        entry['values'] = ['Entered', 'Approved', 'Received', 'Invoiced', 'Closed', 'Cancelled']
    else:
        entry = ttk.Entry(scrollable_frame)

        
    entry.grid(row=row, column=1, padx=10, pady=5)  # Adjusted padx and pady for spacing
    return entry

# Labels and entry fields
labels_and_entries = [
    ("PO HEADER ID:", 33),
    ("PO LINE NUMBER", 34),
    ("ITEM ID:", 35),
    ("PO LINE DESCRIPTION:", 36),

    ("QUANTITY:", 37),
    ("UNIT PRICE:", 38),
    ("LINE TAX AMOUNT", 39),
    ("SUPPORT START DATE:", 40),
    ("SUPPORT END DATE:", 41),
    ("NEED BY DATE:", 42),

    ("PO LINE STATUS:", 43),
    ("SHIP LOCATION", 44),
    ("INVOICE NUMBER:", 45),
    ("INVOICE LINE NUMBER:", 46),
    ("INVOICE DATE:", 47),

    ("INVOICE PAID:", 48),
    ("INVOICE AMOUNT:", 49),
    ("PO LINE COMMENTS:", 50),
]


#ENTRY FIELDS
entry_fields = []
for label_text, row in labels_and_entries:
    entry = add_label_and_entry(label_text, row)
    entry_fields.append(entry)

def insert_inventory_onhand():
    try:
        connection = pyodbc.connect('Driver={SQL Server};'
                       'Server=LAPTOP-687KHBP5\SQLEXPRESS;'
                        'Database=InfraDB;'
                        'Trusted_Connection=yes;')
        connection.autocommit = True

        # Get the current date and time as a string
        current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Assuming username is passed as a command-line argument or an empty string if not provided
        username = sys.argv[1] if len(sys.argv) > 1 else ''

        # Use a tuple to unpack the entry fields for the SQL query
        query_params = tuple(entry.get() for entry in entry_fields)

        # Add current_date and username to the tuple for query_params
        query_params += (current_date, current_date)

        # Use parameterized query to avoid SQL injection and handle date conversion
        connection.execute("""
            INSERT INTO PO_LINES 
            (PO_HEADER_ID, PO_LINE_NUMBER, ITEM_ID, PO_LINE_DESCRIPTION, QUANTITY, UNIT_PRICE,
            LINE_TAX_AMOUNT, SUPPORT_START_DATE, SUPPORT_END_DATE, NEED_BY_DATE, PO_LINE_STATUS,
            SHIP_LOCATION, INVOICE_NUMBER, INVOICE_LINE_NUMBER, INVOICE_DATE, INVOICE_PAID, INVOICE_AMOUNT, PO_LINE_COMMENTS,
             CREATION_DATE, LAST_UPDATE_DATE) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)

        """, query_params)

        info_label_inventory.configure(text="INSERTION COMPLETED!")

    except pyodbc.Error as ex:
        print("CONNECTION FAILED", ex)

def reset():
    # Reset all entry fields to empty strings
    for entry in entry_fields:
        entry.delete(0, tk.END)


def cancel():
    app.destroy()

# Create a new frame for the buttons
button_frame = tk.Frame(app)
button_frame.place(relx=0.1, rely=0.8, relwidth=0.8)

# Function to create bold font
def get_bold_font():
    return font.Font(weight="bold")

# Create buttons with bold text
insert_button = tk.Button(button_frame, text="INSERT", command=insert_inventory_onhand,
                          foreground="black", background="#9ccc65", font=font.Font(size=10, weight="bold"), width=7, height=1)
insert_button.grid(row=0, column=0, pady=(10, 5), padx=50)

reset_button = tk.Button(button_frame, text="RESET", command=reset,
                         foreground="black", background="#64b5f6", font=font.Font(size=10, weight="bold"), width=7, height=1)
reset_button.grid(row=0, column=1, pady=(10, 5), padx=50)

cancel_button = tk.Button(button_frame, text="EXIT", command=cancel,
                          foreground="black", background="#ef5350", font=font.Font(size=10, weight="bold"), width=7, height=1)
cancel_button.grid(row=0, column=2, pady=(10, 5), padx=50)

info_label_inventory = ttk.Label(app, text="3S Technologies - ITEM MASTER")
info_label_inventory.place(relx=0.1, rely=0.95)  # Adjusted y-position

app.mainloop()