import tkinter as tk
from tkinter import ttk
import pyodbc

def fetch_data(item_id):
    connection_string = 'Driver={SQL Server};' \
                        'Server=LAPTOP-687KHBP5\SQLEXPRESS;'\
                        'Database=InfraDB;' \
                        'Trusted_Connection=yes;'
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    query = f"SELECT * FROM Inventory_Onhand WHERE ITEM_ID = '{item_id}'"
    cursor.execute(query)
    data = cursor.fetchone()
    cursor.close()
    connection.close()
    return data

def update_data(item_id, install_location, project_code, quantity, ip_address, subnet_mask, gateway, comments, last_po_num, last_po_price, renewal_date, notes):
    connection_string = 'Driver={SQL Server};' \
                        'Server=LAPTOP-687KHBP5\SQLEXPRESS;'\
                        'Database=InfraDB;'\
                        'Trusted_Connection=yes;'
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    query = f"UPDATE Inventory_Onhand SET INSTALL_LOCATION='{install_location}', PROJECT_CODE='{project_code}', " \
            f"QUANTITY='{quantity}', IP_ADDRESS='{ip_address}', SUBNET_MASK='{subnet_mask}', " \
            f"GATEWAY='{gateway}', COMMENTS='{comments}', LAST_PO_NUM='{last_po_num}', " \
            f"LAST_PO_PRICE='{last_po_price}', RENEWAL_DATE='{renewal_date}', NOTES='{notes}' " \
            f"WHERE ITEM_ID='{item_id}'"
    cursor.execute(query)
    connection.commit()
    cursor.close()
    connection.close()

def on_button_click():
    item_id = entry_item_id.get()
    data = fetch_data(item_id)
    if data:
        entry_install_location.delete(0, tk.END)
        entry_install_location.insert(0, data.INSTALL_LOCATION)
        entry_project_code.delete(0, tk.END)
        entry_project_code.insert(0, data.PROJECT_CODE)
        entry_quantity.delete(0, tk.END)
        entry_quantity.insert(0, data.QUANTITY)
        entry_ip_address.delete(0, tk.END)
        entry_ip_address.insert(0, data.IP_ADDRESS)
        entry_subnet_mask.delete(0, tk.END)
        entry_subnet_mask.insert(0, data.SUBNET_MASK)
        entry_gateway.delete(0, tk.END)
        entry_gateway.insert(0, data.GATEWAY)
        entry_comments.delete(0, tk.END)
        entry_comments.insert(0, data.COMMENTS)
        entry_last_po_num.delete(0, tk.END)
        entry_last_po_num.insert(0, data.LAST_PO_NUM)
        entry_last_po_price.delete(0, tk.END)
        entry_last_po_price.insert(0, data.LAST_PO_PRICE)
        entry_renewal_date.delete(0, tk.END)
        entry_renewal_date.insert(0, data.RENEWAL_DATE)
        entry_notes.delete(0, tk.END)
        entry_notes.insert(0, data.NOTES)

def on_update_button_click():
    item_id = entry_item_id.get()
    install_location = entry_install_location.get()
    project_code = entry_project_code.get()
    quantity = entry_quantity.get()
    ip_address = entry_ip_address.get()
    subnet_mask = entry_subnet_mask.get()
    gateway = entry_gateway.get()
    comments = entry_comments.get()
    last_po_num = entry_last_po_num.get()
    last_po_price = entry_last_po_price.get()
    renewal_date = entry_renewal_date.get()
    notes = entry_notes.get()
    update_data(item_id, install_location, project_code, quantity, ip_address, subnet_mask, gateway, comments, last_po_num, last_po_price, renewal_date, notes)
    label_update_info = tk.Label(root, text="UPDATED SUCCESSFULLY!!!", bg="lightgrey")
    label_update_info.grid(row=4, column=0, pady=10)

def on_reset_button_click():
    entry_item_id.delete(0, tk.END)
    entry_install_location.delete(0, tk.END)
    entry_project_code.delete(0, tk.END)
    entry_quantity.delete(0, tk.END)
    entry_ip_address.delete(0, tk.END)
    entry_subnet_mask.delete(0, tk.END)
    entry_gateway.delete(0, tk.END)
    entry_comments.delete(0, tk.END)
    entry_last_po_num.delete(0, tk.END)
    entry_last_po_price.delete(0, tk.END)
    entry_renewal_date.delete(0, tk.END)
    entry_notes.delete(0, tk.END)

root = tk.Tk()
root.title("Fetch and Update Data in Database")

fields_frame = ttk.Frame(root)
fields_frame.grid(row=0, column=0, padx=10, pady=10)

label_enter_item_number = tk.Label(fields_frame, text="ENTER   ITEM ID   TO FETCH DATA FROM INVENTORY ON HAND", bg='lightgrey')
label_enter_item_number.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

label_item_id = tk.Label(fields_frame, text="ITEM ID:".upper())
label_item_id.grid(row=1, column=0, padx=10, pady=10)
entry_item_id = tk.Entry(fields_frame)
entry_item_id.grid(row=1, column=1, padx=10, pady=10)

button_fetch_data = tk.Button(fields_frame, text="Fetch Data", command=on_button_click, bg="lightgreen")
button_fetch_data.grid(row=1, column=2, padx=10, pady=10)

label_install_location = tk.Label(fields_frame, text="INSTALL LOCATION:".upper())
label_install_location.grid(row=2, column=0, padx=10, pady=10)
entry_install_location = tk.Entry(fields_frame)
entry_install_location.grid(row=2, column=1, padx=10, pady=10)

label_project_code = tk.Label(fields_frame, text="PROJECT CODE:".upper())
label_project_code.grid(row=3, column=0, padx=10, pady=10)
entry_project_code = tk.Entry(fields_frame)
entry_project_code.grid(row=3, column=1, padx=10, pady=10)

label_quantity = tk.Label(fields_frame, text="QUANTITY:".upper())
label_quantity.grid(row=4, column=0, padx=10, pady=10)
entry_quantity = tk.Entry(fields_frame)
entry_quantity.grid(row=4, column=1, padx=10, pady=10)

label_ip_address = tk.Label(fields_frame, text="IP ADDRESS:".upper())
label_ip_address.grid(row=5, column=0, padx=10, pady=10)
entry_ip_address = tk.Entry(fields_frame)
entry_ip_address.grid(row=5, column=1, padx=10, pady=10)

label_subnet_mask = tk.Label(fields_frame, text="SUBNET MASK:".upper())
label_subnet_mask.grid(row=6, column=0, padx=10, pady=10)
entry_subnet_mask = tk.Entry(fields_frame)
entry_subnet_mask.grid(row=6, column=1, padx=10, pady=10)

label_gateway = tk.Label(fields_frame, text="GATEWAY:".upper())
label_gateway.grid(row=7, column=0, padx=10, pady=10)
entry_gateway = tk.Entry(fields_frame)
entry_gateway.grid(row=7, column=1, padx=10, pady=10)

label_comments = tk.Label(fields_frame, text="COMMENTS:".upper())
label_comments.grid(row=8, column=0, padx=10, pady=10)
entry_comments = tk.Entry(fields_frame)
entry_comments.grid(row=8, column=1, padx=10, pady=10)

label_last_po_num = tk.Label(fields_frame, text="LAST PO NUMBER:".upper())
label_last_po_num.grid(row=9, column=0, padx=10, pady=10)
entry_last_po_num = tk.Entry(fields_frame)
entry_last_po_num.grid(row=9, column=1, padx=10, pady=10)

label_last_po_price = tk.Label(fields_frame, text="LAST PO PRICE:".upper())
label_last_po_price.grid(row=10, column=0, padx=10, pady=10)
entry_last_po_price = tk.Entry(fields_frame)
entry_last_po_price.grid(row=10, column=1, padx=10, pady=10)

label_renewal_date = tk.Label(fields_frame, text="RENEWAL DATE:".upper())
label_renewal_date.grid(row=11, column=0, padx=10, pady=10)
entry_renewal_date = tk.Entry(fields_frame)
entry_renewal_date.grid(row=11, column=1, padx=10, pady=10)

label_notes = tk.Label(fields_frame, text="NOTES:".upper())
label_notes.grid(row=12, column=0, padx=10, pady=10)
entry_notes = tk.Entry(fields_frame)
entry_notes.grid(row=12, column=1, padx=10, pady=10)

# Buttons for update, reset, and exit
button_update_data = tk.Button(root, text="Update", command=on_update_button_click, bg="lightgreen")
button_update_data.grid(row=3, column=0, padx=5, pady=5)

button_reset = tk.Button(root, text="Reset", command=on_reset_button_click, bg="lightblue")
button_reset.grid(row=3, column=1, padx=5, pady=5)

button_exit = tk.Button(root, text="Exit", command=exit, bg="lightcoral")
button_exit.grid(row=3, column=2, padx=5, pady=5)

label_update_info = tk.Label(root, text="3S-Technologies Update Item Master")
label_update_info.grid(row=5, column=0, pady=10)

root.mainloop()
