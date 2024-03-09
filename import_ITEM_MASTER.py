import pandas as pd
import pyodbc
from tkinter import Tk, filedialog

def choose_excel_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )

    return file_path

def import_excel_to_sql_server(server, database):
    # Choose Excel file interactively
    excel_file_path = choose_excel_file()

    if not excel_file_path:
        print("No file selected. Exiting.")
        return

    # Read Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file_path)

    # Database connection string
    connection_string = f'Driver={{SQL Server}};Server={server};Database={database};Trusted_Connection=yes;'

    # Establish database connection
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    try:
        # Iterate through each row in the DataFrame and insert into the ITEM_MASTER table
        for index, row in df.iterrows():
            cursor.execute("""
                INSERT INTO ITEM_MASTER (
                    ITEM_NUMBER, ITEM_DESCRIPTION, ITEM_TYPE, MANUFACTURER_CODE,
                    ITEM_CATEGORY, CPU, MEMORY, DISKS, UOM, ENABLED_FLAG
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            str(row['ITEM_NUMBER']), str(row['ITEM_DESCRIPTION']), str(row['ITEM_TYPE']),
            str(row['MANUFACTURER_CODE']), str(row['ITEM_CATEGORY']), str(row['CPU']),
            str(row['MEMORY']), str(row['DISKS']), str(row['UOM']), str(row['ENABLED_FLAG']))

        # Commit the transaction
        connection.commit()
        print("Data imported successfully.")

    except pyodbc.Error as ex:
        print("Error during data import:", ex)
        connection.rollback()

    finally:
        # Close the database connection
        connection.close()

# Specify the database information
server_name = 'LAPTOP-687KHBP5\SQLEXPRESS'
database_name = 'InfraDB'

# Call the function to import data from Excel to SQL Server
import_excel_to_sql_server(server_name, database_name)
