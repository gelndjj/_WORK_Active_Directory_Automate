from tkinter import *
from tkinter import ttk, messagebox, filedialog, font, simpledialog
import sqlite3, csv, os, subprocess, sys, shutil, time, glob, shlex, datetime, ctypes, random, string, tkinter as tk
from faker import Faker
from tkinter.scrolledtext import ScrolledText

try:
    # Try to set DPI awareness to make text and elements clear
    ctypes.windll.shcore.SetProcessDpiAwareness(1) # 1: System DPI aware, 2: Per monitor DPI aware
except AttributeError:
    # Fallback if SetProcessDpiAwareness does not exist (Windows versions < 8.1)
    ctypes.windll.user32.SetProcessDPIAware()

root = Tk()
root.title('Active Directory Automate - Users')
root.geometry("1230x830")
root.resizable(False, False)
root.iconbitmap("icons.ico")
root.tk_setPalette(background="#ececec")

# Create a Style instance
style = ttk.Style()

# Set the theme to "clam"
style.theme_use("clam")

# Create a Notebook (tab container)
notebook = ttk.Notebook(root)

# Create a Frame for the first tab
info_tab = ttk.Frame(notebook)
tools_tab = ttk.Frame(notebook)

# Add the tab to the notebook with a name
notebook.add(info_tab, text='Users Information')
notebook.add(tools_tab, text='Tools')

# Place the notebook on the root window
notebook.pack(expand=True, fill='both')

# Add Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Initialize fake_record_count as a global variable
fake_record_count = 0

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# Create the 'domain' folder if it doesn't exist
domain_folder = os.path.join(script_dir, "domain")
os.makedirs(domain_folder, exist_ok=True)

# Create the 'addresses' folder if it doesn't exist
addresses_folder = os.path.join(script_dir, "addresses")
os.makedirs(addresses_folder, exist_ok=True)

# Create the 'Organization unit' folder if it doesn't exist
organization_unit_folder = os.path.join(script_dir, "organization_unit")
os.makedirs(organization_unit_folder, exist_ok=True)

def get_application_path():
    if getattr(sys, 'frozen', False):
        # Running as an executable
        return os.path.dirname(sys.executable)
    else:
        # Running as a regular Python script
        return os.path.dirname(os.path.abspath(__file__))

def create_new_database():
    # Prompt the user to enter the name for the new database
    new_db_name = simpledialog.askstring("New Database", "Enter name for the new database:")

    if new_db_name:
        # Ensure the database name ends with '.db'
        if not new_db_name.endswith('.db'):
            new_db_name += '.db'

        # Inside create_new_database function
        script_dir = get_application_path()

        new_db_path = os.path.join(script_dir, new_db_name)
        # Create a database or connect to one that exists
        conn = sqlite3.connect(new_db_path)

        # Create a cursor instance
        c = conn.cursor()

        # Create a customers table (if it doesn't exist)
        c.execute("""CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY,
            last_name TEXT,
            first_name TEXT,
            display_name TEXT,
            description TEXT,
            email TEXT,
            sam TEXT,
            upn TEXT,
            address TEXT,
            city TEXT,
            state TEXT,
            zipcode TEXT,
            company TEXT,
            job_title TEXT,
            department TEXT,
            cid TEXT
        )""")

        # Commit changes and close the connection
        conn.commit()
        conn.close()

        # Extract the filename with extension from the full path
        new_db_filename = os.path.basename(new_db_path)

        # Set the new database filename as the selected database in the ComboBox
        db_combobox.set(new_db_filename)
        populate_db_combobox()

def populate_csv_combobox():
    try:
        #script_dir = os.path.dirname(__file__)
        #addresses_folder = os.path.join(script_dir, "addresses")

        # Get a list of CSV files in the "addresses" folder
        csv_files = [f for f in os.listdir(addresses_folder) if f.endswith('.csv')]

        # Set the values of the ComboBox to the list of CSV files
        csv_combobox['values'] = csv_files
    except FileNotFoundError:
        messagebox.showerror("Error", "The 'addresses' folder was not found.")

def populate_db_combobox():
    global db_combobox  # Declare db_combobox as a global variable

    # Inside populate_db_combobox function
    script_dir = get_application_path()
    db_files = [f for f in os.listdir(script_dir) if f.endswith('.db')]

    # Set the values of the ComboBox to the list of database files
    db_combobox['values'] = db_files

# Function to load the selected database and display its contents in the Treeview
def load_selected_db(event):
    selected_db = db_combobox.get()

    if selected_db:
        try:
            # Connect to the selected database
            conn = sqlite3.connect(selected_db)
            c = conn.cursor()

            # Clear existing Treeview data
            my_tree.delete(*my_tree.get_children())

            # Fetch data from the database and display it in the Treeview
            c.execute("SELECT * FROM customers")
            data = c.fetchall()
            for record in data:
                my_tree.insert("", "end", values=record)

            # Commit changes and close the connection
            conn.commit()
            conn.close()
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"Error loading selected database: {e}")

def query_database():
    # Check if a database is selected in the ComboBox
    selected_db = db_combobox.get()

    if selected_db:
        try:
            # Connect to the selected database
            conn = sqlite3.connect(selected_db)
            c = conn.cursor()

            # Clear existing Treeview data
            my_tree.delete(*my_tree.get_children())

            # Fetch data from the database and display it in the Treeview
            c.execute("SELECT * FROM customers")
            data = c.fetchall()
            for record in data:
                my_tree.insert("", "end", values=record)

            # Commit changes and close the connection
            conn.commit()
            conn.close()
        except sqlite3.Error as e:
            pass  # Ignore errors silently

def search_records():
    lookup_record = search_entry.get()

    # Close the search box
    search.destroy()

    # Clear the Treeview
    for record in my_tree.get_children():
        my_tree.delete(record)

    # Get the selected database filename from the ComboBox
    selected_db = db_combobox.get()

    if not selected_db:
        messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
        return

    # Create a database connection for the selected database
    conn = sqlite3.connect(selected_db)
    c = conn.cursor()

    c.execute("SELECT rowid, * FROM customers WHERE last_name LIKE ?", (f"%{lookup_record}%",))
    records = c.fetchall()

    # Add our data to the screen
    global count
    count = 0

    for record in records:
        if count % 2 == 0:
            my_tree.insert(parent='', index='end', iid=count, text='',
                           values=(
                               record[1], record[2], record[0], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12], record[13], record[14], record[15]),
                           tags=('evenrow',))
        else:
            my_tree.insert(parent='', index='end', iid=count, text='',
                           values=(
                               record[1], record[2], record[3], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12], record[13], record[14], record[15]),
                           tags=('oddrow',))
        # Increment counter
        count += 1

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()

def lookup_records():
    global search_entry, search

    search = Toplevel(root)
    search.title("Search Records")
    search.geometry("400x200")
    search.iconbitmap("icons.ico")

    # Create label frame
    search_frame = LabelFrame(search, text="Last Name")
    search_frame.pack(padx=10, pady=10)

    # Add entry box
    search_entry = ttk.Entry(search_frame, font=("Helvetica", 18))
    search_entry.pack(pady=20, padx=20)

    # Add button
    search_button = ttk.Button(search, text="Search Records", command=search_records)
    search_button.pack(padx=20, pady=20)

# Set the theme to "clam" initially
style.theme_use("clam")

# Add Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Configure our menu
option_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="Options", menu=option_menu)

# Search Menu
search_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="Search", menu=search_menu)

# Drop down menu
option_menu.add_separator()
option_menu.add_command(label="Exit", command=root.quit)

# Drop down menu
search_menu.add_command(label="Search", command=lookup_records)
search_menu.add_separator()
search_menu.add_command(label="Reset", command=query_database)

# Create a Treeview Frame
tree_frame = ttk.Frame(info_tab)
tree_frame.pack(pady=10, fill=Y)

# Create a vertical scrollbar
tree_scroll_y = ttk.Scrollbar(tree_frame, orient='vertical')
tree_scroll_y.pack(side=RIGHT, fill=Y)

# Create a horizontal scrollbar
tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal')
tree_scroll_x.pack(side=BOTTOM, fill=X)

# Create The Treeview inside a separate frame
treeview_frame = ttk.Frame(tree_frame)
treeview_frame.pack(fill=BOTH, expand=True)

my_tree = ttk.Treeview(
    treeview_frame,
    selectmode="extended",
    style="Treeview",
    yscrollcommand=tree_scroll_y.set,
    xscrollcommand=tree_scroll_x.set  # Add horizontal scrollbar command
)
my_tree.pack(fill=BOTH, expand=True)

# Configure the Treeview to use the vertical scrollbar
tree_scroll_y.config(command=my_tree.yview)

# Configure the Treeview to use the horizontal scrollbar
tree_scroll_x.config(command=my_tree.xview)

# Create a StringVar to track the display_name
display_name_var = StringVar()

# Define Our Columns
my_tree['columns'] = (
"ID", "Last Name", "First Name", "Display Name", "Description", "Email", "SAM", "UPN", "Address", "City", "State",
"Zipcode", "Company", "Job Title", "Department", "Custom Identifier")

# Format Our Columns
my_tree.column("#0", width=0, stretch=NO)
my_tree.column("ID", anchor=CENTER, width=40)
my_tree.column("Last Name", anchor=CENTER, width=150)
my_tree.column("First Name", anchor=CENTER, width=150)
my_tree.column("Display Name", anchor=CENTER, width=200)
my_tree.column("Description", anchor=CENTER, width=150)
my_tree.column("Email", anchor=CENTER, width=220)
my_tree.column("SAM", anchor=CENTER, width=80)
my_tree.column("UPN", anchor=CENTER, width=170)
my_tree.column("Address", anchor=CENTER, width=170)
my_tree.column("City", anchor=CENTER, width=130)
my_tree.column("State", anchor=CENTER, width=100)
my_tree.column("Zipcode", anchor=CENTER, width=100)
my_tree.column("Company", anchor=CENTER, width=170)
my_tree.column("Job Title", anchor=CENTER, width=170)
my_tree.column("Department", anchor=CENTER, width=170)
my_tree.column("Custom Identifier", anchor=CENTER, width=170)

# Create Headings
my_tree.heading("#0", text="", anchor=W)
my_tree.heading("ID", text="ID", anchor=CENTER)
my_tree.heading("Last Name", text="Last Name", anchor=CENTER)
my_tree.heading("First Name", text="First Name", anchor=CENTER)
my_tree.heading("Display Name", text="Display Name", anchor=CENTER)
my_tree.heading("Description", text="Description", anchor=CENTER)
my_tree.heading("Email", text="Email", anchor=CENTER)
my_tree.heading("SAM", text="SAM", anchor=CENTER)
my_tree.heading("UPN", text="UPN", anchor=CENTER)
my_tree.heading("Address", text="Address", anchor=CENTER)
my_tree.heading("City", text="City", anchor=CENTER)
my_tree.heading("State", text="State", anchor=CENTER)
my_tree.heading("Zipcode", text="Zipcode", anchor=CENTER)
my_tree.heading("Company", text="Company", anchor=CENTER)
my_tree.heading("Job Title", text="Job Title", anchor=CENTER)
my_tree.heading("Department", text="Department", anchor=CENTER)
my_tree.heading("Custom Identifier", text="Custom Identifier", anchor=CENTER)

# Create Striped Row Tags
my_tree.tag_configure('oddrow')
my_tree.tag_configure('evenrow')

# Add Record Entry Boxes
data_frame = LabelFrame(info_tab, text="Record", border=5, bg="#dcdad5")
data_frame.place(x=15, y=270)

id_var = StringVar()
id_var.set("Auto-Generated")
id_entry = ttk.Entry(data_frame, textvariable=id_var, state='disabled')
id_entry.grid(row=0, column=1, padx=10, pady=5)

id_label = Label(data_frame, text="ID", font=("Academy Engraved", 10), bg="#dcdad5")
id_label.grid(row=0, column=0, padx=10, pady=10)

ln_label = Label(data_frame, text="Last Name", font=("Academy Engraved", 10), bg="#dcdad5")
ln_label.grid(row=0, column=2, padx=10, pady=10)
ln_entry = ttk.Entry(data_frame)
ln_entry.grid(row=0, column=3, padx=10, pady=10)

fn_label = Label(data_frame, text="First Name", font=("Academy Engraved", 10), bg="#dcdad5")
fn_label.grid(row=0, column=4, padx=10, pady=10)
fn_entry = ttk.Entry(data_frame)
fn_entry.grid(row=0, column=5, padx=10, pady=10)

dn_label = Label(data_frame, text="Display Name", font=("Academy Engraved", 10), bg="#dcdad5")
dn_label.grid(row=0, column=6, padx=10, pady=10)
dn_entry = ttk.Entry(data_frame, textvariable=display_name_var)
dn_entry.grid(row=0, column=7, padx=10, pady=10)

dc_label = Label(data_frame, text="Description", font=("Academy Engraved", 10), bg="#dcdad5")
dc_label.grid(row=1, column=0, padx=10, pady=10)
dc_entry = ttk.Entry(data_frame)
dc_entry.grid(row=1, column=1, padx=10, pady=10)

em_label = Label(data_frame, text="Email", font=("Academy Engraved", 10), bg="#dcdad5")
em_label.grid(row=1, column=2, padx=10, pady=10)
em_entry = ttk.Entry(data_frame)
em_entry.grid(row=1, column=3, padx=10, pady=10)

sm_label = Label(data_frame, text="SAM", font=("Academy Engraved", 10), bg="#dcdad5")
sm_label.grid(row=1, column=4, padx=10, pady=10)
sm_entry = ttk.Entry(data_frame)
sm_entry.grid(row=1, column=5, padx=10, pady=10)

up_label = Label(data_frame, text="UPN", font=("Academy Engraved", 10), bg="#dcdad5")
up_label.grid(row=1, column=6, padx=10, pady=10)
up_entry = ttk.Entry(data_frame)
up_entry.grid(row=1, column=7, padx=10, pady=10)

address_label = Label(data_frame, text="Address", font=("Academy Engraved", 10), bg="#dcdad5")
address_label.grid(row=2, column=0, padx=10, pady=10)
address_entry = ttk.Entry(data_frame)
address_entry.grid(row=2, column=1, padx=10, pady=10)

city_label = Label(data_frame, text="City", font=("Academy Engraved", 10), bg="#dcdad5")
city_label.grid(row=2, column=2, padx=10, pady=10)
city_entry = ttk.Entry(data_frame)
city_entry.grid(row=2, column=3, padx=10, pady=10)

state_label = Label(data_frame, text="State", font=("Academy Engraved", 10), bg="#dcdad5")
state_label.grid(row=2, column=4, padx=10, pady=10)
state_entry = ttk.Entry(data_frame)
state_entry.grid(row=2, column=5, padx=10, pady=10)

zipcode_label = Label(data_frame, text="Zipcode", font=("Academy Engraved", 10), bg="#dcdad5")
zipcode_label.grid(row=2, column=6, padx=10, pady=10)
zipcode_entry = ttk.Entry(data_frame)
zipcode_entry.grid(row=2, column=7, padx=10, pady=10)

cp_label = Label(data_frame, text="Company", font=("Academy Engraved", 10), bg="#dcdad5")
cp_label.grid(row=4, column=0, padx=10, pady=10)
cp_entry = ttk.Entry(data_frame)
cp_entry.grid(row=4, column=1, padx=10, pady=10)

jt_label = Label(data_frame, text="Job Title", font=("Academy Engraved", 10), bg="#dcdad5")
jt_label.grid(row=4, column=2, padx=10, pady=10)
jt_entry = ttk.Entry(data_frame)
jt_entry.grid(row=4, column=3, padx=10, pady=10)

dp_label = Label(data_frame, text="Department", font=("Academy Engraved", 10), bg="#dcdad5")
dp_label.grid(row=4, column=4, padx=10, pady=10)
dp_entry = ttk.Entry(data_frame)
dp_entry.grid(row=4, column=5, padx=10, pady=10)

cid_label = Label(data_frame, text="Custom Identifier", font=("Academy Engraved", 10), bg="#dcdad5")
#cid_label.grid(row=5, column=0, padx=10, pady=10)
cid_entry = ttk.Entry(data_frame)
#cid_entry.grid(row=5, column=1, padx=10, pady=10)

def remove():
    selected_items = my_tree.selection()

    if not selected_items:
        messagebox.showwarning("No Selection", "Please select one or more records to delete.")
        return

    try:
        # Get the selected database from the ComboBox
        selected_db = db_combobox.get()

        if not selected_db:
            messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
            return

        # Create a connection to the selected database
        conn = sqlite3.connect(selected_db)
        c = conn.cursor()

        for item in selected_items:
            # Get the row_id from the selected item
            row_id = my_tree.item(item, "values")[0]

            # Delete the record from the selected database
            c.execute("DELETE FROM customers WHERE id=?", (row_id,))

            # Delete the selected item from the Treeview
            my_tree.delete(item)

        # Commit changes and close the connection
        conn.commit()
        conn.close()

        # Clear the entry boxes
        clear_entries()

        # Show a message indicating successful deletion
        messagebox.showinfo("Deleted!", "Selected Records Have Been Deleted!")

    except sqlite3.Error as e:
        # Handle any database-related errors here
        messagebox.showerror("Error", f"Error deleting record(s): {e}")

# Remove all records
def remove_all():
    # Add a little message box for fun
    response = messagebox.askyesno("WOAH!!!!", "This Will Delete EVERYTHING From The Table\nAre You Sure?!")

    # Add logic for the message box
    if response == 1:
        # Clear the Treeview
        for record in my_tree.get_children():
            my_tree.delete(record)

        try:
            # Get the currently selected database from the ComboBox
            selected_db = db_combobox.get()

            if selected_db:
                # Connect to the selected database
                conn = sqlite3.connect(selected_db)
                c = conn.cursor()

                # Drop the 'customers' table
                c.execute("DROP TABLE IF EXISTS customers")

                # Recreate the 'customers' table
                c.execute("""CREATE TABLE IF NOT EXISTS customers (
                    id INTEGER PRIMARY KEY,
                    last_name TEXT,
                    first_name TEXT,
                    display_name TEXT,
                    description TEXT,
                    email TEXT,
                    sam TEXT,
                    upn TEXT,
                    address TEXT,
                    city TEXT,
                    state TEXT,
                    zipcode TEXT,
                    company TEXT,
                    job_title TEXT,
                    department TEXT,
                    cid TEXT
                )""")

                # Commit changes
                conn.commit()

                # Close the connection
                conn.close()

                # Clear entry boxes if filled
                clear_entries()
        except Exception as e:
            messagebox.showerror("Error", f"Error removing all records: {e}")

# Clear entry boxes
def clear_entries():
    # Clear entry boxes
    id_entry.delete(0, END)
    ln_entry.delete(0, END)
    fn_entry.delete(0, END)
    dn_entry.delete(0, END)
    dc_entry.delete(0, END)
    em_entry.delete(0, END)
    sm_entry.delete(0, END)
    up_entry.delete(0, END)
    address_entry.delete(0, END)
    city_entry.delete(0, END)
    state_entry.delete(0, END)
    zipcode_entry.delete(0, END)
    cp_entry.delete(0, END)
    jt_entry.delete(0, END)
    dp_entry.delete(0, END)
    cid_entry.delete(0, END)

# Select Record
def select_record(e):
    # Clear entry boxes
    id_entry.delete(0, END)
    ln_entry.delete(0, END)
    fn_entry.delete(0, END)
    dn_entry.delete(0, END)
    dc_entry.delete(0, END)
    em_entry.delete(0, END)
    sm_entry.delete(0, END)
    up_entry.delete(0, END)
    address_entry.delete(0, END)
    city_entry.delete(0, END)
    state_entry.delete(0, END)
    zipcode_entry.delete(0, END)
    cp_entry.delete(0, END)
    jt_entry.delete(0, END)
    dp_entry.delete(0, END)
    cid_entry.delete(0, END)

    # Grab record Number
    selected = my_tree.focus()
    # Grab record values
    values = my_tree.item(selected, 'values')

    # output to entry boxes
    id_entry.insert(0, values[0])
    ln_entry.insert(0, values[1])
    fn_entry.insert(0, values[2])
    dn_entry.insert(0, values[3])
    dc_entry.insert(0, values[4])
    em_entry.insert(0, values[5])
    sm_entry.insert(0, values[6])
    up_entry.insert(0, values[7])
    address_entry.insert(0, values[8])
    city_entry.insert(0, values[9])
    state_entry.insert(0, values[10])
    zipcode_entry.insert(0, values[11])
    cp_entry.insert(0, values[12])
    jt_entry.insert(0, values[13])
    dp_entry.insert(0, values[14])
    cid_entry.insert(0, values[15])

def update_record():
    # Grab the record number
    selected = my_tree.focus()

    if not selected:
        messagebox.showwarning("No Selection", "Please select a record to update.")
        return

    try:
        # Get the selected database from the ComboBox
        selected_db = db_combobox.get()

        if not selected_db:
            messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
            return

        # Get the original ID of the selected record
        original_id = my_tree.item(selected, 'values')[0]

        # Update record in the Treeview
        my_tree.item(selected, text="", values=(
            original_id,  # Keep the original ID
            ln_entry.get(), fn_entry.get(), dn_entry.get(), dc_entry.get(), em_entry.get(),
            sm_entry.get(), up_entry.get(), address_entry.get(), city_entry.get(), state_entry.get(),
            zipcode_entry.get(), cp_entry.get(), jt_entry.get(), dp_entry.get(), cid_entry.get()))

        # Create a connection to the selected database
        with sqlite3.connect(selected_db) as conn:
            c = conn.cursor()

            # Update the record in the selected database excluding the "ID" column
            c.execute("""UPDATE customers SET
                last_name = ?,
                first_name = ?,
                display_name = ?,
                description = ?,
                email = ?,
                sam = ?,
                upn = ?,
                address = ?,
                city = ?,
                state = ?,
                zipcode = ?,
                company = ?,
                job_title = ?,
                department = ?,
                cid = ?
                WHERE id = ?""",
                (
                    ln_entry.get(),
                    fn_entry.get(),
                    dn_entry.get(),
                    dc_entry.get(),
                    em_entry.get(),
                    sm_entry.get(),
                    up_entry.get(),
                    address_entry.get(),
                    city_entry.get(),
                    state_entry.get(),
                    zipcode_entry.get(),
                    cp_entry.get(),
                    jt_entry.get(),
                    dp_entry.get(),
                    cid_entry.get(),
                    original_id  # Use the original ID for the WHERE clause
                ))

            # Commit changes (not needed when using the 'with' context manager)
            # conn.commit()
            clear_entries()

    except sqlite3.Error as e:
        # Handle any database-related errors here
        messagebox.showerror("Error", f"Error updating record: {e}")


def add_record():
    try:
        # Get the currently selected database from the ComboBox
        selected_db = db_combobox.get()

        if selected_db:
            # Connect to the selected database
            conn = sqlite3.connect(selected_db)
            c = conn.cursor()

            # Get the next available ID from the database
            c.execute("SELECT MAX(id) FROM customers")
            max_id = c.fetchone()[0]
            next_id = max_id + 1 if max_id else 1

            # Retrieve the first name and last name entries
            first_name = fn_entry.get().strip()
            last_name = ln_entry.get().strip()

            # Create the cid without spaces, or replace spaces with another character if preferred
            cid_value = f"{first_name.lower()}{last_name.lower()}"
            cid_value = cid_value.replace(" ", "")  # Remove spaces for cid
            cid_value = cid_value[:16]  # Truncate CID to 16 characters

            # Add two random capital letters to ensure uniqueness
            cid_value += ''.join(random.choices(string.ascii_uppercase, k=2))

            # Add New Record (including 'cid' field with proper handling for single word requirement)
            c.execute(
                "INSERT INTO customers (id, last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (next_id, last_name, first_name, dn_entry.get(), dc_entry.get(), em_entry.get(), sm_entry.get(),
                 up_entry.get(), address_entry.get(), city_entry.get(), state_entry.get(), zipcode_entry.get(),
                 cp_entry.get(), jt_entry.get(), dp_entry.get(), cid_value))

            # Commit changes and close the connection
            conn.commit()
            conn.close()

            clear_entries()

            # Optionally, refresh the displayed records in the Treeview
            my_tree.delete(*my_tree.get_children())
            query_database()
    except Exception as e:
        messagebox.showerror("Error", f"Error adding record: {e}")

def force_uppercase(*args): # Function to force uppercase
    last_name = last_name_var.get().upper()
    last_name_var.set(last_name)

last_name_var = StringVar()
last_name_var.trace_add("write", force_uppercase)
ln_entry = ttk.Entry(data_frame, textvariable=last_name_var)
ln_entry.grid(row=0, column=3, padx=10, pady=10)


def update_display_name():
    # Get the original names from entry fields and immediately define modified versions
    original_first_name = fn_entry.get().strip()
    original_last_name = ln_entry.get().strip()

    # Predefine modified names outside the if block to ensure they are always available
    modified_first_name = original_first_name.replace(" ", "-").lower()
    modified_last_name = original_last_name.replace(" ", "-").lower()

    if original_first_name and original_last_name:
        # Determine the email provider/domain
        email_domain = "@example.com"  # Default domain; adjust as needed
        if "@" in em_entry.get():
            email_parts = em_entry.get().split("@")
            email_domain = "@" + email_parts[-1] if len(email_parts) > 1 else email_domain

        # Construct email, SAM, and UPN values
        email_var = f"{modified_first_name}.{modified_last_name}{email_domain}"
        sam_var = f"{modified_first_name}.{modified_last_name}"
        upn_var = sam_var

        # Enforce length restrictions for SAM
        if len(sam_var) > 20:  # Adjust the length limit as necessary
            sam_var = 'Too long, please shorten'

        # Update the respective fields with the new values
        em_entry.delete(0, 'end')
        em_entry.insert(0, email_var)
        sm_entry.delete(0, 'end')
        sm_entry.insert(0, sam_var)
        up_entry.delete(0, 'end')
        up_entry.insert(0, upn_var)
    else:
        # Clear the fields if either name is empty
        em_entry.delete(0, 'end')
        sm_entry.delete(0, 'end')
        up_entry.delete(0, 'end')

    # Update the display name using the original names
    display_name = f"{original_last_name.upper()} {original_first_name}".strip()
    display_name_var.set(display_name)

    # Set additional fields to "None" if empty
    fields_to_check = [dc_entry, address_entry, city_entry, state_entry, zipcode_entry, cp_entry, jt_entry, dp_entry]
    for field in fields_to_check:
        if not field.get().strip():
            field.delete(0, 'end')
            field.insert(0, "None")

    # Handle CID value
    cid_var = f"{modified_first_name}{modified_last_name}"
    cid_var = cid_var[:16]  # Truncate CID to 16 characters
    # Add two random capital letters
    cid_var += ''.join(random.choices(string.ascii_uppercase, k=2))

    cid_entry.delete(0, 'end')
    cid_entry.insert(0, cid_var)

# Call the function
update_display_name()

def set_default_if_empty(event, entry_field):
    """Set the entry field to 'None' if it is empty when focus is lost."""
    if not entry_field.get().strip():
        entry_field.delete(0, END)
        entry_field.insert(0, "None")
def load_csv_data(csv_filename):
    try:
        with open(csv_filename, 'r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Populate the existing Entry widgets with CSV data
                cp_entry.delete(0, END)
                cp_entry.insert(0, row['Company'])
                address_entry.delete(0, END)
                address_entry.insert(0, row['Address'])
                city_entry.delete(0, END)
                city_entry.insert(0, row['City'])
                state_entry.delete(0, END)
                state_entry.insert(0, row['State'])
                zipcode_entry.delete(0, END)
                zipcode_entry.insert(0, row['Zipcode'])
    except FileNotFoundError:
        messagebox.showerror("Error", "CSV file not found.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def update_email_domain(csv_filename):
    try:
        with open(csv_filename, 'r') as file:
            reader = csv.DictReader(file)

            for row in reader:
                csv_part = row['Status']  # Assume 'Status' is the part to append

                if csv_part:
                    # Get original names, replace spaces for compound names
                    original_first_name = fn_entry.get().strip()
                    original_last_name = ln_entry.get().strip()
                    modified_first_name = original_first_name.replace(" ", "-").lower()
                    modified_last_name = original_last_name.replace(" ", "-").lower()

                    # Construct the base part of the email, SAM, and UPN using modified names
                    base_part = f"{modified_first_name}.{modified_last_name}"

                    # Update Email
                    updated_email = f"{base_part}{csv_part}"
                    em_entry.delete(0, 'end')
                    em_entry.insert(0, updated_email)

                    # Update SAM (use only the part up to the '@' character in csv_part if present)
                    sam_part = csv_part.split('@')[0] if '@' in csv_part else csv_part
                    updated_sam = f"{base_part}{sam_part}"
                    sm_entry.delete(0, 'end')
                    sm_entry.insert(0, updated_sam)

                    # Update UPN to use the full csv_part
                    updated_upn = f"{base_part}{csv_part}"
                    up_entry.delete(0, 'end')
                    up_entry.insert(0, updated_upn)

                else:
                    messagebox.showerror("Error", "No text to append from CSV.")
                break  # Assuming you want to update based on the first row only

    except FileNotFoundError:
        messagebox.showerror("Error", "CSV file not found.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def list_csv_files(directory):
    """List all .csv files in a given directory."""
    return [f for f in os.listdir(directory) if f.endswith('.csv')]

filename_to_directory = {}  # Global mapping

def update_combobox_with_csv_files(combobox):
    global filename_to_directory
    filename_to_directory.clear()  # Clear previous entries
    csv_files = []
    directories = {
        '---Addresses---': 'addresses',
        '---Domain---': 'domain',
        '---Organization Unit---': 'organization_unit',
    }

    for separator, directory in directories.items():
        csv_files.append(separator)  # Add separator for visual grouping
        files = list_csv_files(directory)
        for file in files:
            filename_to_directory[file] = directory  # Map filename to directory
        csv_files.extend(files)  # Add files from directory

    combobox['values'] = csv_files

def create_edition_section():
    global ed_csv_combobox

    edition_section_frame = ttk.LabelFrame(tools_tab, text="Edition Section")
    edition_section_frame.grid(row=0, column=3, padx=10, pady=10, sticky='nw')
    edition_section_frame.grid_columnconfigure(0, weight=1)  # Make the column within the frame expandable

    ed_csv_combobox = ttk.Combobox(edition_section_frame, state='readonly', width=20)
    ed_csv_combobox.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
    ed_csv_combobox.bind('<<ComboboxSelected>>', lambda event: on_combobox_select(event, text_box))

    text_box = tk.Text(edition_section_frame, height=10, width=50)
    text_box.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

    update_combobox_with_csv_files(ed_csv_combobox)

    def save_changes():
        selected_file = ed_csv_combobox.get()
        directory = filename_to_directory.get(selected_file, None)  # Use the updated approach to get directory
        if directory:
            csv_filepath = os.path.join(directory, selected_file)
            content = text_box.get("1.0", "end-1c")  # Ensure this captures all text

            try:
                with open(csv_filepath, 'w') as file:
                    file.write(content)
                print("File saved successfully.")
            except Exception as e:
                print(f"Failed to save file: {e}")
        else:
            print(f"Could not find directory for file: {selected_file}")

    save_button = ttk.Button(edition_section_frame, text="Edit & Save", command=save_changes)
    save_button.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

def get_directory_from_selection(selected_file):
    """Return the directory based on the prefix in the selected file."""
    if '---Addresses---' in selected_file:
        return 'addresses'
    elif '---Domain---' in selected_file:
        return 'domain'
    elif '---Organization Unit---' in selected_file:
        return 'organization_unit'
    else:
        return None

last_valid_selection = None  # Keep track of the last valid selection

def on_combobox_select(event, text_widget):
    global last_valid_selection
    selected = event.widget.get()

    # Check if the selected item is a separator
    if selected.startswith('---'):
        # If a separator is selected, revert to the last valid selection (if any)
        event.widget.set(last_valid_selection if last_valid_selection else '')
        print("Separator selected, reverting to last valid selection.")
    else:
        # Update last valid selection if a valid file is selected
        last_valid_selection = selected
        directory = filename_to_directory.get(selected, None)
        if directory:
            filename = selected.split(': ')[-1] if ': ' in selected else selected
            csv_filepath = os.path.join(directory, filename)
            try:
                with open(csv_filepath, 'r') as file:
                    content = file.read()
                    text_widget.delete("1.0", "end")
                    text_widget.insert("1.0", content)
            except FileNotFoundError:
                print(f"File not found: {csv_filepath}")
        print(f"Directory not found for selected item: {selected}")

def populate_domain_csv_combobox():
    try:
        # domain_folder = os.path.join(os.path.dirname(__file__), "domain")
        # Get a list of CSV files in the "domain" folder
        csv_files = [f for f in os.listdir(domain_folder) if f.endswith('.csv')]
        # Set the values of the ComboBox to the list of CSV files
        domain_csv_combobox['values'] = csv_files
    except FileNotFoundError:
        messagebox.showerror("Error", "The 'domain' folder was not found.")

def populate_ou_csv_combobox():
    try:
        csv_files = [f for f in os.listdir(organization_unit_folder) if f.endswith('.csv')]
        # Set the values of the ComboBox to the list of CSV files
        ou_combobox['values'] = csv_files
    except FileNotFoundError:
        messagebox.showerror("Error", "The 'Organization unit' folder was not found.")

def on_csv_select(event):
    selected_csv = csv_combobox.get()
    if selected_csv:
        #script_dir = os.path.dirname(__file__)
        #addresses_folder = os.path.join(script_dir, "addresses")
        csv_filename = os.path.join(addresses_folder, selected_csv)

        if os.path.exists(csv_filename):
            load_csv_data(csv_filename)
        else:
            messagebox.showerror("Error", f"CSV file not found: {csv_filename}")

def on_domain_csv_select(event):
    selected_csv = domain_csv_combobox.get()
    if selected_csv:
        csv_filename = os.path.join("domain", selected_csv)  # Assuming CSV files are in the "domain" folder
        update_email_domain(csv_filename)

def on_ou_csv_select(event):
    selected_csv = ou_combobox.get()
    if selected_csv:
        csv_filename = os.path.join("organization_unit", selected_csv)  # Assuming CSV files are in the "Organization_unit folder" folder

def add_fake_record():
    try:
        # Create an instance of the Faker class
        fake = Faker()

        # Generate fake data
        last_name = fake.last_name().upper()
        first_name = fake.first_name()
        display_name = f"{last_name.upper()} {first_name}"
        description = "Employee"  # Set Description as "Employee"

        # Set email, UPN, and SAM based on first_name and last_name
        email = f"{first_name.lower()}.{last_name.lower()}@example.com"
        sam = f"{first_name.lower()}.{last_name.lower()}"
        upn = f"{first_name.lower()}.{last_name.lower()}"
        cid = f"cid.{first_name.lower()}.{last_name.lower()}"

        address = fake.street_address()
        city = fake.city()
        state = fake.state_abbr()
        zipcode = fake.zipcode()
        company = fake.company()
        job_title = fake.job()

        # Generate the department in capital letters
        department = fake.word().upper()

        # Get the next available ID from the selected database
        selected_db = db_combobox.get()

        if selected_db:
            conn = sqlite3.connect(selected_db)
            c = conn.cursor()
            c.execute("SELECT MAX(id) FROM customers")
            max_id = c.fetchone()[0]
            next_id = max_id + 1 if max_id else 1

            # Insert the fake record into the Treeview using the next available ID
            my_tree.insert(parent='', index='end', iid=next_id, text='',
                           values=(
                               next_id, last_name, first_name, display_name, description, email, sam, upn, address, city,
                               state, zipcode, company, job_title, department, cid),
                           tags=('oddrow' if next_id % 2 != 0 else 'evenrow'))

            # Insert the fake record into the selected database
            c.execute(
                "INSERT INTO customers (id, last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (next_id, last_name, first_name, display_name, description, email, sam, upn, address, city, state,
                 zipcode, company, job_title, department, cid))
            conn.commit()
            conn.close()
        else:
            messagebox.showwarning("Warning", "No database selected. Please select a database.")
    except Exception as e:
        messagebox.showerror("Error", f"Error adding fake record: {e}")

def add_multiple_fake_records():
    # Get the selected database from the ComboBox
    selected_db = db_combobox.get()
    for _ in range(50):
        add_fake_record()

        # Create a database or connect to the selected one
        conn = sqlite3.connect(selected_db)
        c = conn.cursor()
        # Commit changes
        conn.commit()
        # Close our connection
        conn.close()

def sync_records_to_ad():
    try:
        # Get the directory where the Python script is located
        script_dir = os.path.dirname(sys.argv[0])

        # Construct the full path to the PowerShell script
        script_path = os.path.join(script_dir, "ada_users_server.ps1")

        # Get the selected OU path from the combobox
        selected_csv = ou_combobox.get()
        if not selected_csv or selected_csv == 'Select Organization Unit':
            print("Please select a valid Organization Unit from the combobox.")
            return

        # Assuming the CSV has only one line which is the OU path
        csv_path = os.path.join(script_dir, "organization_unit", selected_csv)
        with open(csv_path, 'r') as f:
            ou_path = f.readline().strip()

        # Get the selected database from the combobox
        selected_db = db_combobox.get()
        if not selected_db or selected_db == 'Select DataBase':
            print("Please select a valid database from the combobox.")
            return

        # Run the PowerShell script using subprocess and pass the OU path and selected DB as arguments
        subprocess.run(["powershell.exe", "-File", script_path, ou_path, selected_db], shell=True, check=True)
        print("Sync Database to AD completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error running PowerShell script: {e}")

def export_to_csv():
    # Get the selected database from the ComboBox
    selected_db = db_combobox.get()

    if not selected_db:
        messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
        return

    try:
        # Extract the database name without the extension
        db_name = selected_db.split('.')[0]

        # Get the current date and time in yyyymmdd-hhmmss format
        current_datetime = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")

        # Form the CSV file name
        csv_filename = f"{db_name}_{current_datetime}.csv"

        # Open a CSV file for writing with the new file name
        with open(csv_filename, 'w', newline='') as csv_file:
            # Create a CSV writer
            csv_writer = csv.writer(csv_file)

            # Write the header row
            header = ["ID", "Last Name", "First Name", "Display Name", "Description", "Email", "SAM", "UPN", "Address",
                      "City", "State", "Zipcode", "Company", "Job Title", "Department", "Custom Identifier"]
            csv_writer.writerow(header)

            # Connect to the selected database
            conn = sqlite3.connect(selected_db)
            c = conn.cursor()

            # Fetch all records from the database
            c.execute("SELECT * FROM customers")
            records = c.fetchall()

            # Write each record to the CSV file
            for record in records:
                csv_writer.writerow(record)

            # Close the database connection
            conn.close()

        messagebox.showinfo("Export Complete", f"All records from {selected_db} have been exported to {csv_filename}")
    except sqlite3.Error as e:
        # Handle any database-related errors here
        messagebox.showerror("Error", f"Error exporting records: {e}")
def import_csv_data():
    # Ask the user to select a CSV file for import
    csv_filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

    if csv_filename:
        try:
            # Get the selected database from the ComboBox
            selected_db = db_combobox.get()

            if not selected_db:
                messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
                return

            # Open the selected CSV file and read its contents
            with open(csv_filename, 'r') as file:
                csv_reader = csv.reader(file)

                # Skip the header row
                next(csv_reader, None)

                # Create a database connection for the selected database
                conn = sqlite3.connect(selected_db)
                cursor = conn.cursor()

                # Get the next available ID in the selected database
                cursor.execute("SELECT MAX(id) FROM customers")
                max_id = cursor.fetchone()[0]
                next_id = max_id + 1 if max_id else 1

                # Iterate over the CSV rows and insert them into the selected database
                for row in csv_reader:
                    # Ensure that the row has all 15 values (add empty strings for missing values)
                    while len(row) < 15:
                        row.append('')

                    cursor.execute(
                        "INSERT INTO customers (id, last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        [next_id] + row[1:])  # Start from the second column (index 1)

                    next_id += 1  # Increment the ID for the next record

                # Commit changes and close the database connection
                conn.commit()
                conn.close()

                # Refresh the displayed records in the Treeview
                query_database()

                messagebox.showinfo("Success", f"CSV data imported successfully into {selected_db}.")
        except FileNotFoundError:
            messagebox.showerror("Error", "CSV file not found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def insert_data_into_database(data_rows):
    try:
        # Get the selected database from the ComboBox
        selected_db = db_combobox.get()

        if not selected_db:
            messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
            return

        # Create a database connection for the selected database
        conn = sqlite3.connect(selected_db)
        cursor = conn.cursor()

        # Insert the data into the selected database
        for row in data_rows:
            # Ensure that the row has all 15 values (add empty strings for missing values)
            while len(row) < 15:
                row.append('')

            cursor.execute(
                "INSERT INTO customers (id, last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                row)

        # Commit changes and close the database connection
        conn.commit()
        conn.close()

        # Refresh the displayed records in the TreeView
        query_database()

        messagebox.showinfo("Success", f"Data inserted successfully into {selected_db}.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def delete_duplicates():
    # Get the selected database from the ComboBox
    selected_db = db_combobox.get()

    if not selected_db:
        messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
        return

    try:
        # Create a database connection
        conn = sqlite3.connect(selected_db)
        c = conn.cursor()

        # Identify duplicates based on a unique criteria, e.g., email
        c.execute("""
            DELETE FROM customers
            WHERE id NOT IN (
                SELECT MIN(id)
                FROM customers
                GROUP BY email  -- Change 'email' to the column you want to check for duplicates
            )
        """)

        # Commit changes
        conn.commit()

        # Close the connection
        conn.close()

        # Clear The Treeview Table
        my_tree.delete(*my_tree.get_children())

        # Refresh the displayed records in the Treeview
        query_database()

        messagebox.showinfo("Duplicates Deleted", "Duplicate records have been deleted.")
    except sqlite3.Error as e:
        # Handle any database-related errors here
        messagebox.showerror("Error", f"Error deleting duplicates: {e}")

def confirm_delete_duplicates():
    result = messagebox.askquestion("Confirm Deletion", "Are you sure you want to delete duplicate records?", icon='warning')
    if result == 'yes':
        delete_duplicates()
    else:
        messagebox.showinfo("Cancelled", "Deletion of duplicate records cancelled.")

# Function to reassign IDs in sequential order
def reassign_ids():
    # Get the selected database from the ComboBox
    selected_db = db_combobox.get()

    if not selected_db:
        messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
        return

    try:
        # Create a database connection
        conn = sqlite3.connect(selected_db)
        c = conn.cursor()

        # Retrieve all records ordered by the current ID
        c.execute("SELECT * FROM customers ORDER BY id")

        # Fetch all records
        records = c.fetchall()

        # Close the current cursor
        c.close()

        # Open a new cursor
        c = conn.cursor()

        # Initialize a new ID value
        new_id = 1

        # Update the ID of each record to be sequential
        for record in records:
            current_id = record[0]
            c.execute("UPDATE customers SET id = ? WHERE id = ?", (new_id, current_id))
            new_id += 1

        # Commit changes
        conn.commit()

        # Close the connection
        conn.close()

        # Clear The Treeview Table
        my_tree.delete(*my_tree.get_children())

        # Refresh the displayed records in the Treeview
        query_database()
    except sqlite3.Error as e:
        # Handle any database-related errors here
        messagebox.showerror("Error", f"Error reassigning IDs: {e}")


def backup_database():
    try:
        # Get the selected database filename from the ComboBox
        selected_db = db_combobox.get()

        if not selected_db:
            messagebox.showwarning("No Database Selected", "Please select a database from the ComboBox.")
            return

        # Generate a timestamp for the backup file name with date and time separated
        timestamp = time.strftime("%Y%m%d-%H%M%S")

        # Extract the original database file name and extension
        db_name = os.path.basename(selected_db)
        base_name, extension = os.path.splitext(db_name)

        # Specify the backup file name and path
        backup_filename = f"{base_name}_{timestamp}{extension}"

        # Create a backup by copying the selected database file
        shutil.copyfile(selected_db, backup_filename)

        messagebox.showinfo("Backup Complete", f"Backup saved as {backup_filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Error creating backup: {e}")

def browse_db():
    global db_combobox, my_tree  # Assuming db_combobox and my_tree are defined globally

    def get_application_path():
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    db_file = filedialog.askopenfilename(filetypes=[("SQLite Database Files", "*.db")])
    if db_file:
        script_dir = get_application_path()
        new_db_path = os.path.join(script_dir, os.path.basename(db_file))

        if os.path.dirname(db_file) != script_dir:
            try:
                shutil.copy(db_file, new_db_path)
            except Exception as e:
                print(f"Error copying database: {e}")
                return

        try:
            db_combobox['values'] = tuple(os.path.basename(f) for f in glob.glob(os.path.join(script_dir, "*.db")))
            db_combobox.set(os.path.basename(new_db_path))

            # Call load_selected_db to load and display the database
            load_selected_db(None)
        except Exception as e:
            print(f"Error loading new database: {e}")

def edit_sync_script():
    def find_text(event=None):
        search_query = search_entry.get()

        if search_query:
            start_pos = "1.0"
            script_text.tag_remove("highlight", "1.0", END)  # Clear any previous highlights

            while True:
                found_pos = script_text.search(search_query, start_pos, nocase=1, stopindex=END)

                if not found_pos:
                    break

                end_pos = f"{found_pos}+{len(search_query)}c"
                script_text.tag_add("highlight", found_pos, end_pos)
                start_pos = end_pos

            script_text.tag_config("highlight", background="red")
            script_text.see(INSERT)

        script_text.tag_remove(SEL, "1.0", END)

    def next_result(event=None):
        find_text()
        script_text.tag_remove(SEL, "1.0", END)
        script_text.tag_add(SEL, "1.0", script_text.search(INSERT, "1.0+1c", stopindex=END))

    def prev_result(event=None):
        find_text()
        script_text.tag_remove(SEL, "1.0", END)
        script_text.tag_add(SEL, "1.0", script_text.search(INSERT, "1.0-1c", stopindex="1.0"))

    # Create a new window for script editing
    script_editor = Toplevel(root)
    script_editor.title("Edit Sync Database Script")
    script_editor.iconbitmap("icons.ico")

    # Create a custom font with bold
    custom_font = font.Font(size=10, weight="bold")

    # Create a Text widget for script editing with vertical and horizontal scrollbars
    script_text = ScrolledText(script_editor, wrap=WORD, width=80, height=30, font=custom_font)
    script_text.pack(padx=10, pady=10, fill=BOTH, expand=True)

    # Set the background color to blue (#0000FF) and text color to yellow (#FFFF00)
    script_text.configure(bg="#012456", fg="#eeefeb")

    # Create an entry widget for entering search queries
    search_entry = ttk.Entry(script_editor)
    search_entry.pack(padx=10, pady=5, fill=X)

    # Create a "Find" button to initiate the search
    find_button = ttk.Button(script_editor, text="Find", command=find_text)
    find_button.pack(padx=10, pady=5)

    search_entry.bind("<Return>", find_text)

    # Bind arrow keys for navigation
    script_text.bind("<Up>", prev_result)
    script_text.bind("<Down>", next_result)

    # Load the script from the existing PowerShell file
    with open("ada_users_server.ps1", "r") as script_file:
        script_content = script_file.read()
        script_text.insert("1.0", script_content)

    def save_script():
        # Save the edited script back to the file
        edited_script = script_text.get("1.0", END)
        with open("ada_users_server.ps1", "w") as script_file:
            script_file.write(edited_script)

        messagebox.showinfo("Script Saved", "The script has been saved.")

    def close_script_editor():
        script_editor.destroy()

    # Create buttons for saving and closing
    save_button = ttk.Button(script_editor, text="Save Script", command=save_script)
    save_button.pack(side=LEFT, padx=10, pady=10)

    close_button = ttk.Button(script_editor, text="Close", command=close_script_editor)
    close_button.pack(side=RIGHT, padx=10, pady=10)

def create_address_csv():
    # Create a LabelFrame in the Tools tab without making it expand to fill the tab
    address_frame = ttk.LabelFrame(tools_tab, text="Create Address")
    address_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')  # Align to north-west without expanding

    # Define the fields
    fields = ['CSV Filename', 'Company', 'Address', 'City', 'State', 'Zipcode']
    entries = {}

    for i, field in enumerate(fields):
        # Create label for each field
        label = ttk.Label(address_frame, text=field)
        label.grid(row=i, column=0, padx=5, pady=5, sticky='w')

        # Create entry field for each label with a specific width
        entry = ttk.Entry(address_frame, width=25)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
        entries[field] = entry

    def save_to_csv():
        # Get values from the entry fields
        data = {field: entries[field].get() for field in fields[1:]}  # Exclude CSV Filename from data dict

        # Get the CSV filename from the entry field
        csv_filename = entries['CSV Filename'].get() + '.csv'

        # Specify the CSV file path (you can change the directory as needed)
        csv_directory = 'addresses'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, csv_filename)

        # Write data to the CSV file
        with open(csv_filepath, 'w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(fields[1:])  # Write header row (exclude CSV Filename)
            csv_writer.writerow(list(data.values()))  # Write data row

        populate_csv_combobox()
        update_combobox_with_csv_files(ed_csv_combobox)

    # Create Save button just below the sixth field
    save_button = ttk.Button(address_frame, text="Save Custom Address", command=save_to_csv)
    save_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

def create_sam_at_domain_csv():
    # Create a LabelFrame in the Tools tab for the sAM@domain CSV
    sam_at_domain_frame = ttk.LabelFrame(tools_tab, text="Create sAM@domain CSV")
    # Position it accordingly in the grid within the tools_tab
    sam_at_domain_frame.grid(row=2, column=0, padx=10, pady=10, sticky='nw')

    # Define labels and entry fields
    fields = ['CSV File Name', 'Status']
    entries = {}

    for i, field in enumerate(fields):
        # Create label for each field
        label = ttk.Label(sam_at_domain_frame, text=field)
        label.grid(row=i, column=0, padx=5, pady=5, sticky='w')

        # Create entry field for each label with width 20
        entry = ttk.Entry(sam_at_domain_frame, width=20)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
        entries[field] = entry

    # Adjusted variable names for clarity
    csv_filename_entry = entries['CSV File Name']
    status_entry = entries['Status']

    def save_to_csv():
        # Get the values from the entry fields
        status_value = status_entry.get()
        csv_filename = csv_filename_entry.get()

        # Specify the CSV file path
        csv_directory = 'domain'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, f'{csv_filename}.csv')

        # Write data to the CSV file
        with open(csv_filepath, 'w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Status'])
            csv_writer.writerow([status_value])

        populate_domain_csv_combobox()
        update_combobox_with_csv_files(ed_csv_combobox)

    # Create a button to save the data to CSV
    save_button = ttk.Button(sam_at_domain_frame, text="Save Custom sAM", command=save_to_csv)
    save_button.grid(row=len(fields), column=0, columnspan=2, padx=10, pady=10, sticky='ew')
def create_ou_csv():
    # Create a LabelFrame in the Tools tab for the Organization Unit CSV
    ou_frame = ttk.LabelFrame(tools_tab, text="Create Organization Unit CSV")
    # Position it to the right of the first LabelFrame if using grid, or adjust accordingly
    ou_frame.grid(row=1, column=0, padx=10, pady=10, sticky='nw')

    # Define labels and entry fields
    fields = ['CSV File Name', 'Organization Unit']
    entries = {}

    for i, field in enumerate(fields):
        # Create label for each field
        label = ttk.Label(ou_frame, text=field)
        label.grid(row=i, column=0, padx=5, pady=5, sticky='w')

        # Create entry field for each label with width 20
        entry = ttk.Entry(ou_frame, width=20)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
        entries[field] = entry

    # Adjusted variable names to reflect their purpose clearly
    csv_filename_entry = entries['CSV File Name']
    organization_unit_entry = entries['Organization Unit']

    def save_to_csv():
        # Get the values from the entry fields
        organization_unit_value = organization_unit_entry.get()
        csv_filename = csv_filename_entry.get()

        # Specify the CSV file path
        csv_directory = 'organization_unit'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, f'{csv_filename}.csv')

        # Write data directly to the file
        with open(csv_filepath, 'w') as csvfile:
            csvfile.write(organization_unit_value + '\n')

        populate_ou_csv_combobox()
        update_combobox_with_csv_files(ed_csv_combobox)

    # Create a button to save the data to CSV
    save_button = ttk.Button(ou_frame, text="Save Organization Unit Path", command=save_to_csv)
    save_button.grid(row=len(fields), column=0, columnspan=2, padx=10, pady=10, sticky='ew')

def enter_new_password():
    # Create a TopLevel window for entering the new password
    password_window = Toplevel(root)
    password_window.title("Enter New Password")
    password_window.iconbitmap("icons.ico")

    new_password_label = Label(password_window, text="New Password:")
    new_password_label.pack()

    new_password_entry = ttk.Entry(password_window, show="*")
    new_password_entry.pack()

    def save_password():
        new_password = new_password_entry.get()
        password_window.destroy()
        change_password(new_password)

    save_button = ttk.Button(password_window, text="Save Password", command=save_password)
    save_button.pack()

def change_password(new_password):
    # Iterate over selected rows and call the PowerShell script for each user
    for item in my_tree.selection():
        custom_identifier = my_tree.item(item, 'values')[15]  # Assuming 'Custom Identifier' is in the 15th column

        if not custom_identifier:
            print(f"Skipping row {item}: Custom Identifier not found.")
            continue

        # Quote the new password to handle special characters
        quoted_new_password = shlex.quote(new_password)

        # Construct the PowerShell command to execute the script
        ps_command = [
            'powershell.exe',
            '-File',
            'change_password.ps1',
            '-customIdentifier',
            custom_identifier,
            '-newPassword',
            quoted_new_password
        ]

        # Execute the PowerShell script
        try:
            result = subprocess.run(ps_command, capture_output=True, text=True, check=True)
            print(result.stdout)
        except subprocess.CalledProcessError as e:
            print(f"Error changing password for Custom Identifier '{custom_identifier}': {e}")

    print("Password change completed.")

def toggle_selected_users():
    # Iterate over selected rows and call the PowerShell script for each user
    for item in my_tree.selection():
        custom_identifier = my_tree.item(item, 'values')[15]  # Assuming 'Custom Identifier' is in the 15th column

        if not custom_identifier:
            print(f"Skipping row {item}: Custom Identifier not found.")
            continue

        # Construct the PowerShell command to execute the script
        ps_command = [
            'powershell.exe',
            '-File',
            'toggle_ad_account.ps1',  # The name of your PowerShell script
            '-customIdentifier',
            custom_identifier
        ]

        # Execute the PowerShell script
        try:
            result = subprocess.run(ps_command, capture_output=True, text=True, check=True)
            print(result.stdout)
        except subprocess.CalledProcessError as e:
            print(f"Error toggling account for Custom Identifier '{custom_identifier}': {e}")

    print("Account toggle operation completed.")

def get_OU_and_extract():
    # Get the selected OU path CSV from the combobox
    selected_csv = ou_combobox.get()

    # Check if the combobox is empty or has the default value
    if not selected_csv or selected_csv == 'Select Organization Unit':
        messagebox.showwarning("Warning", "Please select a valid OU from Organization Unit Management.")
        return

    # Construct the path to the CSV inside the 'Organization_unit' folder
    full_csv_path = os.path.join('organization_unit', selected_csv)

    # Read the OU path from the CSV
    try:
        with open(full_csv_path, 'r') as f:
            ou_path = f.readline().strip()
    except Exception as e:
        print(f"Error reading OU from CSV: {e}")
        return

    # Surround the OU with quotation marks
    OU_quoted = f'"{ou_path}"'

    # Now you can pass the quoted OU value to the PowerShell script and execute it
    extract_ad_to_csv(OU_quoted)

def show_user_ou():
    selected_items = my_tree.selection()
    if len(selected_items) != 1:
        messagebox.showerror("Error", "Please select a single user.")
        return

    custom_identifier = my_tree.item(selected_items[0], 'values')[15]  # Assuming 'Custom Identifier' is in the 15th column

    ps_command = [
        'powershell.exe',
        '-File',
        'get_user_ou.ps1',
        '-customIdentifier',
        custom_identifier
    ]

    try:
        result = subprocess.run(ps_command, capture_output=True, text=True, check=True)
        ou = result.stdout.strip()
        messagebox.showinfo("User OU", f"The user is in OU: {ou}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Error retrieving OU for Custom Identifier '{custom_identifier}': {e}")

def copy_sam_to_employeeid():
    selected_items = my_tree.selection()
    if not selected_items:
        messagebox.showerror("Error", "Please select at least one user.")
        return

    success_count = 0
    error_messages = []

    for item in selected_items:
        custom_identifier = my_tree.item(item, 'values')[15]  # Assuming 'Custom Identifier' is in the 15th column

        ps_command = [
            'powershell.exe',
            '-File',
            'copy_sam_to_employeeid.ps1',
            '-customIdentifier',
            custom_identifier
        ]

        try:
            subprocess.run(ps_command, capture_output=True, text=True, check=True)
            success_count += 1
        except subprocess.CalledProcessError as e:
            error_messages.append(f"Error for '{custom_identifier}': {e}")

    # Display results
    if success_count == len(selected_items):
        messagebox.showinfo("Operation Completed", "SAMAccountName copied to employeeID for all selected users.")
    elif success_count > 0:
        messagebox.showinfo("Partial Completion",
                            f"SAMAccountName copied to employeeID for {success_count} out of {len(selected_items)} selected users.")

    if error_messages:
        messagebox.showerror("Errors Encountered", "\n".join(error_messages))


def erase_employeeid_for_multiple_users():
    selected_items = my_tree.selection()
    if not selected_items:
        messagebox.showerror("Error", "Please select at least one user.")
        return

    success_count = 0
    error_messages = []

    for item in selected_items:
        custom_identifier = my_tree.item(item, 'values')[15]  # Assuming 'Custom Identifier' is in the 15th column

        ps_command = [
            'powershell.exe',
            '-File',
            'erase_employeeid.ps1',
            '-customIdentifier',
            custom_identifier
        ]

        try:
            subprocess.run(ps_command, capture_output=True, text=True, check=True)
            success_count += 1
        except subprocess.CalledProcessError as e:
            error_messages.append(f"Error for '{custom_identifier}': {e}")

    # Display results
    if success_count == len(selected_items):
        messagebox.showinfo("Operation Completed", "employeeID erased for all selected users.")
    elif success_count > 0:
        messagebox.showinfo("Partial Completion",
                            f"employeeID erased for {success_count} out of {len(selected_items)} selected users.")

    if error_messages:
        messagebox.showerror("Errors Encountered", "\n".join(error_messages))

def extract_ad_to_csv(OU):
    # Modify the PowerShell script invocation to pass the OU as an argument
    process = subprocess.Popen(["powershell", ".\extract_OU_csv.ps1", OU], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = process.communicate()
    if err:
        print(f"Error: {err}")
    else:
        print(f"Output: {out}")

def scan_ou_users():
    # Path to the PowerShell script
    script_path = './extract_OU_users_list_sep.ps1'

    # Running the PowerShell script
    try:
        result = subprocess.run(['powershell', '-ExecutionPolicy', 'Unrestricted', '-File', script_path], capture_output=True, text=True, check=True)
        print("Script output:", result.stdout)
    except subprocess.CalledProcessError as e:
        print("Error running script:", e.stderr)

    populate_ou_csv_combobox()


# Configure our menu
ps_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="PowerShell Commands", menu=ps_menu)
py_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="Python Commands", menu=py_menu)
db_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="Database Commands", menu=db_menu)
# Drop down menu
ps_menu.add_command(label="Where - Selected User OU", command=show_user_ou)
ps_menu.add_separator()
ps_menu.add_command(label="Extract OU to CSV - DB Layout", command=get_OU_and_extract)
ps_menu.add_separator()
ps_menu.add_command(label="Copy SAM to employeeID", command=copy_sam_to_employeeid)
ps_menu.add_command(label="Erase SAM from employeeID", command=erase_employeeid_for_multiple_users)
ps_menu.add_command(label="Scan AD OU's Users", command=scan_ou_users)
ps_menu.add_separator()
ps_menu.add_command(label="Edit Sync DB Script", command=edit_sync_script)
py_menu.add_command(label="Add One Fake Record", command=add_fake_record)
py_menu.add_command(label="Add Multiple Fake Record", command=add_multiple_fake_records)
db_menu.add_command(label="Create New Database", command=create_new_database)
db_menu.add_command(label="Browse Database", command=browse_db)
db_menu.add_separator()
db_menu.add_command(label="Backup Current Database", command=backup_database)
db_menu.add_separator()
db_menu.add_command(label="Import Record From CSV", command=import_csv_data)
db_menu.add_command(label="Export Database To CSV", command=export_to_csv)
db_menu.add_separator()

# Bind the update_display_name function to the first name and last name entry fields
fn_entry.bind("<KeyRelease>", lambda event: update_display_name())
ln_entry.bind("<KeyRelease>", lambda event: update_display_name())

# Add Buttons
command_frame = LabelFrame(info_tab, text="Database Management", border=5, bg="#dcdad5")
command_frame.place(x=15, y=500)

address_frame = LabelFrame(tools_tab, text="Create Address", border=5, bg="#dcdad5")
address_frame.place(x=10, y=10)

commanddb_frame = LabelFrame(info_tab, text="Synchronization Process", border=5, bg="#dcdad5")
commanddb_frame.place(x=15, y=680)

autofill_frame = LabelFrame(info_tab, text="Auto Filling", border=5, bg="#dcdad5")
autofill_frame.place(x=15, y=595)

add_button = ttk.Button(command_frame, text="Add Record", style="Custom.TButton", command=add_record)
add_button.grid(row=0, column=0, padx=9, pady=10, sticky='w')

update_button = ttk.Button(command_frame, text="Update Record", style="Custom.TButton", command=update_record)
update_button.grid(row=0, column=1, padx=9, pady=10, sticky='w')

remove_button = ttk.Button(command_frame, text="Remove Record", style="Custom.TButton", command=remove)
remove_button.grid(row=0, column=2, padx=9, pady=10, sticky='w')

remove_all_button = ttk.Button(command_frame, text="Remove All Records", style="Custom.TButton", command=remove_all)
remove_all_button.grid(row=0, column=3, padx=9, pady=10, sticky='w')

select_record_button = ttk.Button(command_frame, text="Clear Entry Boxes", style="Custom.TButton", command=clear_entries)
select_record_button.grid(row=0, column=4, padx=9, pady=10, sticky='w')

delete_duplicates_button = ttk.Button(command_frame, text="Delete Duplicates", style="Custom.TButton", command=confirm_delete_duplicates)
delete_duplicates_button.grid(row=0, column=5, padx=9, pady=10, sticky='w')

reassign_ids_button = ttk.Button(command_frame, text="Reassign IDs", style="Custom.TButton", command=reassign_ids)
reassign_ids_button.grid(row=0, column=6, padx=9, pady=10, sticky='w')

style.configure("Custom.TButton", relief='solid')
sync_button = ttk.Button(commanddb_frame, text="Sync Database to AD", style="Custom.TButton", command=sync_records_to_ad)
sync_button.grid(row=0, column=2, padx=10, pady=10, sticky='w')

password_button = ttk.Button(commanddb_frame, text="Change User Password", style="Custom.TButton", command=enter_new_password)
password_button.grid(row=0, column=3, padx=10, pady=10, sticky='w')

disable_button = ttk.Button(commanddb_frame, text="Enable/Disable User(s)", style="Custom.TButton", command=toggle_selected_users)
disable_button.grid(row=0, column=4, padx=10, pady=10, sticky='w')

# Create a ComboBox to list available CSV files
csv_combobox = ttk.Combobox(autofill_frame, width=15)
csv_combobox.grid(row=0, column=0, padx=10, pady=10)
csv_combobox.bind("<<ComboboxSelected>>", on_csv_select)
csv_combobox.set('Address')
populate_csv_combobox()

# Create the ComboBox for selecting CSV files from the "domain" folder
domain_csv_combobox = ttk.Combobox(autofill_frame, width=15)
domain_csv_combobox.grid(row=0, column=1, padx=10, pady=10)
domain_csv_combobox.bind("<<ComboboxSelected>>", on_domain_csv_select)
domain_csv_combobox.set('sAM@domain')
populate_domain_csv_combobox()

# Create the ComboBox for selecting databases
db_combobox = ttk.Combobox(commanddb_frame, width=19)
db_combobox.grid(row=0, column=0, padx=10, pady=10, sticky='w')
db_combobox.bind("<<ComboboxSelected>>", load_selected_db)
db_combobox.set('Select DataBase')
populate_db_combobox()

# Create a ComboBox to list OU
ou_combobox = ttk.Combobox(commanddb_frame, width=26)
ou_combobox.grid(row=0, column=1, padx=9, pady=10)
ou_combobox.bind("<<ComboboxSelected>>", on_ou_csv_select)
ou_combobox.set('Select Organization Unit')
populate_ou_csv_combobox()

# Bind the treeview
my_tree.bind("<ButtonRelease-1>", select_record)

# Run to pull data from database on start
query_database()

# Call the functions to create widgets in the Tools tab
create_address_csv()
create_ou_csv()
create_sam_at_domain_csv()
create_edition_section()

root.mainloop()