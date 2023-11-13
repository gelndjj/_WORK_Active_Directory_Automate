from tkinter import *
from tkinter import ttk, messagebox, filedialog, font, simpledialog
import sqlite3, csv, os, subprocess, sys, shutil, time, glob, shlex, datetime
from faker import Faker
from tkinter.scrolledtext import ScrolledText

root = Tk()
root.title('Active Directory Automate - Users')
root.geometry("925x805")
root.resizable(False, False)
root.tk_setPalette(background="#ececec")

# Create a Style instance
style = ttk.Style()

# Set the theme to "clam"
style.theme_use("clam")

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

# Create the 'organisation unit' folder if it doesn't exist
organisation_unit_folder = os.path.join(script_dir, "organisation_unit")
os.makedirs(organisation_unit_folder, exist_ok=True)

def create_new_database():
    # Prompt the user to enter the name for the new database
    new_db_name = simpledialog.askstring("New Database", "Enter name for the new database:")

    if new_db_name:
        # Ensure the database name ends with '.db'
        if not new_db_name.endswith('.db'):
            new_db_name += '.db'

        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Create a new database file in the script directory
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

    # Get a list of database files in the script directory
    db_files = [f for f in os.listdir() if f.endswith('.db')]

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
tree_frame = ttk.Frame(root)
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
data_frame = LabelFrame(root, text="Record", border=5)
data_frame.place(x=15, y=285)

id_var = StringVar()
id_var.set("Auto-Generated")
id_entry = ttk.Entry(data_frame, textvariable=id_var, state='disabled')
id_entry.grid(row=0, column=1, padx=10, pady=5)

id_label = Label(data_frame, text="ID", font=("Academy Engraved", 10))
id_label.grid(row=0, column=0, padx=10, pady=10)

ln_label = Label(data_frame, text="Last Name", font=("Academy Engraved", 10))
ln_label.grid(row=0, column=2, padx=10, pady=10)
ln_entry = ttk.Entry(data_frame)
ln_entry.grid(row=0, column=3, padx=10, pady=10)

fn_label = Label(data_frame, text="First Name", font=("Academy Engraved", 10))
fn_label.grid(row=0, column=4, padx=10, pady=10)
fn_entry = ttk.Entry(data_frame)
fn_entry.grid(row=0, column=5, padx=10, pady=10)

dn_label = Label(data_frame, text="Display Name", font=("Academy Engraved", 10))
dn_label.grid(row=1, column=0, padx=10, pady=10)
dn_entry = ttk.Entry(data_frame, textvariable=display_name_var)
dn_entry.grid(row=1, column=1, padx=10, pady=10)

dc_label = Label(data_frame, text="Description", font=("Academy Engraved", 10))
dc_label.grid(row=1, column=2, padx=10, pady=10)
dc_entry = ttk.Entry(data_frame)
dc_entry.grid(row=1, column=3, padx=10, pady=10)

em_label = Label(data_frame, text="Email", font=("Academy Engraved", 10))
em_label.grid(row=1, column=4, padx=10, pady=10)
em_entry = ttk.Entry(data_frame)
em_entry.grid(row=1, column=5, padx=10, pady=10)

sm_label = Label(data_frame, text="SAM", font=("Academy Engraved", 10))
sm_label.grid(row=2, column=0, padx=10, pady=10)
sm_entry = ttk.Entry(data_frame)
sm_entry.grid(row=2, column=1, padx=10, pady=10)

up_label = Label(data_frame, text="UPN", font=("Academy Engraved", 10))
up_label.grid(row=2, column=2, padx=10, pady=10)
up_entry = ttk.Entry(data_frame)
up_entry.grid(row=2, column=3, padx=10, pady=10)

address_label = Label(data_frame, text="Address", font=("Academy Engraved", 10))
address_label.grid(row=2, column=4, padx=10, pady=10)
address_entry = ttk.Entry(data_frame)
address_entry.grid(row=2, column=5, padx=10, pady=10)

city_label = Label(data_frame, text="City", font=("Academy Engraved", 10))
city_label.grid(row=3, column=0, padx=10, pady=10)
city_entry = ttk.Entry(data_frame)
city_entry.grid(row=3, column=1, padx=10, pady=10)

state_label = Label(data_frame, text="State", font=("Academy Engraved", 10))
state_label.grid(row=3, column=2, padx=10, pady=10)
state_entry = ttk.Entry(data_frame)
state_entry.grid(row=3, column=3, padx=10, pady=10)

zipcode_label = Label(data_frame, text="Zipcode", font=("Academy Engraved", 10))
zipcode_label.grid(row=3, column=4, padx=10, pady=10)
zipcode_entry = ttk.Entry(data_frame)
zipcode_entry.grid(row=3, column=5, padx=10, pady=10)

cp_label = Label(data_frame, text="Company", font=("Academy Engraved", 10))
cp_label.grid(row=4, column=0, padx=10, pady=10)
cp_entry = ttk.Entry(data_frame)
cp_entry.grid(row=4, column=1, padx=10, pady=10)

jt_label = Label(data_frame, text="Job Title", font=("Academy Engraved", 10))
jt_label.grid(row=4, column=2, padx=10, pady=10)
jt_entry = ttk.Entry(data_frame)
jt_entry.grid(row=4, column=3, padx=10, pady=10)

dp_label = Label(data_frame, text="Department", font=("Academy Engraved", 10))
dp_label.grid(row=4, column=4, padx=10, pady=10)
dp_entry = ttk.Entry(data_frame)
dp_entry.grid(row=4, column=5, padx=10, pady=10)

cid_label = Label(data_frame, text="Custom Identifier", font=("Academy Engraved", 10))
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

    # outpus to entry boxes
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

            # Add New Record (including 'id' field)
            c.execute(
                "INSERT INTO customers (id, last_name, first_name, display_name, description, email, sam, upn, address, city, state, zipcode, company, job_title, department, cid) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (next_id, ln_entry.get(), fn_entry.get(), dn_entry.get(), dc_entry.get(), em_entry.get(), sm_entry.get(),
                 up_entry.get(), address_entry.get(), city_entry.get(), state_entry.get(), zipcode_entry.get(),
                 cp_entry.get(), jt_entry.get(), dp_entry.get(), cid_entry.get()))

            # Commit changes
            conn.commit()

            # Close the connection
            conn.close()

            clear_entries()

            # Clear the Treeview Table
            my_tree.delete(*my_tree.get_children())

            # Refresh the displayed records in the Treeview
            query_database()
    except Exception as e:
        messagebox.showerror("Error", f"Error adding record: {e}")


def update_display_name():
    first_name = fn_entry.get()
    last_name = ln_entry.get()

    # Get the current email value
    current_email = em_entry.get()

    # Extract the email provider (part after the '@') if it's present
    email_provider = "@"
    if "@" in current_email:
        email_parts = current_email.split("@")
        email_provider = "@" + email_parts[-1]

    # Set the display name
    display_name_var.set(f"{last_name.upper()} {first_name}")

    # Generate the email value with the same provider
    email_var = f"{first_name.lower()}.{last_name.lower()}{email_provider}" if first_name and last_name else ""

    # Update the email field
    em_entry.delete(0, END)
    em_entry.insert(0, email_var)

    # Update the UPN field
    upn_var = f"{first_name.lower()}.{last_name.lower()}" if first_name and last_name else ""
    up_entry.delete(0, END)
    up_entry.insert(0, upn_var)

    # Update the SAM field
    sam_var = f"{first_name.lower()}.{last_name.lower()}" if first_name and last_name else ""
    sm_entry.delete(0, END)
    sm_entry.insert(0, sam_var)

    # Check if the CID entry is not already set and the first name and last name are not empty
    if not cid_entry.get().strip() and first_name and last_name:
        # Fill the CID entry with the correct value
        cid_var = f"cid.{first_name.lower()}.{last_name.lower()}"
        cid_entry.delete(0, END)
        cid_entry.insert(0, cid_var)

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
                    first_name = fn_entry.get().lower()
                    last_name = ln_entry.get().lower()

                    # Construct the base part of the email, SAM, and UPN
                    base_part = f"{first_name}.{last_name}"

                    # Update Email
                    updated_email = f"{base_part}{csv_part}"
                    em_entry.delete(0, 'end')
                    em_entry.insert(0, updated_email)

                    # Update SAM (up to the '@' character in csv_part)
                    sam_part = csv_part.split('@')[0] if '@' in csv_part else csv_part
                    updated_sam = f"{base_part}{sam_part}"
                    sm_entry.delete(0, 'end')
                    sm_entry.insert(0, updated_sam)

                    # Update UPN
                    updated_upn = f"{base_part}{csv_part}"
                    up_entry.delete(0, 'end')
                    up_entry.insert(0, updated_upn)

                else:
                    messagebox.showerror("Error", "No text to append from CSV.")
    except FileNotFoundError:
        messagebox.showerror("Error", "CSV file not found.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

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
        csv_files = [f for f in os.listdir(organisation_unit_folder) if f.endswith('.csv')]
        # Set the values of the ComboBox to the list of CSV files
        ou_combobox['values'] = csv_files
    except FileNotFoundError:
        messagebox.showerror("Error", "The 'organisation unit' folder was not found.")

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
        csv_filename = os.path.join("organisation_unit", selected_csv)  # Assuming CSV files are in the "organisation_unit folder" folder

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
        if not selected_csv or selected_csv == 'Select Organisation Unit':
            print("Please select a valid Organisation Unit from the combobox.")
            return

        # Assuming the CSV has only one line which is the OU path
        csv_path = os.path.join(script_dir, "organisation_unit", selected_csv)
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
    db_file = filedialog.askopenfilename(filetypes=[("SQLite Database Files", "*.db")])

    if db_file:
        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Check if the selected database file is in the script directory
        if os.path.dirname(db_file) == script_dir:
            # Update the ComboBox with available database names
            db_combobox['values'] = tuple(os.path.basename(f) for f in glob.glob("*.db"))
            db_combobox.set(os.path.basename(db_file))  # Set the ComboBox to the selected database
        else:
            # Copy the selected database file to the script directory
            new_db_path = os.path.join(script_dir, os.path.basename(db_file))
            shutil.copy(db_file, new_db_path)

            # Update the ComboBox with available database names
            db_combobox['values'] = tuple(os.path.basename(f) for f in glob.glob("*.db"))
            db_combobox.set(os.path.basename(new_db_path))  # Set the ComboBox to the selected database

        # Load and display the selected database
        load_selected_db(None)  # Call the function with an event parameter as None

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
    # Create a TopLevel window
    address_window = Toplevel(root)
    address_window.geometry('300x320')
    address_window.title("Create Address")

    # Create a label and entry field for the CSV filename
    Label(address_window, text="CSV Filename").pack()
    filename_entry = ttk.Entry(address_window)
    filename_entry.pack()

    # Create labels and entry fields for other headers
    fields = ['Company', 'Address', 'City', 'State', 'Zipcode']
    entries = {}

    for field in fields:
        Label(address_window, text=field).pack()
        entry = ttk.Entry(address_window)
        entry.pack()
        entries[field] = entry

    def save_to_csv():
        # Get values from the entry fields
        data = [entries[field].get() for field in fields]

        # Get the CSV filename from the entry field
        csv_filename = filename_entry.get() + '.csv'

        # Specify the CSV file path (you can change the directory as needed)
        csv_directory = 'addresses'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, csv_filename)

        # Write data to the CSV file
        with open(csv_filepath, 'w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(fields)
            csv_writer.writerow(data)

        # Close the window after saving
        address_window.destroy()
        populate_csv_combobox()

    # Create a button to save the data to CSV
    save_button = ttk.Button(address_window, text="Save", command=save_to_csv)
    save_button.pack(padx=10, pady=10)

def create_sam_at_domain_csv():
    # Create a TopLevel window
    sam_at_domain_window = Toplevel(root)
    sam_at_domain_window.geometry('300x150')
    sam_at_domain_window.title("Create sAM@domain CSV")

    # Create labels and entry fields for the 'Status' header and CSV file name
    Label(sam_at_domain_window, text="CSV File Name").pack()
    csv_filename_entry = ttk.Entry(sam_at_domain_window)
    csv_filename_entry.pack()

    Label(sam_at_domain_window, text="Status").pack()
    status_entry = ttk.Entry(sam_at_domain_window)
    status_entry.pack()

    def save_to_csv():
        # Get the values from the entry fields
        status_value = status_entry.get()
        csv_filename = csv_filename_entry.get()

        # Specify the CSV file path (you can change the directory as needed)
        csv_directory = 'domain'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, f'{csv_filename}.csv')

        # Write data to the CSV file
        with open(csv_filepath, 'w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Status'])
            csv_writer.writerow([status_value])

        # Close the window after saving
        sam_at_domain_window.destroy()
        populate_domain_csv_combobox()

    # Create a button to save the data to CSV
    save_button = ttk.Button(sam_at_domain_window, text="Save", command=save_to_csv)
    save_button.pack(padx=10, pady=10)

def create_ou_csv():
    # Create a TopLevel window
    ou_window = Toplevel(root)
    ou_window.geometry('300x150')
    ou_window.title("Create Organisation Unit CSV")

    # Create labels and entry fields for the 'Status' header and CSV file name
    Label(ou_window, text="CSV File Name").pack()
    csv_filename_entry = ttk.Entry(ou_window)
    csv_filename_entry.pack()

    Label(ou_window, text="Organisation Unit").pack()
    status_entry = ttk.Entry(ou_window)
    status_entry.pack()

    def save_to_csv():
        # Get the values from the entry fields
        status_value = status_entry.get()
        csv_filename = csv_filename_entry.get()

        # Specify the CSV file path (you can change the directory as needed)
        csv_directory = 'organisation_unit'
        os.makedirs(csv_directory, exist_ok=True)
        csv_filepath = os.path.join(csv_directory, f'{csv_filename}.csv')

        # Write data directly to the file
        with open(csv_filepath, 'w') as csvfile:
            csvfile.write(status_value + '\n')

        # Close the window after saving
        ou_window.destroy()
        populate_ou_csv_combobox()

    # Create a button to save the data to CSV
    save_button = ttk.Button(ou_window, text="Save", command=save_to_csv)
    save_button.pack(padx=10, pady=10)

def enter_new_password():
    # Create a TopLevel window for entering the new password
    password_window = Toplevel(root)
    password_window.title("Enter New Password")

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
    if not selected_csv or selected_csv == 'Select Organisation Unit':
        messagebox.showwarning("Warning", "Please select a valid OU from Organisation Unit Management.")
        return

    # Construct the path to the CSV inside the 'organisation_unit' folder
    full_csv_path = os.path.join('organisation_unit', selected_csv)

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
ad_menu = Menu(my_menu, tearoff=0)
my_menu.add_cascade(label="AD Commands", menu=ad_menu)
# Drop down menu
ad_menu.add_command(label="Where - Selected User OU", command=show_user_ou)
ad_menu.add_separator()
ad_menu.add_command(label="Extract OU to CSV - DB Layout", command=get_OU_and_extract)
ad_menu.add_separator()
ad_menu.add_command(label="Copy SAM to employeeID", command=copy_sam_to_employeeid)
ad_menu.add_command(label="Erase SAM from employeeID", command=erase_employeeid_for_multiple_users)

# Bind the update_display_name function to the first name and last name entry fields
fn_entry.bind("<KeyRelease>", lambda event: update_display_name())
ln_entry.bind("<KeyRelease>", lambda event: update_display_name())

# Add Buttons
command_frame = LabelFrame(root, text="Database Management", border=5)
command_frame.place(x=760, y=285)

commanddb_frame = LabelFrame(root, text="Database Storing", border=5)
commanddb_frame.place(x=400, y=525)

mgmtou_frame = LabelFrame(root, text="Organisation Unit Management", border=5)
mgmtou_frame.place(x=400, y=675)

autofill_frame = LabelFrame(root, text="Auto Filling", border=5)
autofill_frame.place(x=622, y=675)

ps_frame = LabelFrame(root, text="PowerShell Commands", border=5)
ps_frame.place(x=15, y=525)

py_frame = LabelFrame(root, text="Python Commands", border=5)
py_frame.place(x=200, y=525)

update_button = ttk.Button(command_frame, text="Update Record", style="Custom.TButton", command=update_record)
update_button.grid(row=1, column=0, padx=10, pady=10, sticky='w')

add_button = ttk.Button(command_frame, text="Add Record", style="Custom.TButton", command=add_record)
add_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')

remove_all_button = ttk.Button(command_frame, text="Remove All Records", style="Custom.TButton", command=remove_all)
remove_all_button.grid(row=6, column=0, padx=10, pady=10, sticky='w')

remove_button = ttk.Button(command_frame, text="Remove Record", style="Custom.TButton", command=remove)
remove_button.grid(row=5, column=0, padx=10, pady=10, sticky='w')

select_record_button = ttk.Button(command_frame, text="Clear Entry Boxes", style="Custom.TButton", command=clear_entries)
select_record_button.grid(row=2, column=0, padx=10, pady=10, sticky='w')

fake_button = ttk.Button(py_frame, text="Add One Fake Record", style="Custom.TButton", command=add_fake_record)
fake_button.grid(row=0, column=0, padx=10, pady=16, sticky='w')

fake_mul_button = ttk.Button(py_frame, text="Add Multiple Fake Record", style="Custom.TButton", command=add_multiple_fake_records)
fake_mul_button.grid(row=1, column=0, padx=10, pady=16, sticky='w')

export_button = ttk.Button(py_frame, text="Export Database to CSV", style="Custom.TButton", command=export_to_csv)
export_button.grid(row=2, column=0, padx=10, pady=16, sticky='w')

import_csv_button = ttk.Button(py_frame, text="Import Record from CSV", style="Custom.TButton", command=import_csv_data)
import_csv_button.grid(row=3, column=0, padx=10, pady=18, sticky='w')

delete_duplicates_button = ttk.Button(command_frame, text="Delete Duplicates", style="Custom.TButton", command=confirm_delete_duplicates)
delete_duplicates_button.grid(row=4, column=0, padx=10, pady=10, sticky='w')

reassign_ids_button = ttk.Button(command_frame, text="Reassign IDs", style="Custom.TButton", command=reassign_ids)
reassign_ids_button.grid(row=3, column=0, padx=10, pady=10, sticky='w')

backup_button = ttk.Button(commanddb_frame, text="Backup Current Database", style="Custom.TButton", command=backup_database)
backup_button.grid(row=1, column=1, padx=35, pady=17, sticky='w')

browse_button = ttk.Button(commanddb_frame, text="Browse Database", style="Custom.TButton", command=browse_db)
browse_button.grid(row=1, column=0, padx=10, pady=17, sticky='w')

create_new_db_button = ttk.Button(commanddb_frame, text="Create New Database", style="Custom.TButton", command=create_new_database)
create_new_db_button.grid(row=0, column=1, padx=35, pady=10, sticky='w')

style.configure("Custom.TButton", relief='solid')
sync_button = ttk.Button(ps_frame, text="Sync Database to AD", style="Custom.TButton", command=sync_records_to_ad)
sync_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')

edit_script_button = ttk.Button(ps_frame, text="Edit Sync DB Script", style="Custom.TButton", command=edit_sync_script)
edit_script_button.grid(row=1, column=0, padx=10, pady=10, sticky='w')

password_button = ttk.Button(ps_frame, text="Change User Password", style="Custom.TButton", command=enter_new_password)
password_button.grid(row=2, column=0, padx=10, pady=10, sticky='w')

disable_button = ttk.Button(ps_frame, text="Enable/Disable User(s)", style="Custom.TButton", command=toggle_selected_users)
disable_button.grid(row=3, column=0, padx=10, pady=10, sticky='w')

create_address_button = ttk.Button(autofill_frame, text="Create Address", style="Custom.TButton", command=create_address_csv)
create_address_button.grid(row=0, column=1, padx=10, pady=10, sticky='w')

create_sam_at_domain_button = ttk.Button(autofill_frame, text="Create sAM@domain", style="Custom.TButton", command=create_sam_at_domain_csv)
create_sam_at_domain_button.grid(row=1, column=1, padx=10, pady=12, sticky='w')

create_ou_button = ttk.Button(mgmtou_frame, text="Create Organization Unit Path", style="Custom.TButton", command=create_ou_csv)
create_ou_button.grid(row=1, column=0, padx=10, pady=1, sticky='w')

scan_ou_button = ttk.Button(mgmtou_frame, text="Scan AD OU's Users", style="Custom.TButton", command=scan_ou_users)
scan_ou_button.grid(row=2, column=0, padx=10, pady=1, sticky='w')

# Create a ComboBox to list available CSV files
csv_combobox = ttk.Combobox(autofill_frame, width=15)
csv_combobox.grid(row=0, column=0, padx=10, pady=10)
csv_combobox.bind("<<ComboboxSelected>>", on_csv_select)
csv_combobox.set('Address')
populate_csv_combobox()

# Create the ComboBox for selecting CSV files from the "domain" folder
domain_csv_combobox = ttk.Combobox(autofill_frame, width=15)
domain_csv_combobox.grid(row=1, column=0, padx=10, pady=10)
domain_csv_combobox.bind("<<ComboboxSelected>>", on_domain_csv_select)
domain_csv_combobox.set('sAM@domain')
populate_domain_csv_combobox()

# Create the ComboBox for selecting databases
db_combobox = ttk.Combobox(commanddb_frame, width=14)
db_combobox.grid(row=0, column=0, padx=10, pady=10, sticky='w')
db_combobox.bind("<<ComboboxSelected>>", load_selected_db)
db_combobox.set('Select DataBase')
populate_db_combobox()

# Create a ComboBox to list OU
ou_combobox = ttk.Combobox(mgmtou_frame, width=23)
ou_combobox.grid(row=0, column=0, padx=10, pady=10)
ou_combobox.bind("<<ComboboxSelected>>", on_ou_csv_select)
ou_combobox.set('Select Organisation Unit')
populate_ou_csv_combobox()

# Bind the treeview
my_tree.bind("<ButtonRelease-1>", select_record)

# Run to pull data from database on start
query_database()

root.mainloop()