from tkinter import *
from tkinter.messagebox import showinfo, askquestion
from tkinter.simpledialog import askstring
from tkinter.scrolledtext import ScrolledText
from tkinter import font
from tkinter import ttk
import sqlite3
import os
import platform
from datetime import datetime, date, timedelta

icon = os.path.abspath('./fms.ico')


class Connection:
    def __init__(self):
        self.conn = sqlite3.connect("fms.db")

    def __enter__(self):
        return self.conn

    def __exit__(self, exec_type, exec_val, tb):
        if exec_type:
            self.conn.rollback()
        else:
            self.conn.commit()

        self.conn.close()


class Treeview(ttk.Treeview):
    def __init__(self, parent, headers, parent_self, *args, **kwargs):
        ttk.Treeview.__init__(self, parent,  columns = headers, show="headings", height=2, *args, **kwargs)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self.xview)
        self.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill="y", anchor='w')
        hsb.pack(side='bottom', fill="x")

        self.parent = parent
        self.headers = headers
        self.current_selection = None
        self.parent_self= parent_self
        self.bind("<<TreeviewSelect>>", self.get_selection)
        self._build_tree()

    def set_headers(self, headers):
        self.headers = headers

    def _build_tree(self):
        for col in self.headers:
            self.heading(col, text=col, anchor='w',command=lambda c=col: self.sortby(self, c, 0))
            # adjust the column's width to the header string
            self.column(col, anchor='nw', width=100)

    def sortby(self, tree, col, descending):
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        data.sort(reverse=descending)
        for ix, item in enumerate(data):
            tree.move(item[1], '', ix)
        tree.heading(col, command=lambda col=col: self.sortby(tree, col, int(not descending)))

    def fill_tree(self):
        # Delete children and repopulate
        self.delete(*self.get_children())

        if self.register is not None:
            for item in self.register:
                self.insert('', 'end', values=item)  # Returns row-id
                # adjust column's width if necessary to fit each value
                try:
                    for ix, val in enumerate(item):
                        col_w = font.Font().measure(val)
                        if self.column(self.headers[ix], width=None) < col_w:
                            self.column(self.headers[ix], width=col_w)
                except TypeError:
                    raise TypeError("Tree_list must be a list of tuples")

    def get_selection(self, event=None):
        if event:
            current_selection = event.widget.focus()
            rows = self.item(current_selection)['values']
            self.current_selection = rows
            if hasattr(self.parent_self, 'Find'):
                self.parent_self.Find(REF = rows[0])

    def set_register(self, register):
        self.register = register
        self.update_tree()

    def update_tree(self):
        self.clear()
        self.fill_tree()

    def clear(self):
        self.delete(*self.get_children())

    def get_all(self):
        return [item for item in self.get_children()]


class Dialog(Toplevel):
    def __init__(self, title):
        super().__init__()

        self.title(title)
        if platform.system() =="Windows":
            self.iconbitmap(icon)
        self.resizable(0, 0)
        self.focus()
        self.grab_set()
        self.configure(padx=4)
        self.geometry("+300+20")


class AskString(Toplevel):
    def __init__(self, parent, master, title, prompt):
        super().__init__()

        self.parent = parent
        self.title(title)
        if platform.system() =="Windows":
            self.iconbitmap(icon)
            
        self.resizable(0, 0)
        self.focus()
        self.transient(master)
        self.grab_set()
        self.configure(padx=4)
        self.geometry("+500+20")

        Label(self, text=prompt, font='calibri 12').grid(row=0, columnspan=2)

        self.string = ttk.Entry(self, width=25)
        self.string.grid(row=1, columnspan=2)
        self.string.configure(font='calibri 12')
        self.string.bind("<Return>", self.submit)

        self.okay = ttk.Button(self, text='OK', command=self.submit)
        self.cancel = ttk.Button(self, text='Cancel', command=self.destroy)

        self.okay.grid(row=2, column=0, pady=4)
        self.cancel.grid(row=2, column=1, pady=4)


    def submit(self, event=None):
        val = self.string.get()
        self.destroy()
        self.parent.FindComplainant(val)


class MyScrolledText(ScrolledText):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def insert(self, index=None, val=None):
        super().insert("1.0", val)

    def get(self, index=None, end=None):
        return super().get('1.0', 'end')

    def delete(self, index, end):
        super().delete("1.0", 'end')



class Base:
    fields = []

    def build_toolbar(self, parent):
        self.toolbar = Frame(parent)
        self.toolbar.pack(padx=80,fill=X, pady=10, side=BOTTOM)

    def add_tool_buttons(self):
        self.buttons = {}

        btns = ["Save", "Clear", "Update", "Find", "DEL"]
        commands = [self.Save, self.Clear,self.Update, self.Find, self.Delete]

        for btn, cmd in zip(btns, commands):
            b= ttk.Button(self.toolbar, text=btn, command=cmd, cursor='hand2')
            b.pack(side=LEFT, padx=4)
            self.buttons[btn] = b


    def build_interface(self, title, parent):
        frame = LabelFrame(parent, 
            text=title, 
            font='Arial 14', 
            fg='blue', padx=5, pady=5)
        frame.pack()

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", background='powderblue')

        for i, field in enumerate(self.fields):
            l=Label(frame, text=field.upper(), font='Calibri 12')
            l.grid(row=i, column=0, sticky='e')

            if field not in ["File Sent To", "Location of File"]:
                if field == "Remarks":
                    entry = MyScrolledText(frame, width=32,
                        wrap=WORD, padx=4,pady=4, borderwidth=2,
                        background='honeydew', height=4)
                else:
                    entry = ttk.Entry(frame, width=34)

            else:
                if field in ["File Sent To"]:
                    entry = ttk.Combobox(frame, width=32,
                        values=["DPP", "RSA"])
                elif field in ["Location of File"]:
                    entry = ttk.Combobox(frame, width=32,
                        values=["LPPU", "CID HEADQTRS"])

            entry.grid(row=i, column=1, pady=4)
            entry.configure(font='Consolas 14')

            key = field.upper().replace(" ", "_")
            self.entries[key] = entry

    def Save(self):
        values = []
        for ent in self.entries.values():
            try:
                values.append(ent.get().strip())
            except TclError:
                pass

        values = tuple(values)
        keys = ", ".join([f.upper().replace(" ", "_") for f in self.fields])
        
        SQL = """INSERT INTO %s (%s) VALUES(""" % (self.table, keys)
        SQL += """ "%s",""" * len(values) % (values)
        SQL = SQL[:-1] + ")"

        with Connection() as conn:
            cur = conn.cursor()
            try:
                cur.execute(SQL)
            except sqlite3.IntegrityError:
                showinfo("SaveError", "This record already exists")
            except Exception as e:
                showinfo("SaveError", str(e))
            else:
                showinfo("Done", "Saved record with file number: %s"%values[0])


    def exists_in_records(self):
        try:
            REF = self.entries["ORIGINAL_REF_NO"].get()
            sql = "SELECT * FROM %s WHERE ORIGINAL_REF_NO='%s'"%(self.table, REF)

        except KeyError:
            REF = self.entries["CURRENT_REF_NO"].get()
            sql = "SELECT * FROM %s WHERE CURRENT_REF_NO='%s'"%(self.table, REF)
        
        with Connection() as conn:
            cur = conn.cursor()
            cur.execute(sql)
            results = cur.fetchone()

            if results:
                return True

            else:
                return False


    def Update(self):
        if not self.exists_in_records():
            showinfo("Aborted", "This Ref Number is not in records")
            return False

        values = []
        for ent in self.entries.values():
            try:
                values.append(ent.get())
            except TclError:
                pass

        values = tuple(values)
        keys = ", ".join([f.upper().replace(" ", "_") for f in self.fields])

        SQL = """UPDATE {} SET """.format(self.table)

        for key, value in zip(keys.split(", "), values):
            SQL += """ %s = "%s" ,""" % (key, value)

        if "ORIGINAL_REF_NO" in self.entries:
            SQL = SQL[:-1] + """ WHERE ORIGINAL_REF_NO = "%s" """%values[0]
        else:
            SQL = SQL[:-1] + """ WHERE CURRENT_REF_NO = "%s" """%values[0]

        with Connection() as conn:
            cur = conn.cursor()
            try:
                cur.execute(SQL)
            except Exception as e:
                showinfo("Update Error", str(e))
            else:
                showinfo("Success","Updated the File record: %s"%values[0])


    def Find(self, event=None, REF=None):
        if "ORIGINAL_REF_NO" in self.entries:
            if REF:
                SQL = "SELECT * FROM %s WHERE ORIGINAL_REF_NO='%s'"%(
                    self.table, REF)
            else:
                SQL = "SELECT * FROM %s WHERE ORIGINAL_REF_NO='%s'"%(
                    self.table, self.entries['ORIGINAL_REF_NO'].get())
        else:
            if REF:
                SQL = "SELECT * FROM %s WHERE CURRENT_REF_NO='%s'"%(
                    self.table, REF)
            else:
                SQL = "SELECT * FROM %s WHERE CURRENT_REF_NO='%s'"%(
                    self.table, self.entries['CURRENT_REF_NO'].get())

        with Connection() as conn:
            cur = conn.cursor()
            try:
                cur.execute(SQL)
            except Exception as e:
                showinfo("Lookup Error", str(e))
            else:
                results = cur.fetchall()
                colnames = [d[0] for d in cur.description]
                results_dict = [dict(zip(colnames, r)) for r in results]

                if results_dict:
                    for key, val in results_dict[0].items():
                        self.entries[key].delete(0, END)
                        self.entries[key].insert(0, val)


    def Delete(self):
        ref = self.entries[self.fields[0].upper().replace(" ","_")].get()
        if not ref:
            return False

        ans = askquestion("Delete", "Are you sure you want to delete"
            " this record. \n REF NO: %s"%ref)
        if ans !='yes':
            return False

        if not self.exists_in_records():
            showinfo("Delete", 'This record does not exist in database')
            return False

        if "ORIGINAL_REF_NO" in self.entries:
            SQL = "DELETE FROM %s WHERE ORIGINAL_REF_NO='%s'"%(self.table, 
                self.entries['ORIGINAL_REF_NO'].get())
        else:
            SQL = "DELETE FROM %s WHERE CURRENT_REF_NO='%s'"%(self.table, 
                self.entries['CURRENT_REF_NO'].get())

        with Connection() as conn:
            try:
                cur = conn.cursor()
                cur.execute(SQL)
            except Exception as e:
                showinfo("Info", str(e))
            else:
                showinfo("Success", "Delete record successfully")


    def Clear(self):
        for e in self.entries.values():
            try:
                e.delete(0, END)
            except TclError:
                pass

        try:
            self.entries['ORIGINAL_REF_NO'].focus()
        except KeyError:
            self.entries['CURRENT_REF_NO'].focus()


    def FindComplainant(self, complainant=None):
        if complainant:
            SQL = "SELECT * FROM {} WHERE COMPLAINANT LIKE '%{}%'".format(
                self.table, complainant) 
        else:
            SQL = "SELECT * FROM {}".format(self.table)

        with Connection() as conn:
            cur = conn.cursor()
            try:
                cur.execute(SQL)
            except Exception as e:
                showinfo("Lookup Error", str(e))
            else:
                results = cur.fetchall()
                colnames = [d[0].replace("_", " ") for d in cur.description]
                if results:
                    self.show_tree(colnames, results)


    def show_tree(self, headers, data):
        top = Toplevel()
        top.title("FMS: Advanced Search")
        if platform.system() =="Windows":
            top.iconbitmap(icon)
        top.geometry("1000x600+50+10")

        label = Label(top, text=self.__class__.__name__ +\
            " (Click to select a record and fill it in the form)",
            font='Arial 18 bold', fg='blue')
        label.pack(pady=10)

        self.tree = Treeview(top, headers, self)
        self.tree.set_register(data)
        self.tree.pack(expand=1, fill=BOTH)


    def FindAll(self):
        self.FindComplainant(complainant=None)
  

    def fill_form(self, results_dict):
        for key, val in results_dict[0].items():
            self.entries[key].delete(0, END)
            self.entries[key].insert(0, val)


class FilesSentToDPP(Base):
    table = 'files_sent_to_dpp'
    
    def __init__(self, frame):
        self.entries = {}
        self.fields = [
        "Original REF NO", "Current REF NO",
        "Complainant", "Suspect", "Offence",
        "Investigating Officer",
        "Date Sent", 
        "Date Returned",
        "File Sent To", 
        "Remarks"]
        self.frame = frame

        self.create_table()

    def create_table(self):
        conn = sqlite3.connect("fms.db")
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS files_sent_to_dpp
            (ORIGINAL_REF_NO VARCHAR(15) PRIMARY KEY, 
            CURRENT_REF_NO VARCHAR(15) UNIQUE,
            COMPLAINANT VARCHAR(30), 
            SUSPECT VARCHAR(30), 
            OFFENCE VARCHAR(50),
            INVESTIGATING_OFFICER VARCHAR(30), 
            DATE_SENT DATE, 
            DATE_RETURNED DATE(15),
            FILE_SENT_TO VARCHAR(20),
            REMARKS VARCHAR(200))
            ''')
    

    def build(self):
        super().build_interface("Files Sent to DPP/RSA", self.frame)


class CourtGoingFiles(Base):
    table = "court_going"

    def __init__(self, frame):
        self.entries = {}
        self.fields = [
                    "Current REF NO",
                    "Complainant", "Suspect", "Offence",
                    "Investigating Officer",
                    "Date Sent to Court", 
                    "Date Next in Court", 
                    "Status Of Case"
                    ]

        self.frame = frame
        self.create_table()

    def create_table(self):
        conn = sqlite3.connect("fms.db")
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS court_going
            ( 
            CURRENT_REF_NO VARCHAR(15) PRIMARY KEY,
            COMPLAINANT VARCHAR(30), 
            SUSPECT VARCHAR(30), 
            OFFENCE VARCHAR(50),
            INVESTIGATING_OFFICER VARCHAR(30), 
            DATE_SENT_TO_COURT DATE, 
            DATE_NEXT_IN_COURT DATE, 
            STATUS_OF_CASE VARCHAR(30))
            ''')

    def build(self):
        super().build_interface("Court Going Files", self.frame)
    

class PutAwayFiles(Base):
    table = "putaway"

    def __init__(self, frame):
        self.entries = {}
        self.fields = [
        "Original REF NO", "Current REF NO",
        "Complainant", "Suspect", "Offence",
        "Location of File", "Status", "Date Sent"]

        self.frame = frame
        self.create_table()


    def create_table(self):
        conn = sqlite3.connect("fms.db")
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS putaway
            (ORIGINAL_REF_NO VARCHAR(15) PRIMARY KEY, 
            CURRENT_REF_NO VARCHAR(15) UNIQUE,
            COMPLAINANT VARCHAR(30), 
            SUSPECT VARCHAR(30), 
            OFFENCE VARCHAR(50),
            LOCATION_OF_FILE VARCHAR(30), 
            STATUS VARCHAR(30), 
            DATE_SENT DATE)
            ''')
    

    def build(self):
        super().build_interface("Put Away Files", self.frame)


class AllocationToInvestigators(Base):
    table = "allocation"
    
    def __init__(self, frame):
        self.entries = {}
        self.fields = [
        "Original REF NO", "Current REF NO",
        "Complainant", "Suspect", "Offence",
        "Investigating Officer", "Date of Allocation"]

        self.frame = frame
        self.create_table()

    def create_table(self):
        conn = sqlite3.connect("fms.db")
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS allocation
            (ORIGINAL_REF_NO VARCHAR(15) PRIMARY KEY, 
            CURRENT_REF_NO VARCHAR(15) UNIQUE,
            COMPLAINANT VARCHAR(30), 
            SUSPECT VARCHAR(30), 
            OFFENCE VARCHAR(50),
            INVESTIGATING_OFFICER VARCHAR(30), 
            DATE_OF_ALLOCATION DATE)
            ''')
        
    def build(self):
        super().build_interface("Allocation To Investigators", self.frame)


# Main gui framework

class Main:
    def __init__(self, parent):
        self.parent = parent
        self.parent.geometry("600x450+300+50")
        self.parent.protocol("WM_DELETE_WINDOW", self.Close)
        if platform.system() =="Windows":
            self.parent.iconbitmap(icon)
        self.parent.resizable(0,0)

        static = Label(self.parent, 
            text='FILE MANAGEMENT SYSTEM (CID)',
            font='Arial 18 bold roman', fg='green')
        static.pack(fill=X, pady=10)

        self.container = Frame(parent)
        self.container.pack(side=RIGHT, expand=1)

        self.menubar()
        self.switch_frame(PutAwayFiles)
    

    def menubar(self):
        menubar = Menu(self.parent)
        self.parent.configure(menu=menubar)

        filemenu = Menu(menubar)
        filemenu.add_command(label="Files Sent To DPP OR RSA",
            command = self.ShowDPPFiles)
        filemenu.add_command(label="File Allocation to Investigators",
            command = self.ShowAllocationFiles)

        filemenu.add_command(label="Court Going files",
            command = self.ShowCourtFiles)
        filemenu.add_command(label="Put Away files",
            command = self.ShowPutAwayFiles)
        filemenu.add_command(label="Exit",
            command = self.Close)

        searchmenu = Menu(menubar, tearoff=0)
        searchmenu.add_command(label="Search by Complainant",
            command=self.AdvancedSearch)
        searchmenu.add_command(label="All Department Files",
            command=self.AllDepartmentFiles)

        analysisMenu = Menu(menubar, tearoff=0)
        analysisMenu.add_command(
            label="Analyse by Date/Month/Year",
            command=self.AnalysisWindow)

        aboutmenu = Menu(menubar, tearoff=0)
        aboutmenu.add_command(label='Help', 
            command=self.HelpDialog)
        aboutmenu.add_command(label='Developer', 
            command=self.DeveloperDialog)

        menubar.add_cascade(label="Switch Files", menu=filemenu)
        menubar.add_cascade(label="Advanced Search", menu=searchmenu)
        menubar.add_cascade(label="Analysis", menu=analysisMenu)
        menubar.add_cascade(label="About", menu=aboutmenu)


    def AnalysisWindow(self):
        win = SQLWindow("File Analysis")
        win.resizable(1,1)
        win.geometry('1100x600')


    def HelpDialog(self):
        dialog = Dialog("FMS Help")
        dialog.geometry("600x650")

        Label(dialog, 
            text="File Management System (CID)",
            font='Consolas 16 bold', 
            fg='red').pack()

        helptext="""
The File Management System was developed as a critical need
for an organised database management system to help a ardent
records personel at CID headquaters Kampala, in an attempt 
to streamline file lookup for clients.

The system is divided into four Parts:
1) Put Away Files
2) Files Sent to DPP or Resident state attorney(RSA)
3) Files Allocated to Investigating Officers
4) Court going files

SAVING RECORDS
1) From the `Switch Files` menu, switch to the appropriate
department.
2) Enter the File Details, including the Original File Reference
number and Current Reference number. For files where the the referece
number has not been changed, e.g GEF 1002/2016, the Original REF NO
and Current REF NO are the same.
3) After filling the required information, click `SAVE' button.
You should see a message that the record has been saved 
successfully, otherwise the record won't be saved.

UPDATING RECORDS
If the files acquires a new REF NO, put the new REF NO
in the field of Current REF NO and click `UPDATE` button.
To update other fields, first find the file record you want
by the Original Reference Number, enter the new information
and update. Be sure to see confirmation the the record has been
updated successfully.

FINDING RECORDS
Enter the Original Ref No and hit `FIND` button to fetch
records for that unique number.
To view all file details for a selected department,
click 'AdvancedSearch' menu item and choose `AllDepartmentFiles`.
Click on an entry to view details in the form.
To Search by Complainant, choose 'Find By Complainant' from the
the same menu item. Enter the name or part of the name of the complainant
and hit enter/OK.
        """

        text = Label(dialog, text=helptext, justify='left',
            font='calibri 10')
        text.pack(anchor='nw', expand=1, fill=BOTH)

    def DeveloperDialog(self):
        dialog = Dialog("About FMS Developer")
        
        text = """
        Dev Name: Dr Abiira Nathan Kyarugahi
        Contact: 07854581/0700198736
        Email: nabiira2by2@gmail.com
        website: abiiranathan.pythonanywhere.com
        Twitter: @abiiranathan
        """

        dev = Label(dialog, text=text, font='Calibri 12', justify='left')
        dev.pack(anchor='w')


    def Close(self, event=None):
        # answer = askquestion("Quit", 
        #     "Are you sure you want to close the program?")
        # if answer=='yes':
        self.parent.destroy()

    def ShowDPPFiles(self):
        self.parent.geometry("620x580")
        self.switch_frame(FilesSentToDPP)

    def ShowAllocationFiles(self):
        self.parent.geometry("600x450")
        self.switch_frame(AllocationToInvestigators)

    def ShowCourtFiles(self):
        self.parent.geometry("600x450")
        self.switch_frame(CourtGoingFiles)

    def ShowPutAwayFiles(self):
        self.parent.geometry("600x450")
        self.switch_frame(PutAwayFiles)
        
    def switch_frame(self, cls):
        self.refresh_container()
        self.window = cls(self.container)
        self.window.build()
        self.window.build_toolbar(self.container)
        self.window.add_tool_buttons()
        
        try:
            self.window.entries["ORIGINAL_REF_NO"].bind("<Return>", 
                self.window.Find)
            self.window.entries["ORIGINAL_REF_NO"].focus()

        except:
            self.window.entries["CURRENT_REF_NO"].focus()
            self.window.entries["CURRENT_REF_NO"].bind("<Return>", 
                self.window.Find)

    def refresh_container(self):
        self.container.destroy()
        self.container = Frame(self.parent)
        self.container.pack(expand=1, fill=BOTH)
        self.parent.update() 


    def AdvancedSearch(self):
        AskString(self.window, self.parent,"Name", "Enter name of complainant")
        
    def AllDepartmentFiles(self):
        self.window.FindAll()


class SQLWindow(Dialog):
    def __init__(self, title):
        super().__init__(title)

        self.QUERY_RESULTS = []
        self.FILES = ["FILES SENT TO DPP", 
                    "FILES SENT TO RSA",
                    "PUT AWAY FILES", 
                    "FILES ALLOCATED TO INVESTIGATORS",
                    "COURT GOING FILES"]

        self.configure(padx=4, pady=4)
        self.option_add('*Label*font', 'Calibri 12 bold')

        toolbar = LabelFrame(self, text='  Specify A range of dates OR a month and year or a year  ', padx=5, pady=5)
        toolbar.pack(expand=0, fill=X)

        toolbar0 = Frame(self)
        toolbar0.pack(side=TOP, anchor='w', fill=X)

        toolbar1 = Frame(toolbar, relief='raised', bd=2, padx=4)
        toolbar1.pack(side=LEFT, fill=BOTH, padx=4, anchor='s')

        toolbar2 = Frame(toolbar, relief='raised', bd=2)
        toolbar2.pack(side=LEFT, fill=BOTH, padx=4, anchor='s')

        toolbar3 = Frame(toolbar,  relief='raised', bd=2)
        toolbar3.pack(side=LEFT, fill=BOTH, padx=4, anchor='s')

        self.main = Frame(self)
        self.main.pack(expand=1, fill=BOTH)


        choice_label = Label(toolbar0, text="WHICH FILES?", font='Calibri 14 bold')
        choice_label.grid(row=1, column=0, sticky='w')

        self.choice_entry = ttk.Combobox(toolbar0, width=40,
                            values=[
                            "FILES SENT TO DPP", 
                            "FILES SENT TO RSA",
                            "PUT AWAY FILES", 
                            "FILES ALLOCATED TO INVESTIGATORS"])
        self.choice_entry.configure(foreground='navyblue')
        self.choice_entry.grid(row=1, column=1, sticky='w')
        self.choice_entry.configure(font='Calibri 12 bold')
        
        # Left side
        fromlabel = Label(toolbar1, text="FROM DATE")
        fromlabel.grid(row=2, column=0, sticky='w')

        tolabel   = Label(toolbar1, text="TO DATE")
        tolabel.grid(row=3, column=0, sticky='w')

        self.from_entry = ttk.Entry(toolbar1, width=30)
        self.from_entry.configure(foreground='navyblue')
        self.from_entry.configure(font='Calibri 12 bold')
        self.from_entry.grid(row=2, column=1, sticky='w', pady=5)
        
        self.to_entry   = ttk.Entry(toolbar1, width=30)
        self.to_entry.configure(foreground='navyblue')
        self.to_entry.configure(font='Calibri 12 bold')
        self.to_entry.grid(row=3, column=1, sticky='w', pady=5)
        
        submit1 = ttk.Button(toolbar1, text='SUBMIT', command=self.ReQueryRange)
        submit1.grid(row=4, column=0, columnspan=2, pady=2)

        # Total
        self.totalFile = Label(self, text="", fg='blue', font='Arial 12 bold')
        self.totalFile.pack(anchor=W)


        # Middle side
        Label(toolbar2, text="MONTH (e.g 01)", 
                            font='Calibri 12 bold'
                            ).grid(row=0, column=0)
        Label(toolbar2, text="YEAR (e.g 2017)", 
                            font='Calibri 12 bold'
                            ).grid(row=1, column=0)

        self.month = ttk.Entry(toolbar2, width=10)
        self.month.configure(font='Calibri 14')
        self.month.grid(row=0, column=1, pady=4, padx=4, sticky='w')

        self.year   = ttk.Entry(toolbar2, width=10)
        self.year.configure(font='Calibri 14')
        self.year.grid(row=1, column=1, pady=4, padx=4, sticky='w')

        submit2 = ttk.Button(toolbar2, text='SUBMIT', command=self.ReQueryMonth)
        submit2.grid(row=4, column=0, columnspan=2, pady=2)

        # By year
        yr = Label(toolbar3, text="YEAR (e.g 2017)", font='Calibri 12 bold')
        yr.grid(row=1, column=0)

        self.fullyear   = ttk.Entry(toolbar3, width=10)
        self.fullyear.configure(font='Calibri 14')
        self.fullyear.grid(row=1, column=1, pady=4, padx=4, sticky='w')
        submit3 = ttk.Button(toolbar3, text='SUBMIT', command=self.ReQueryYear)
        submit3.grid(row=4, column=0, columnspan=2, pady=2)

        # Bottom Tree
        self.tree = Treeview(self.main, [], self)
        self.tree.pack(expand=1, fill=BOTH, pady=10)

        # Set Defaults
        now = datetime.now().date()
        diff = timedelta(days=30)
        first_date = now - diff

        _from = first_date.strftime('%d-%m-%Y')
        _to = now.strftime('%d-%m-%Y')

        self.choice_entry.insert(0, "FILES SENT TO DPP")
        self.from_entry.insert(0, _from)
        self.to_entry.insert(0, _to)

    def getQueryByDate(self):
        dept  = self.choice_entry.get()
        _from = self.from_entry.get().replace("/", '-')
        _to = self.to_entry.get().replace("/", '-')

        if dept == 'FILES SENT TO DPP':
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%d-%m-%Y', DATE_SENT) >= '{}'
            AND strftime('%d-%m-%Y', DATE_SENT) <= '{}'
            AND FILE_SENT_TO='DPP'
            """.format(_from, _to)

        elif dept == "FILES SENT TO RSA":
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%d-%m-%Y', DATE_SENT) >= '{}'
            AND strftime('%d-%m-%Y', DATE_SENT) <= '{}'
            AND FILE_SENT_TO='RSA'
            """.format(_from, _to)
            
        elif dept == 'PUT AWAY FILES':
            SQL = """
            SELECT * FROM putaway 
            WHERE strftime('%d-%m-%Y', DATE_SENT) >= '{}'
            AND strftime('%d-%m-%Y', DATE_SENT) <= '{}' 
            """
            SQL = SQL.format(_from, _to)

        elif dept == 'COURT GOING FILES':
            SQL = """
            SELECT * FROM court_going 
            WHERE strftime('%d-%m-%Y', DATE_SENT_TO_COURT) >= '{}'
            AND strftime('%d-%m-%Y', DATE_SENT_TO_COURT) <= '{}' """
            SQL = SQL.format(_from, _to)

        elif dept == 'FILES ALLOCATED TO INVESTIGATORS':
            SQL = """
            SELECT * FROM allocation 
            WHERE strftime('%d-%m-%Y', DATE_OF_ALLOCATION) >= '{}'
            AND strftime('%d-%m-%Y', DATE_OF_ALLOCATION) <= '{}' """
            SQL = SQL.format(_from, _to)

        return SQL

    def getQueryByMonth(self):
        dept  = self.choice_entry.get()
        month = self.month.get()
        year = self.year.get()

        if dept == 'FILES SENT TO DPP':
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%m-%Y', DATE_SENT) == '{}-{}'
            AND FILE_SENT_TO='DPP'
            """.format(month, year)

        elif dept == "FILES SENT TO RSA":
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%m-%Y', DATE_SENT) == '{}-{}'
            AND FILE_SENT_TO='RSA'
            """.format(month, year)
            
        elif dept == 'PUT AWAY FILES':
            SQL = """
            SELECT * FROM putaway 
            WHERE strftime('%m-%Y', DATE_SENT) == '{}-{}' 
            """
            SQL = SQL.format(month, year)

        elif dept == 'COURT GOING FILES':
            SQL = """
            SELECT * FROM court_going 
            WHERE strftime('%m-%Y', DATE_SENT_TO_COURT) == '{}-{}' """
            SQL = SQL.format(month, year)

        elif dept == 'FILES ALLOCATED TO INVESTIGATORS':
            SQL = """
            SELECT * FROM allocation 
            WHERE strftime('%m-%Y', DATE_OF_ALLOCATION) == '{}-{}' """
            SQL = SQL.format(month, year)

        return SQL

    def getQueryByYear(self):
        dept  = self.choice_entry.get()
        year = self.fullyear.get()

        if dept == 'FILES SENT TO DPP':
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%Y', DATE_SENT) == '{}'
            AND FILE_SENT_TO='DPP'
            """.format(year)

        elif dept == "FILES SENT TO RSA":
            SQL = """
            SELECT * FROM files_sent_to_dpp 
            WHERE strftime('%Y', DATE_SENT) == '{}'
            AND FILE_SENT_TO='RSA'
            """.format(year)
            
        elif dept == 'PUT AWAY FILES':
            SQL = """
            SELECT * FROM putaway 
            WHERE strftime('%Y', DATE_SENT) == '{}' 
            """
            SQL = SQL.format(year)

        elif dept == 'COURT GOING FILES':
            SQL = """
            SELECT * FROM court_going 
            WHERE strftime('%Y', DATE_SENT_TO_COURT) == '{}' """
            SQL = SQL.format(year)

        elif dept == 'FILES ALLOCATED TO INVESTIGATORS':
            SQL = """
            SELECT * FROM allocation 
            WHERE strftime('%Y', DATE_OF_ALLOCATION) == '{}' """
            SQL = SQL.format(year)

        return SQL

    def ReQueryMonth(self):
        sql = self.getQueryByMonth()
        with Connection() as conn:
            cur= conn.cursor()
            cur.execute(sql)
            results = cur.fetchall()
            colnames = [d[0].replace("_", " ") for d in cur.description]
            self.handleResult(results, colnames)

            if results:
                self.totalFile['text'] = "TOTAL: %s %s IN THE MONTH-YEAR(%s-%s)"%(len(results), 
                    self.choice_entry.get(), self.month.get(), self.year.get())
            else: 
                self.totalFile['text'] = "TOTAL: 0 %s IN THE MONTH-YEAR(%s-%s)"%( 
                    self.choice_entry.get(), self.month.get(), self.year.get())


    def ReQueryYear(self):
        sql = self.getQueryByYear()
        with Connection() as conn:
            cur= conn.cursor()
            cur.execute(sql)
            results = cur.fetchall()
            
            if results:
                self.totalFile['text'] = "TOTAL: %s %s"%(
                    len(results), 
                    self.choice_entry.get() + " IN THE YEAR %s"%self.fullyear.get()) 
            else: 
                self.totalFile['text'] = "TOTAL: 0 %s"%( 
                self.choice_entry.get() + " IN THE YEAR %s"%self.fullyear.get())  

            colnames = [d[0].replace("_", " ") for d in cur.description]
            self.handleResult(results, colnames)

    def ReQueryRange(self):
        sql = self.getQueryByDate()
        with Connection() as conn:
            cur= conn.cursor()
            cur.execute(sql)
            results = cur.fetchall()
            colnames = [d[0].replace("_", " ") for d in cur.description]
            self.handleResult(results, colnames)

            if results:
                self.totalFile['text'] = "TOTAL: %s %s"%(
                    len(results), 
                    self.choice_entry.get() + " FROM %s TO %s"%(
                        self.from_entry.get(), self.to_entry.get()) )
            else: 
                self.totalFile['text'] = "TOTAL: 0 %s"%( 
                    self.choice_entry.get() + " FROM %s TO %s"%(
                        self.from_entry.get(), self.to_entry.get()) )  
    
    def handleResult(self, results, colnames):
        if results:
            register = results
        else:
            register = []
            colnames=[]

        self.main.destroy()
        self.main = Frame(self)
        self.tree = Treeview(self.main, colnames, self)
        self.tree.pack(expand=1, fill=BOTH)
        self.tree.set_register(register)
        self.main.pack(expand=1, fill=BOTH)

def main():
    root = Tk()
    
    app = Main(root)
    root.title("FMS")
    root.mainloop()

if __name__ == '__main__':
    main()

    
    

 
