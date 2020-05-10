import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from project_utils import *
from table_utils import *

LARGE_FONT = ("Verdana", 12)
SMALL_FONT = ("Verdana", 10)
TABLES = ["All", "Bags", "Components_Inventory", "Components", "Demand", "Products", "Products_Inventory", "MD_Components", "MD_Customer", "MD_Formulas", "MD_Products", "MD_Machines"]


# FUNCTIONS
def create_window():
	new_window = Toplevel(window)

# GUI
class MainApp(tk.Tk):
	def __init__(self, *args, **kwargs):
		tk.Tk.__init__(self, *args, **kwargs)
		tk.Tk.iconbitmap(self, default="assets/project_icon.ico") # has to be ico
		tk.Tk.wm_title(self, "SQL Tables Importer")

		container = tk.Frame(self)
		container.grid(row=0, column=0, sticky="nsew")

		self.frames = {}
		for f in (StartPage, TablesPage):
			frame = f(container, self)
			self.frames[f] = frame
			frame.grid(row=1, column=0, sticky="nsew")

		self.show_frame(StartPage)

	def show_frame(self, cont):
		frame = self.frames[cont]
		frame.tkraise()

class StartPage(tk.Frame):
	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)

		self.start_image_label = tk.Label(self)
		self.start_image_label.image = tk.PhotoImage(file="assets/project_image2.png")
		self.start_image_label["image"] = self.start_image_label.image
		self.start_image_label.grid(row=0, column=0, sticky="w")

		start_label = tk.Label(self, text="Connect to Database", font=LARGE_FONT)
		start_label.grid(row=1, column=0, sticky="w")

		# ENTRIES	
		server_label = tk.Label(self, text="Enter Server name", fg="black", font=SMALL_FONT).grid(row=2, column=0, sticky="w")
		server_entry = tk.Entry(self, width=30, bg="white")
		server_entry.grid(row=2, column=1, sticky="w")

		database_label = tk.Label(self, text="Enter Database name", fg="black", font=SMALL_FONT).grid(row=3, column=0, sticky="w")
		database_entry = tk.Entry(self, width=30, bg="white")
		database_entry.grid(row=3, column=1, sticky="w")

		trusted_label = tk.Label(self, text="Trusted Connection", fg="black", font=SMALL_FONT).grid(row=4, column=0, sticky="w")
		trusted_entry = tk.Entry(self, width=30, bg="white")
		trusted_entry.grid(row=4, column=1, sticky="w")

		user_label = tk.Label(self, text="Enter Username", fg="black", font=SMALL_FONT).grid(row=5, column=0, sticky="w")
		user_entry = tk.Entry(self, width=30, bg="white")
		user_entry.grid(row=5, column=1, sticky="w")

		pwd_label = tk.Label(self, text="Enter Password", fg="black", font=SMALL_FONT).grid(row=6, column=0, sticky="w")
		pwd_entry = tk.Entry(self, width=30, bg="white", show="*")
		pwd_entry.grid(row=6, column=1, sticky="w")

		# BUTTON
		## lambda allows the function to be executed on button press instead of on program startup
		connect_button = ttk.Button(self, text="Connect", command=lambda: combine_functions(connection(server_entry.get(), database_entry.get(), trusted_entry.get(), user_entry.get(), pwd_entry.get()), load_tables(), controller.show_frame(TablesPage)))
		connect_button.grid(row=7, column=1, sticky="w")

class TablesPage(tk.Frame):
	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)

		self.tables_image_label = tk.Label(self)
		self.tables_image_label.image = tk.PhotoImage(file="assets/project_image2.png")
		self.tables_image_label["image"] = self.tables_image_label.image
		self.tables_image_label.grid(row=0, column=0, sticky="w")

		tables_label = tk.Label(self, text="Import SQL Tables", fg="black", font=LARGE_FONT)
		tables_label.grid(row=1, column=0, sticky="w")

		demand_label = tk.Label(self, text="Create Demand", fg="black", font=LARGE_FONT).grid(row=2, column=0, sticky="w")

		self.demand_folder_path = tk.StringVar(self)
		demand_output_path = str(self.demand_folder_path)
		demand_browse_entry = tk.Entry(self, textvariable=self.demand_folder_path, width=30, bg="white")
		demand_browse_entry.grid(row=3, column=1, sticky="w")
		demand_browse_button = ttk.Button(self, text="Choose Folder", command=lambda: self.demand_browse_folder())
		demand_browse_button.grid(row=3, column=2, sticky="w")

		demand_entry_label = tk.Label(self, text="Week Number")
		demand_entry_label.grid(row=4, column=0, sticky="w")
		demand_entry = tk.Entry(self, width=30, bg="white")
		demand_entry.grid(row=4, column=1, sticky="w")
		demand_button = ttk.Button(self, text="Import", command=lambda: create_format_demand(demand_entry.get()))
		demand_button.grid(row=4, column=2, sticky="w")

		xlsx_label = tk.Label(self, text="Import Tables", fg="black", font=LARGE_FONT).grid(row=5, column=0, sticky="w")
		xlsx_entry = tk.Entry(self, width=30, bg="white")

		self.xlsx_folder_path = tk.StringVar(self)
		xlsx_output_path = str(self.xlsx_folder_path)
		xlsx_browse_button = ttk.Button(self, text="Choose Folder", command=lambda: self.xlsx_browse_folder())
		xlsx_browse_button.grid(row=6, column=2, sticky="w")
		xlsx_browse_entry = tk.Entry(self, textvariable=self.xlsx_folder_path, width=30, bg="white")
		xlsx_browse_entry.grid(row=6, column=1, sticky="w")

		xlsx_chosen_file = tk.StringVar(self)
		xlsx_chosen_file.set("All")
		xlsx_dropdown = tk.OptionMenu(self, xlsx_chosen_file, *TABLES)
		xlsx_dropdown.grid(row=6, column=0, sticky="w")
		xlsx_button = ttk.Button(self, text="Import", command=lambda: table_to_xlsx(xlsx_chosen_file.get()))
		xlsx_button.grid(row=7, column=2, sticky="w")

	def xlsx_browse_folder(self):
		xlsx_folder_selected = filedialog.askdirectory(initialdir=".")
		self.xlsx_folder_path.set(xlsx_folder_selected)
		os.chdir(xlsx_folder_selected)

	def demand_browse_folder(self):
		demand_folder_selected = filedialog.askdirectory(initialdir=".")
		self.demand_folder_path.set(demand_folder_selected)
		os.chdir(demand_folder_selected)

app = MainApp()

app.mainloop()