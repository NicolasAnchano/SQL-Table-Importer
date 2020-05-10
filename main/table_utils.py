import pyodbc
import pandas as pd

def connection(server, database, trusted, uid, pwd, driver="{SQL Server}"):
	"""
	driver (str) - Client, be it MSSQL, MySQL, etc.
	server (str) - Server name.
	trusted (str) - If it's a trusted connection
	uid (str) - Username
	pwd (str) - Password
	"""
	global conn
	conn = pyodbc.connect("""
						  Driver={};
						  Server={};
						  Database={};
						  Trusted_Connection={};
						  UID={};
						  PWD={};
						  """.format(driver, server, database, trusted, uid, pwd))
	global cursor
	cursor = conn.cursor()

def load_tables():
	global table_names
	table_names = []
	for table in cursor.tables():
		if "tbl_" in table.table_name:
			table_names.append(table.table_name)
	cursor.cancel()
	
	global table_dict
	table_dict = {}
	for table in table_names:
		data = pd.read_sql("SELECT * FROM "+table, conn)
		df = pd.DataFrame(data)
		table_dict[str(table[4:])] = df # Removes "tbl_" from names

	conn.close()
	#return table_dict

def table_to_xlsx(sel_table="All"):
	"""
	sel_table (str) - specific table to export - default exports all.
	"""
	if sel_table != "All":
		table_dict[sel_table].to_excel(sel_table+".xlsx", index=False)
	else:
		for table_name, table_df in table_dict.items():
			table_df.to_excel(table_name+".xlsx", index=False)
		
def table_to_csv(sel_table="All"):
	"""
	sel_table (str) - specific table to export - default exports all.
	"""
	if sel_table != "All":
		table_dict[sel_table].to_csv(sel_table+".csv", index=False)
	else:
		for table_name, table_df in table_dict.items():
			table_df.to_csv(table_name+".csv", index=False)
		
def create_format_demand(n_week):
	"""
	num_of_week (int) - Week number to filter the demand and inventory data - Only the number as input
	"""
	num_of_week = "WEEK " + str(n_week)	
	# LOADING DATA

	demand_sql = table_dict["Demand"]
	demand_sql = demand_sql[demand_sql["Week"] == num_of_week]

	product_inv_sql = table_dict["Product Inventory"]
	product_inv_sql = product_inv_sql[product_inv_sql["Week"] == num_of_week]

	component_inv_sql = table_dict["Components Inventory"]
	component_inv_sql = component_inv_sql[component_inv_sql["Week"] == num_of_week]

	locations = get_json_directory("locations")
	product_dict = get_json_directory("product_dict")
	formulas = get_json_directory("formulas")
	product_client = get_json_directory("product_client")
	
	#DEMAND
	
	## CREATING DEMAND
	dem_data = {
		"Product": [prd for prd in demand_sql["Product"].tolist()],
		"Description": ["null"] * len(demand_sql),
		"Customer": ["null"] * len(demand_sql),
		"Formula": ["null"] * len(demand_sql),
		"Inventory": [0] * len(demand_sql),
		"Priority Product": [0] * len(demand_sql),
		"Priority": [0] * len(demand_sql),
		"Initial Date": [rmd for rmd in demand_sql["Initial_Date"].tolist()],
		"Demand (Pounds)": [dem for dem in demand_sql["Measure"].tolist()],
		"Due Date": [dd for dd in demand_sql["Due_Date"].tolist()],
		"Location": [loc for loc in demand_sql["Location"].tolist()]
		}
	demand = pd.DataFrame(dem_data)
	
	## FORMAT
	demand["Product"] = demand["Product"].astype(str)
	demand["Initial Date"] = pd.to_datetime(demand["Initial Date"], format="%Y-%m-%d")
	demand["Due Date"] = pd.to_datetime(demand["Due Date"], format="%Y-%m-%d")
	
	for i_d, description in locations.items():
		demand.loc[demand["Location"] == i_d, "Location"] = description
		
	## ADDING DATA
	for i_d, data in product_dict.items():
		for prop, value in data.items():
			if prop == "desc":
				demand.loc[demand["Product"] == i_d, "Description"] = value
				
	for i_d, formula in formulas.items():
		demand.loc[demand["Product"] == i_d, "Formula"] = formula
		
	for i_d, client in product_client.items():
		demand.loc[demand["Product"] == i_d, "Customer"] = client

	for ix, row in product_inv_sql.iterrows():
		prd = row["Product"]
		inv = row["Measure"]
		demand.loc[demand["Product"] == prd, "Inventory"] = inv
	
	# COMPONENTS INVENTORY
	comp_data = {
		"Component": [comp for comp in component_inv_sql["Component"].tolist()],
		"Inventory": [inv for inv in component_inv_sql["Measure"].tolist()],
		"Location": [loc for loc in component_inv_sql["Location"].tolist()]
		}
		
	comp_inv = pd.DataFrame(comp_data)
	
	for i_d, description in locations.items():
		comp_inv.loc[comp_inv["location"] == i_d, "location"] = description
	
	# EMPTY TAB REQUIRED FOR THE MODEL TO WORK - NULL POINTER EXCEPTION OTHERWISE
	machines = pd.DataFrame(columns=["Machine", "from", "to"])
	
	# DUMPING DATA
	with pd.ExcelWriter("Demand.xlsx") as writer:  
		demand.to_excel(writer, sheet_name="Demand", index=False)
		comp_inv.to_excel(writer, sheet_name="Components Inventory", index=False)
		machines.to_excel(writer, sheet_name="Machines schedule", index=False)