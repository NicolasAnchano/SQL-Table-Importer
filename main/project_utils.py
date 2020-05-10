import os
import json
import pyodbc
import pandas as pd

def load_dic(dic_name):
	with open(dic_name+".json", "r") as fp:
		dic_name = json.load(fp)
	return dic_name

def get_json_directory(dic_name):
	current_folder = os.path.dirname(os.path.realpath(__file__))
	json_file = os.path.join(current_folder, dic_name)
	loaded_json = load_dic(json_file)
	return loaded_json
	
def combine_functions(*funcs):
	def combined_function(*args, **kwargs):
		for f in funcs:
			f(*args, **kwargs)
	return combined_function