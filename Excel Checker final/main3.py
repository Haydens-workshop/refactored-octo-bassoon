import openpyxl as ol
import csv
import pandas as pd
import string



import numpy as np
import os
############################################### functions ############################################################################




def save_csv(filename, result_list, header):
	# name of csv file
	# writing to csv file  
	with open(filename + ".csv", 'w') as csvfile:  
		# creating a csv writer object  
		csvwriter = csv.writer(csvfile)      
		# writing the data rows
		csvwriter.writerow(header)		
		csvwriter.writerows(result_list)
	

def read_excel(filename):
	workbook = ol.load_workbook(filename)
	sheet = workbook.active
	no_of_rows = len(list(sheet.rows))
	no_of_columns = len(list(sheet.columns))
	records = []
	# get all records
	for i in range(1,no_of_rows+1):
		row = []
		for j in range(1, no_of_columns+1):
			row.append(str(sheet.cell(row=i,column=j).value))
		records.append(", ".join(row))
	# return the list
	return records
############################################### funcitons complete ##################################################################

folder_path = input("Enter path to folder: ").replace("\\","\\\\")
filename = input("Enter the name of result file: ")
print("************** Menu ****************")
print("1) Get all common in all excel files")
print("2) Get all unique in all excel files")
choice = int(input("Enter your choice: "))

# get all files in current folder
files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

# to store the records
records = []
# read file
for i in range(0, len(files)):
	if i == 0:
		records.extend([read_excel(folder_path + "\\\\" + files[i])])
	else:
		records.extend([read_excel(folder_path + "\\\\" + files[i])[1:]])

# make it as list of lists to intersection and union operation on it
records = [r for r in records]

if choice == 1:
	# get common intersection
	result = list(set(records[0]).intersection(*records))
elif choice == 2:
	# get union unique
	result = list(set(records[0]).union(*records))

if len(result) == 0:
	print("No common records present in all files")
else:
	result = [r.split(", ") for r in result]
	header = result[0]
	# save csv
	save_csv(filename.replace(".xlsx",""), result,header)
	# Reading the csv file 
	df_new = pd.read_csv(filename.replace(".xlsx","") + '.csv') 
	# saving xlsx file 
	GFG = pd.ExcelWriter(filename) 
	df_new.to_excel(GFG, index = False)   
	GFG.save()
	try:
		os.system("rm " + filename.replace("xlsx","csv"))
	except:
		print("")
