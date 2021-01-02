import openpyxl as ol
import csv
import pandas as pd
import string
import numpy as np
import os

############################################### functions ############################################################################
def save_csv(filename, result_list, header):
	# name of csv file
	header = header.split(", ")
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

# get input
input_file_1  = input("Enter the name of first input file: ")
input_file_2 = input("Enter the name of second input file: ")
filename = input("Enter the name of result file: ")
# read file
records_1 = read_excel(input_file_1)
records_2 = read_excel(input_file_2)

# get common funding information
result = set(records_1[1:]).intersection(records_2[1:])

# set result as records_2
result_union = records_2
# traverse the list
for record in records_1:
	# if not in result add it to result list
	if record not in result:
		# add 
		result_union.append(record)
	# otherwise continue
	else:
		continue

# first convert the result set into a list variable

result = list(result)


# now remove the common from union
result_union_minus_intersection = result_union

for record in result:
	if record not in result_union_minus_intersection:
		result_union_minus_intersection.append(record)

result_union_minus_intersection = [r.split(", ") for r in result_union_minus_intersection]
header = records_1[0]
# save csv
save_csv(filename.replace(".xlsx",""), result_union_minus_intersection,header)

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
