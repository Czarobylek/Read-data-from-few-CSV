import pandas
import numpy
from openpyxl import load_workbook

def ListDif(list_1, list_2):
    #Differencess beetwen two lists
    dif = set(list_1) - set(list_2)
    return list(dif)

machine_name = input("Filler, Cartoner ...: ")
#create excel instance and vcopy existing data
book = load_workbook("Full_" + machine_name + "s.xlsx")
excel_writer = pandas.ExcelWriter('Full_' + machine_name + 's.xlsx', engine='openpyxl')
excel_writer.book = book

#read from excel and sum in one data_frame
y = 1
for x in range(1, 8):
    try:
        parameters_part = pandas.read_excel('Full_' + machine_name + 's.xlsx', sheet_name="HSL_" + str(x))
        parameters_part["Line_number"] = ""
        for index, row in parameters_part.iterrows():
            parameters_part.at[index, "Line_number"] = x
        if y > 1:
            parameters_all = pandas.concat([parameters_all, parameters_part])
        else:
            parameters_all = parameters_part
        y = y + 1
    except:
        print("There is no csv from HSL_ " + str(x))

#create PIVOT with all parameters
parameters_all_pivot = parameters_all.copy()
parameters_all_pivot["ActionParam"] = parameters_all_pivot["ActionParam"] + " (" + parameters_all["Comment_PL"] + ")"
parameters_all_pivot = parameters_all_pivot.pivot_table(index = ['Format','Line_number'], columns = 'ActionParam', values = 'SavedValue')
for row in parameters_all_pivot:
    parameters_all_pivot.to_excel(excel_writer, sheet_name = 'PIVOT')
print("PIVOT table created")

#create PIVOT_v2 with parameters find on HMI displays
mytxt = open(machine_name + "_HMI.txt")
parameters_HMI = []
for line in mytxt:
    line = line.strip('\n')
    parameters_HMI.append(line)

parameters_all_pivot_v2 = parameters_all.copy()
parameters_all_pivot_v2 = parameters_all_pivot_v2.pivot_table(index = ['Format','Line_number'], columns = 'ActionParam', values = 'SavedValue')
columns_names = []
for column in parameters_all_pivot_v2:
    columns_names.append(column)

columns_drop = ListDif(columns_names, parameters_HMI)
parameters_all_pivot_v2 = parameters_all_pivot_v2.drop(columns = columns_drop)

for row in parameters_all_pivot_v2:
    parameters_all_pivot_v2.to_excel(excel_writer, sheet_name = 'PIVOT_v2')
print("PIVOT_v2 table created")

#create PIVOT_finall with parameters
mytxt = open(machine_name + "2_HMI.txt")
parameters_HMI = []
for line in mytxt:
    line = line.strip('\n')
    parameters_HMI.append(line)

parameters_all_pivot_v3 = parameters_all.copy()
parameters_all_pivot_v3 = parameters_all_pivot_v3.pivot_table(index = ['Format','Line_number'], columns = 'ActionParam', values = 'SavedValue')
columns_names = []
for column in parameters_all_pivot_v3:
    columns_names.append(column)

columns_drop = ListDif(columns_names, parameters_HMI)
parameters_all_pivot_v3 = parameters_all_pivot_v3.drop(columns = columns_drop)

for row in parameters_all_pivot_v3:
    parameters_all_pivot_v3.to_excel(excel_writer, sheet_name = 'PIVOT_v3')
print("PIVOT_v3 table created")

#create PIVOT_PLC with tags
parameters_all_pivot_v4 = parameters_all.copy()
parameters_all_pivot_v4 = parameters_all_pivot_v4.pivot_table(index = ['Format','Line_number'], columns = 'ActionParam', values = 'PLC_tag', aggfunc='first')
columns_names = []
for column in parameters_all_pivot_v4:
    columns_names.append(column)

parameters_all_pivot_v4 = parameters_all_pivot_v4.drop(columns = columns_drop)

for row in parameters_all_pivot_v4:
    parameters_all_pivot_v4.to_excel(excel_writer, sheet_name = 'PIVOT_PLC')
print("PIVOT_PLC table created")

#create PIVOT comments
parameters_all_pivot_v5 = parameters_all.copy()
parameters_all_pivot_v5 = parameters_all_pivot_v5.pivot_table(index = ['Format','Line_number'], columns = 'ActionParam', values = 'Comment_PL', aggfunc='first')
columns_names = []
for column in parameters_all_pivot_v5:
    columns_names.append(column)

parameters_all_pivot_v5 = parameters_all_pivot_v5.drop(columns = columns_drop)

for row in parameters_all_pivot_v5:
    parameters_all_pivot_v5.to_excel(excel_writer, sheet_name = 'PIVOT_Comments')
print("PIVOT_Comments table created")

#save excel file
excel_writer.save()
