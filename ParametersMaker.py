import pandas
import numpy

def right(value, count):
    # To get right part of string, use negative first index in slice.
    return value[-count:]
def left(value, count):
    # To get right part of string, use positive first index in slice.
    return value[:count]

machine_name = input("Filler, Cartoner ...: ")
#create excel instance
excel_writer = pandas.ExcelWriter('Full_' + machine_name + 's.xlsx', engine='xlsxwriter')

#open csv files with parameters internal
for x in range(1, 8):
    try:
        print("HSL_" + str(x) + "_processing start")
        with open(machine_name + '_HSL' + str(x) + '_CP_Format.csv') as file_formats, open('DB_HSL' + str(x) + '_' + machine_name + '.csv') as file_DB, open(machine_name + '_HSL' + str(x) + '_language1.csv') as file_comments_1, open(machine_name + '_HSL' + str(x) + '_language2.csv') as file_comments_2:
            data_frame_param = pandas.read_csv(file_formats)
            data_frame_param["PLC_tag"] = ""
            data_frame_param["Comment_PL"] = "Brak w CSV"
            data_frame_param = data_frame_param.drop(columns = ["Format_ID", "VariableType"])

            #fill PLC_tag column (15, 44)
            data_frame_DB = pandas.read_csv(file_DB, header = None, sep = ';', names=[i for i in range(50)])
            for index_, row_ in data_frame_param.iterrows():
                for index, row in data_frame_DB.iterrows():
                    if row_[1] == row[0]:
                        if isinstance(row[44], str):
                            data_frame_param.at[index_, "PLC_tag"] = right(row[44], len(row[44])-11)
                        elif isinstance(row[15], str):
                            data_frame_param.at[index_, "PLC_tag"] = right(row[15], len(row[15])-11)
                        else:
                            data_frame_param.at[index_, "PLC_tag"] = "-xxx-"
                            print("Nie znaleciono PLC_tag dla: " + data_frame_param.at[index_, "PLC_tag"])
                        continue
            print("HSL_" + str(x) + "_tags loaded")

            #fill Comment_PL column (5)
            data_frame_com_1 = pandas.read_csv(file_comments_1, header = None, names=['name_1','type','name_2','de','en','pl'])
            data_frame_com_2 = pandas.read_csv(file_comments_2, header = None, names=['name_1','type','name_2','de','en','pl'])
            data_frame_com = pandas.concat([data_frame_com_1, data_frame_com_2])
            for index_, row_ in data_frame_param.iterrows():
                nockens_str = "LBLMM_" + row_[1]
                nockens_str = left(nockens_str, 21)
                for index, row in data_frame_com.iterrows():
                    if (row_[1] == left(row[0], len(row[0])-8)) or (nockens_str == left(row[0], len(row[0])-8)):
                        data_frame_param.at[index_, "Comment_PL"] = row[5]
                        continue
            print("HSL_" + str(x) + "_comments loaded")

            #convert to excel
            for row in data_frame_param:
                data_frame_param.to_excel(excel_writer, 'HSL_' + str(x), index = False)

    except:
        print("There is no csv from HSL_ " + str(x))

#save excel file
excel_writer.save()
