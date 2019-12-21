import pandas as pd #Excel Workbook Operations

import easygui # GUI operations

import xlrd

easygui.msgbox("select Workbook")
intial_workbook = easygui.fileopenbox()
df_intial = pd.read_excel(intial_workbook)
#header = 1 is to skip the first row in excel
#This is Optional in my file the first row is always empty.

#print("Initial Book:", intial_workbook)
print("Total Rows: {0}".format(len(df_initail)))

output_workbook = easygui.enterbox(msg='Enter the Output File Name')
output_workbook = output_workbook +".xlsx"

#print("Output Workbook Name: ", output_workbook)

df_head = df_intial.head(0)
header_list = []
for col in df_head.columns:
    header_list.append(col)

sheet_data = []
wb = xlrd.open_workbook(intial_workbook)
p = wb.sheet_names()
for y in p:
    sh = wb.sheet_by_name(y)
    for rownum in range(sh.nrows):
        sheet_data.append((sh.row_values(rownum)))

#index 7 is column value in Excel sheet for Skill Select
count = 1

#Java Skill
found_list = []
found_list.append(header_list)
rows_to_be_saved = []
for i in sheet_data:
    if i[7].strip().casefold() == 'Java'.casefold() or i[7].strip().casefold() == 'Core Java'.casefold() or i[7].strip().casefold() == 'Advance Java'.casefold() or i[7].strip().casefold() == 'J2SE'.casefold() or i[7].strip().casefold() == 'J2EE'.casefold() or i[7].strip().casefold() == 'Spring'.casefold() or i[7].strip().casefold() == 'JSP'.casefold() or i[7].strip().casefold() == 'Java Server Pages'.casefold() or i[7].strip().casefold() == 'Servlets'.casefold():
        found_listappend(i)
        count = count+1
    else:
        rows_to_be_saved.append(i)

print("Java Skill Count : ", count)

#DotNet Skill
count = 1
net_found_list = []
net_found_list.append(header_list)
rows_to_be_saved = []
for i in sheet_data:
    if i[7].strip().casefold() == 'ASP'.casefold() or i[7].strip().casefold() == 'ASP.Net Web API'.casefold() or i[7].strip().casefold() == 'ASP.NET'.casefold() or i[7].strip().casefold() == 'ADO.NET'.casefold() or i[7].strip().casefold() == 'C#'.casefold() :
        net_found_listappend(i)
        count = count+1
    else:
        rows_to_be_saved.append(i)

print("Dot Net Skill Count : ", count)

writer = pd.ExcelWriter(output_workbook)
pd.DataFrame(found_list).to_excel(writer,sheet_name='Java Skills', header=False, index=False)
pd.DataFrame(net_found_list).to_excel(writer,sheet_name='Dot Net Skills', header=False, index=False)
writer.save()

easygui.msgbox("Operations Completed Successfully")
