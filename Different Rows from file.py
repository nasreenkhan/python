import pandas as pd #Excel Workbook Operations

import easygui # GUI operations

easygui.msgbox("select Intital Workbook")
intial_workbook = easygui.fileopenbox()
df_intial = pd.read_excel(intial_workbook,header =1)
#header = 1 is to skip the first row in excel
#This is Optional in my file the first row is always empty.

#print("Initial Book:", intial_workbook)
print("Total Rows: {0}".format(len(df_initail)))

easygui.msgbox("select Latest Workbook")
info_workbook = easygui.fileopenbox()
df_info = pd.read_excel(info_workbook,header =1)
#header = 1 is to skip the first row in excel
#This is Optional in my file the first row is always empty.

#print("Info Book:", info_workbook)
print("Total Rows: {0}".format(len(df_info)))

output_workbook = easygui.enterbox(msg='Enter the Output File Name')
output_workbook = output_workbook +".xlsx"

#print("Output Workbook Name: ", output_workbook)


df_Output = pd.merge(df_initial, df_info, on=['Employee No'], how="outer", indicator=True)
df_Output = df_Output[df_Output['_merge'] == 'rightonly']

print("Total Rows: {0}".format(len(df_Output)))

#function to remove _x valued columns from Result Excel
def drop_x(df):
    #list comprehension of the cols that end with '_x'
    to_drop = [x for x in df if x.endswith('_x')]
    df.drop(to_drop, axis=1, inplace=True)

drop_x(df_Output)

#function to rename the value with actual column Names in Excel
def rename_y(df):
    for col in df:
        if col.endswith('_y'):
           df.rename(columns={col:col.rstrip('y')}, inplace=True)

rename_y(df_Output)

#Remvoing the last column which will have "_merge"
df_output.drop(df_output.columns[len(def_4.columns)-1], axis=1, inplace=True)

df_output.to_excel(output_workbook)


# Autofit Excel Rows in Current Sheet
writer = pd.ExcelWriter(output_workbook)
df_output.to_excel(writer, sheet_name='Talent Details', index=False)

workbook = writer.book
worksheet = writer.sheets['Talent Details']

#iterate through each column and set the width == the max length in that column. A padding length of 2 is also added
for i, col in enumerate(df_output):
    #find length of column i
    column_len = df_output[col].astype(str).str.len().max()
    #selecting the length if the column header is larger
    #then the max column value length
    column_len = max(column_len, len(col)) +2
    #set the column length
    worksheet.set_column(i,i,column_len)

writer.save()

easygui.msgbox("Operations Completed Successfully")
