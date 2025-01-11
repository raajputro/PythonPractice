import pandas as pd
import os
import xlsxwriter
import numpy as np
from openpyxl.reader.excel import load_workbook


# Function to plot data in to excel sheet
def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name)#, index=False)
    except Exception as e:
        print(f"From def: write_to_output_file -> An error occurred: {e}")

def write_df_to_output_file(excel_writer, data_frame, output_sheet_name, start_row, start_col, index, header):
    try:
        data_frame.to_excel(excel_writer, sheet_name=output_sheet_name, startcol=start_col, startrow=start_row, index=index, header=header)
    except Exception as e:
        print(f"From def: write_df_to_output_file -> An error occurred: {e}")

def write_dfs_to_output_file(dataframes, output_file, output_sheet_name):
    try:
        s_row = 0
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_len = len(dataframes)
            for x in range(0,df_len):
                df = dataframes[x]
                dl = len(df)
                if x > 0 and x%2==0:
                    s_col = 12
                else:
                    s_col = 0
                #print(f"value of x: {x}, length of df: {dl}, value of s_col: {s_col}, value of s_row: {s_row}")
                if x > 0 and x%2==0:
                    write_df_to_output_file(excel_writer=writer, data_frame=df, output_sheet_name=output_sheet_name, start_row=s_row, start_col=s_col, index=False, header=False)
                    s_row += dl + 2
                else:
                    write_df_to_output_file(excel_writer=writer, data_frame=df, output_sheet_name=output_sheet_name, start_row=s_row, start_col=s_col, index=True, header=True)
                    s_row += dl + 1
                #print(f"value of x: {x}, length of df: {dl}, value of s_col: {s_col}, value of s_row: {s_row}")
    except Exception as e:
        print(f"From def: write_dfs_to_output_file -> An error occurred: {e}")

def create_transposed_frame_from_df(dataframe):
    try:
        return dataframe.T
    except Exception as e:
        print(f"From def: create_transposed_frame_from_df -> An error occurred: {e}")

# Parameters
execution_directory = os.getcwd()
i_file = execution_directory + "\\Data\\Profiling Master- QC.xlsx"
i_s_name = 'Main Data'
o_file = execution_directory + "\\Data\\OutputFile2.xlsx"
o_s_name = 'Scrapped_Sheet'
c_nam = ['Month','Active Listening', 'Verbal Excellence', 'Courteous Approach', 'Identification and Action for Resolution', 'Correct & Complete Information For Resolution (CCIR)', 'Avoid Rude/Unprofessional Behavior/Approach (ARU)','Ownership & Proctiveness (OP)']


# Custom functions to slice main data to specific data
def get_all_data(input_file, input_sheet_name):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df
    except Exception as e:
        print(f"An error occurred {e}")
        return -1

def return_sliced_dataFrame(input_data_frame, filter_column, filter_value, column_names):
    try:
        df = input_data_frame
        return df[df[filter_column]==filter_value][column_names]
    except Exception as e:
        print(f"An error occurred {e}")
        return -1

def return_selected_row_count(input_df,column_name,row_value):
    try:
        count = input_df[input_df[column_name] == row_value].shape[0]
        return count
    except Exception as e:
        print(f"An error occurred {e}")
        return -1

def return_sliced_data_frame(input_df, column_name, row_value):
    try:
        return input_df[input_df[column_name]==row_value]
    except Exception as e:
        print(f"An error occurred {e}")
        return -1


#######################################################################################################################
t_df = get_all_data(input_file=i_file, input_sheet_name=i_s_name)

agent_names = t_df['Agent Name'].unique()
month_names = t_df['Month'].unique()

#######################################################################################################################
sliced_t_df_to_agent = return_sliced_dataFrame(input_data_frame=t_df,filter_column='Agent Name',filter_value=agent_names[0], column_names=c_nam)

month_data = {}
month_data_ratio = {}
for month_name in month_names:
    if month_name != 'Dec':
        t_df_1 = return_sliced_data_frame(sliced_t_df_to_agent,'Month',month_name)
        td_len = len(t_df_1)
        month_count = {}
        month_ratio = {}
        for i in range(1, len(c_nam)):
            cname = c_nam[i]
            month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
            month_ratio[cname] = round((month_count[cname] * 100) / td_len)
        month_data[month_name] = month_count
        month_data_ratio[month_name] = month_ratio

################### DataFrame for Pass/Fail #######################################
t_df_11 = pd.DataFrame(month_data)
#print(t_df_11)
################### DataFrame having ratio values ################################# to test data
#print(month_data_ratio)
t_df_2 = pd.DataFrame(month_data_ratio)
t_df_2['avg'] = round(t_df_2.iloc[:,1:len(t_df_2.columns)].mean(axis=1))
#print(t_df_2)

###################################################################################

com_cnames = ['Verbal Excellence','Avoid Rude/Unprofessional Behavior/Approach (ARU)']
emp_cnames = ['Courteous Approach','Active Listening']
bus_cnames = ['Correct & Complete Information For Resolution (CCIR)','Identification and Action for Resolution']
acc_cnames = ['Ownership & Proctiveness (OP)']

def provide_df(month_names, cnames, month_data_ratio):
    sliced_data = {}
    s_data = {}
    mn_len = len(month_names)
    cn_len = len(cnames)
    for y in range(0,mn_len):
        mname = month_names[y]
        if mname != 'Dec':
            #print(f"value of month: {mname} \n", f"{month_data_ratio[mname]}")
            for z in range(0, cn_len):
                cnam = cnames[z]
                val = month_data_ratio[mname][cnam]
                s_data[cnam] = val
                sliced_data[mname]=s_data
                print(f"value of Z: {z}")
        print(f"value of YYY: {y}")
    print(f"Sliced Data: \n",sliced_data)
    t_df_3 = pd.DataFrame(sliced_data)
    print("t_df_3\n")
    print(t_df_3)
    t_df_3['avg'] = round(t_df_3.iloc[:, 1:len(t_df_3.columns)].mean(axis=1))

    avg_val1 = round(float(t_df_3['avg'].mean()),2)
    avg_val2 = round(float(avg_val1/20),2)
    avg_df = pd.DataFrame({ 'values': [avg_val1, avg_val2] })
    return [t_df_3, avg_df]

com_df = provide_df(month_data_ratio=month_data_ratio,month_names=month_names,cnames = com_cnames)
emp_df = provide_df(month_data_ratio=month_data_ratio,month_names=month_names,cnames = emp_cnames)
bus_df = provide_df(month_data_ratio=month_data_ratio,month_names=month_names,cnames = bus_cnames)
acc_df = provide_df(month_data_ratio=month_data_ratio,month_names=month_names,cnames = acc_cnames)
# #
dfs = [t_df_11, com_df[0], com_df[1], emp_df[0], emp_df[1], bus_df[0], bus_df[1], acc_df[0], acc_df[1]]
# #
write_dfs_to_output_file(dataframes=dfs, output_file=o_file, output_sheet_name=agent_names[0])