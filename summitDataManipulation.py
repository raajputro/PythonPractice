import pandas as pd
import os
from numpy import *

# Function to plot data in to excel sheet
def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name)#, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

# Parameters
execution_directory = os.getcwd()
i_file = execution_directory + "\\Data\\Profiling Master- QC.xlsx"
i_s_name = 'Main Data'
o_file = execution_directory + "\\OutputFile2.xlsx"
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

def return_selected_columns_of_selected_row(input_file, input_sheet_name ,filter_column,filter_value, c_names):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[filter_column]==filter_value][c_names]
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

def return_sliced_data_frame(input_df, column_name, row_value):
    try:
        return input_df[input_df[column_name]==row_value]
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

t_df = return_selected_columns_of_selected_row(input_file=i_file,input_sheet_name=i_s_name,filter_column='Agent Name',filter_value='Shusmoy Kundu',c_names=c_nam)
#print(t_df)

mainData_df = get_all_data(input_file=i_file, input_sheet_name=i_s_name)

agent_names = mainData_df['Agent Name'].unique()

month_names = t_df['Month'].unique()

month_data = {}
for month_name in month_names:
    t_df_1 = return_sliced_data_frame(t_df,'Month',month_name)
    month_count = {}
    for i in range(1, len(c_nam)):
        cname = c_nam[i]
        month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
    month_data[month_name] = month_count

for mData in month_data:
    mDataValue = month_data[mData]
    for i in range(1, len(c_nam)):
        mDataValue_AL = mDataValue[c_nam[i]]
        print(mData,'|', c_nam[i],'|', mDataValue_AL)