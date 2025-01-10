import pandas as pd
import os
import numpy as np

# Function to plot data in to excel sheet
def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name)#, index=False)
    except Exception as e:
        print(f"From def: write_to_output_file -> An error occurred: {e}")

def create_transposed_frame_from_df(dataFrame):
    try:
        return dataFrame.T
    except Exception as e:
        print(f"From def: create_transposed_frame_from_df -> An error occurred: {e}")

def convert_list_to_dataframe_write_to_file(givenlist, givensheet, givenfile):
    try:
        df = pd.DataFrame(givenlist)
        return write_to_output_file(df=df,output_file=givenfile,output_sheet_name=givensheet)
    except Exception as e:
        print(f"From def: convert_list_to_dataframe_write_to_file -> An error occurred: {e}")
        return

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



t_df = get_all_data(input_file=i_file, input_sheet_name=i_s_name)

agent_names = t_df['Agent Name'].unique()
month_names = t_df['Month'].unique()
# print(month_names)
# dec_month = ['Dec']
# month_names = np.setdiff1d(month_names,dec_month)
# print(month_names)
#month_names.remove('Dec')
#print(month_names)
#
sliced_t_df_to_agent = return_sliced_dataFrame(input_data_frame=t_df,filter_column='Agent Name',filter_value=agent_names[0], column_names=c_nam)

month_data = {}
month_data_ratio = {}
for month_name in month_names:
    if month_name != 'Dec':
        t_df_1 = return_sliced_data_frame(sliced_t_df_to_agent,'Month',month_name)
        month_count = {}
        month_ratio = {}
        for i in range(1, len(c_nam)):
            cname = c_nam[i]
            month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
            #print(f"CNAME: {cname} and MOUNTH_COUNT: {month_count[cname]}")
            month_ratio[cname] = round((month_count[cname] * 100) / len(t_df_1))
        month_data[month_name] = month_count
        month_data_ratio[month_name] = month_ratio

#print(month_data)
t_df_2 = pd.DataFrame(month_data_ratio)
#print(t_df_2)
#print(month_data)
# print(f"Length is: {len(month_data_ratio)}")
#
# cname_avg = 0
# for mname in month_names:
#     if mname != 'Dec':
#         cname_avg += month_data_ratio[mname]['Active Listening']
#         print(month_data_ratio[mname]['Active Listening'])
# cname_avg = round(cname_avg/11)
#print(f"Avg: {cname_avg}")
t_df_2['avg'] = round(t_df_2.iloc[:,1:len(t_df_2.columns)].mean(axis=1))
print(t_df_2)
t_df_2.loc[len(t_df_2)] = ['row avg', round(t_df_2.iloc[:,2:len(t_df_2.columns)].mean())]

#print(round(t_df_2.iloc[:,1:len(t_df_2.columns)].mean()))
print(t_df_2)


#convert_list_to_dataframe_write_to_file(givenlist=month_data_ratio, givensheet=agent_names[0], givenfile=o_file)
#
# # exp_df = pd.DataFrame(month_data)
# # exp_df_2 = exp_df.T
# # exp_df_3 = exp_df_2.T
# # #print(exp_df_3)
# # write_to_output_file(exp_df_3,o_file,agent_names[0])
# #
# # for mData in month_data:
# #     mDataValue = month_data[mData]
# #     for i in range(1, len(c_nam)):
# #         mDataValue_AL = mDataValue[c_nam[i]]
# #         #print(mData,'|', c_nam[i],'|', mDataValue_AL)
# #         # print(mData)
# #         # print(c_nam[i])
# #         print(mDataValue_AL)