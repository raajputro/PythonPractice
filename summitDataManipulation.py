import pandas as pd

def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name)#, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

execution_directory = "C:\\Users\\mahbu\\Desktop\\Practices"
i_file = execution_directory + "\\Profiling Master- QC.xlsx"
i_s_name = 'Main Data'
o_file = execution_directory + "\\OutputFile2.xlsx"
o_s_name = 'Scrapped_Sheet'

c_nam = ['Month', 'Active Listening', 'Verbal Excellence', 'Courteous Approach', 'Identification and Action for Resolution', 'Correct & Complete Information For Resolution (CCIR)', 'Avoid Rude/Unprofessional Behavior/Approach (ARU)','Ownership & Proctiveness (OP)']

def return_selected_columns_of_selected_row(input_file, input_sheet_name ,filter_column,filter_value, c_names):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[filter_column]==filter_value][c_names]
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
        #return (input_df[column_name]==row_value).sum()
        count = input_df[input_df[column_name] == row_value].shape[0]
        return count
    except Exception as e:
        print(f"An error occurred {e}")
        return -1

t_df = return_selected_columns_of_selected_row(input_file=i_file,input_sheet_name=i_s_name,filter_column='Agent Name',filter_value='Shusmoy Kundu',c_names=c_nam)
#print(t_df)
#write_to_output_file(df=t_df,output_file=o_file,output_sheet_name=o_s_name)

month_names = t_df['Month'].unique()
#print(month_names)

month_data = {}
for month_name in month_names:
    t_df_1 = return_sliced_data_frame(t_df,'Month',month_name)
    #write_to_output_file(t_df_1,o_file,output_sheet_name=month_name)
    #print(t_df_1)
    month_count = {}
    for i in range(1, len(c_nam)):
        cname = c_nam[i]
        #print(cname)
        month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
        month_data[month_name] = month_count
    print(month_data)
#print(month_data)


#write_to_output_file(df=t_df,output_file=o_file,output_sheet_name=o_s_name)

# month_data = {}
# for month_name in month_names:
#     t_df_1 = return_sliced_data_frame(t_df,'Month',month_name)
#     #write_to_output_file(df=t_df_1,output_file=o_file,output_sheet_name=o_s_name)
#
#     month_count = {}
#     # for cname in c_nam:
#     #     if cname == 'Month':
#     #         month_count[cname]= return_selected_row_count(input_df=t_df_1,column_name=cname,row_value='Jan')
#     #     else:
#     #         month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
#     for i in range(1, len(c_nam)):
#         cname = c_nam[i]
#         month_count[cname] = return_selected_row_count(input_df=t_df_1, column_name=cname, row_value='Pass')
#     month_data[month_name] = month_count
# print(month_data)
# #write_to_output_file(df=t_df,output_file=o_file,output_sheet_name=o_s_name)