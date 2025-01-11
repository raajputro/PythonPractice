import os
import pandas as pd

# Function to plot data in to excel sheet
def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name)#, index=False)
    except Exception as e:
        print(f"From def: write_to_output_file -> An error occurred: {e}")

def write_df_to_output_file(excel_writer, data_frame, output_sheet_name, start_row, start_col, index, header):
    try:
        data_frame.to_excel(excel_writer, sheet_name=output_sheet_name, startcol=start_col, startrow=start_row,
                            index=index, header=header)
    #     #excel_writer.book[output_sheet_name]
    #     worksheet = excel_writer.book[output_sheet_name]
    #     worksheet.delete_cols(1, worksheet.max_column)
    # except KeyError:
    #     worksheet = excel_writer.book.create_sheet(output_sheet_name)
    # data_frame.to_excel(excel_writer, sheet_name=output_sheet_name, startcol=start_col, startrow=start_row, index=index,
    #                     header=header)

    except Exception as e:
        print(f"From def: write_df_to_output_file -> An error occurred: {e}")


def write_dfs_to_output_file(dataframes, output_file, output_sheet_name):
    try:
        s_row = 0
        with (pd.ExcelWriter(output_file, engine='openpyxl') as writer):
            df_len = len(dataframes)
            for x in range(0,df_len):
                df = dataframes[x]
                dl = len(df)
                if x > 0 and x%2==0:
                    s_col = 12
                else:
                    s_col = 0
                #print(f"value of x: {x}, length of df: {dl}, value of s_col: {s_col}, value of s_row: {s_row}")
                if x == 0:
                    write_df_to_output_file(excel_writer=writer, data_frame=df, output_sheet_name=output_sheet_name,
                                            start_row=s_row, start_col=s_col, index=True, header=True)
                    s_row += dl + 5
                elif x > 0 and x%2==0:
                    write_df_to_output_file(excel_writer=writer, data_frame=df, output_sheet_name=output_sheet_name,
                                            start_row=s_row, start_col=s_col, index=False, header=False)
                    s_row += dl + 2
                else:
                    write_df_to_output_file(excel_writer=writer, data_frame=df, output_sheet_name=output_sheet_name,
                                            start_row=s_row, start_col=s_col, index=True, header=True)
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
c_nam = ['Month','Active Listening', 'Verbal Excellence', 'Courteous Approach',
         'Identification and Action for Resolution', 'Correct & Complete Information For Resolution (CCIR)',
         'Avoid Rude/Unprofessional Behavior/Approach (ARU)','Ownership & Proctiveness (OP)']


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
agent_name = agent_names[0]
sliced_t_df_to_agent = return_sliced_dataFrame(input_data_frame=t_df, filter_column='Agent Name',
                                               filter_value=agent_name, column_names=c_nam)
month_data = {}
month_data_ratio = {}
for month_name in month_names:
    if month_name != 'Dec':
        t_df_1 = return_sliced_data_frame(sliced_t_df_to_agent,'Month',month_name)
        td_len = len(t_df_1)
        if td_len == 0:
        #    print(f"for agent name: {agent_name}, month: {month_name}, length: {td_len} and base data frame length: {stdta_len}")
            td_len = 1
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
################### DataFrame having ratio values ################################# to test data
t_df_2 = pd.DataFrame(month_data_ratio)
t_df_2['avg'] = round(t_df_2.iloc[:,1:len(t_df_2.columns)].mean(axis=1))
####################################################################################################################
def slice_month_data_ratio(month_data_ratio, month_names, attributes_name):
    sliced_month_data = {}
    for mnam in month_names:
        if mnam != 'Dec':
            sm_data = {}
            for anam in attributes_name:
                val = month_data_ratio[mnam]
                val2 = val[anam]
                sm_data[anam] = val2
            sliced_month_data[mnam] = sm_data
    test_df = pd.DataFrame(sliced_month_data)
    test_df['avg'] = round(test_df.iloc[:, 1:len(test_df.columns)].mean(axis=1))
    avg_val1 = round(float(test_df['avg'].mean()),2)
    avg_val2 = round(float(avg_val1/20),2)
    avg_df = pd.DataFrame({ 'values': [avg_val1, avg_val2] })
    return [test_df, avg_df]

####################################################################################################################
#creating data frames array, according to following attribute groups
####################################################################################################################
com_cnames = ['Verbal Excellence', 'Avoid Rude/Unprofessional Behavior/Approach (ARU)']
emp_cnames = ['Courteous Approach', 'Active Listening']
bus_cnames = ['Correct & Complete Information For Resolution (CCIR)', 'Identification and Action for Resolution']
acc_cnames = ['Ownership & Proctiveness (OP)']
dfs_array = [t_df_11]
smdr = slice_month_data_ratio(month_data_ratio,month_names,com_cnames)
for x in smdr:
    dfs_array.append(x)
    #print(f"Test DF: \n {x}")
smdr = slice_month_data_ratio(month_data_ratio,month_names,emp_cnames)
for x in smdr:
    dfs_array.append(x)
smdr = slice_month_data_ratio(month_data_ratio,month_names,bus_cnames)
for x in smdr:
    dfs_array.append(x)
smdr = slice_month_data_ratio(month_data_ratio,month_names,acc_cnames)
for x in smdr:
    dfs_array.append(x)
####################################################################################################################
# finally, write all data frames to work sheet
####################################################################################################################
agent_name = agent_name[:30]
write_dfs_to_output_file(dataframes=dfs_array, output_file=o_file, output_sheet_name=agent_name)