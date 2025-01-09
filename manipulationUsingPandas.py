import pandas as pd

def create_new_sheet(input_file, input_sheet_name, output_file, output_sheet_name, row_number, column_indices):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        selected_data = df.iloc[row_number, column_indices]
        new_df = pd.DataFrame([selected_data], columns=df.columns[column_indices])
        with pd.ExcelWriter(output_file) as writer:
            new_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

def create_new_sheets(input_file, input_sheet_name, output_file, output_sheet_name, row_indices, column_indices):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        selected_data = df.iloc[row_indices, column_indices]
        new_df = pd.DataFrame(selected_data, columns=df.columns[column_indices])
        with pd.ExcelWriter(output_file) as writer:
            new_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

def create_data_frame(input_file, input_sheet_name, row_indices, column_indices):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        selected_data = df.iloc[row_indices, column_indices]
        return pd.DataFrame(selected_data, columns=df.columns[column_indices])
    except Exception as e:
        print(f"An error occurred: {e}")

def create_transposed_frame(input_file, input_sheet_name, row_indices, column_indices):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        selected_data = df.iloc[row_indices, column_indices]
        df1= pd.DataFrame(selected_data, columns=df.columns[column_indices])
        return df1.T
    except Exception as e:
        print(f"An error occurred: {e}")

def create_transposed_frame_from_df(dataFrame):
    try:
        return dataFrame.T
    except Exception as e:
        print(f"An error occurred: {e}")

def write_to_output_file(df, output_file, output_sheet_name):
    try:
        with pd.ExcelWriter(output_file) as writer:
            df.to_excel(writer, sheet_name=output_sheet_name, index=False)
    except Exception as e:
        print(f"An error occurred: {e}")

def get_row_index(input_file, input_sheet_name, search_value, search_column):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[search_column] == search_value].index[0]
        #return df[(df==search_value).any(axis=1)]
        #return df.loc[df[search_column] == search_value]
    except IndexError:
        return -1

def get_column_index(input_file, input_sheet_name, column_name):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df.columns.get_loc(column_name)
    except KeyError:
        return -1

def find_duplicate_rows_in_column(input_file, input_sheet_name, column_name):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[column_name].duplicated(keep=False)]
    except Exception as e:
        return print(f"An error occurred: {e}")

def get_selected_column_data_for_row(input_file, input_sheet_name, search_value, search_column):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df.loc[df[search_column] == search_value]
    except Exception as e:
        print(f"An error occurred: {e}")
        return -1

def get_selected_column_data_for_row_from_df(df, search_value, search_column):
    try:
        return df.loc[df[search_column] == search_value]
    except Exception as e:
        print(f"An error occurred: {e}")
        return -1


execution_directory = "C:\\Users\\mahbu\\Desktop\\Practices"
i_file = execution_directory + "\\Profiling Master- QC.xlsx"
i_s_name = 'Main Data'
o_file = execution_directory + "\\OutputFile2.xlsx"
o_s_name = 'Scrapped_Sheet'
#r_num = 1 # here considering first row will have column names, therefore that'd be -1, next
#c_indices = [0, 1, 3, 5, 6, 7]

#create_new_sheet(i_file,o_file,r_num,c_indices)

#r_indices = [0,1,4, 15]
#create_new_sheets(i_file,o_file,r_indices,c_indices)
# s_value = 'Bunkers'
# s_column = 'Name'
#
# r_index = get_row_index(i_file, i_s_name, s_value, s_column)
# if r_index != -1:
#     print(f"Row index of '{s_value}' in '{s_column}': {r_index}")
# else:
#     print(f"'{s_value}' not found in '{s_column}'.")


#c_name = 'Id'

# c_index = get_column_index(i_file, i_s_name, c_name)
# if c_index != -1:
#     print(f"Column index of '{c_name}': {c_index} in the Worksheet: {i_s_name}.")
# else:
#     print(f"'{c_name}' not found in the Worksheet: {i_s_name}.")
#
# c_names = ['Id', 'Agent Name', 'Month', 'Active Listening', 'Verbal Excellence', 'Courteous Approach', 'Identification and Action for Resolution', 'Correct & Complete Information For Resolution (CCIR)', 'Avoid Rude/Unprofessional Behavior/Approach (ARU)','Ownership & Proctiveness (OP)']
# r_values = ['Shusmoy Kundu']
# c_indices = []
# r_indices = []
#
# for c_name in c_names:
#     c_indices.append(get_column_index(i_file,i_s_name,c_name))
#
# for r_value in r_values:
#     r_indices.append(get_row_index(i_file,i_s_name,r_value,c_names[1]))

# print(c_indices)
# print('\n')
# print(r_indices)

#create_new_sheets(i_file, i_s_name, o_file,o_s_name, r_indices, c_indices)

# s_value = 'Bunkers'
# s_column = 'Name'
# d_rows = find_duplicate_rows_in_column(i_file,i_s_name,s_column)
# print(d_rows)

# d = {'col1':[1,2], 'col2':[3,4], 'col3':[5,6]}
# ddf = pd.DataFrame(data=d)
# print(ddf)

#t_df = create_data_frame(input_file=i_file, input_sheet_name=i_s_name, row_indices=r_indices, column_indices=c_indices)
# t_df = get_selected_column_data_for_row(input_file=i_file, input_sheet_name=i_s_name, search_value='Shusmoy Kundu', search_column='Agent Name')
# t_df = get_selected_column_data_for_row_from_df(df=t_df,search_column='Month', search_value='Jan')
# print(t_df)
# write_to_output_file(df=t_df, output_file=o_file, output_sheet_name="NormalData")
# t_df1 = create_transposed_frame_from_df(t_df)
# write_to_output_file(df=t_df1, output_file=o_file, output_sheet_name="TransposedData")

c_nam = ['Month', 'Active Listening', 'Verbal Excellence', 'Courteous Approach', 'Identification and Action for Resolution', 'Correct & Complete Information For Resolution (CCIR)', 'Avoid Rude/Unprofessional Behavior/Approach (ARU)','Ownership & Proctiveness (OP)']

def return_selected_columns_of_selected_row(input_file, input_sheet_name ,filter_column,filter_value, c_names):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[filter_column]==filter_value][c_names]
    except Exception as e:
        print(f"An error occurred {e}")
        return -1

t_dfff = return_selected_columns_of_selected_row(input_file=i_file,input_sheet_name=i_s_name,filter_column='Agent Name',filter_value='Shusmoy Kundu',c_names=c_nam)
print(t_dfff)
