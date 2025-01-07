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

def get_row_index(input_file, input_sheet_name, search_value, search_column):
    try:
        df = pd.read_excel(input_file, sheet_name=input_sheet_name)
        return df[df[search_column] == search_value].index[0]
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



execution_directory = "C:\\Users\\NAJIB\\Desktop\\Practices"
i_file = execution_directory + "\\Excel\\SourceFile2.xlsx"
i_s_name = 'Sheet1'
o_file = execution_directory + "\\Excel\\OutputFile2.xlsx"
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
# c_names = ['Id', 'Name', 'Email', 'CompanyCategoryId', 'ShortName']
# r_values = ['BFS Chicken', 'MZ Resort', 'Northern Bank Limited', 'Indian Spicy King', 'The Food Factory']
# c_indices = []
# r_indices = []
#
# for c_name in c_names:
#     c_indices.append(get_column_index(i_file,i_s_name,c_name))
#
# for r_value in r_values:
#     r_indices.append(get_row_index(i_file,i_s_name,r_value,c_names[1]))
#
# print(c_indices)
# print('\n')
# print(r_indices)

#create_new_sheets(i_file, i_s_name, o_file,o_s_name, r_indices, c_indices)

s_value = 'Bunkers'
s_column = 'Name'
d_rows = find_duplicate_rows_in_column(i_file,i_s_name,s_column)
print(d_rows)