import pandas as pd
from pathlib import Path
import copy

# call up the file's path. assign the file to a variable
input_file = Path.cwd() / 'StressLoads Rev1.xlsx'

extra_file = Path.cwd() / 'extraStressLoads Rev1.xlsx'

# ceate report file's path for storage. assign the file to a variable
output_file = Path.cwd() / 'StressLoadsOutput Rev1.xlsx'

# create a writer object first using Pandas' ExcelWriter. use xlswriter engine for more sophisticated customizations
writer = pd.ExcelWriter(output_file, 
                        engine='xlsxwriter',
                        options={'strings_to_numbers': True})

# read the two files into Python's Panda dataframes. assign file to variables
load_data = pd.read_excel(input_file, 
                        engine='openpyxl', 
                        sheet_name='Table 1',
                        header=0, 
                        usecols='A:AL'                        
                        )

extra_data = pd.read_excel(extra_file, 
                        engine='openpyxl', 
                        sheet_name='Table 1',
                        header=0,
                        usecols='A:AC'
                        )

# 'extraStressLoads.xlsx' data columns not the same as 'StressLoads.xlsx' data columns
# insert extra columns at certain indices to make the two files equal. fill those cells with "No Data".
# this is necessary in order to concatenate the two files later
for _ in range(3):
    extra_data.insert(_, str(_ + 1), "No Data")

for _ in range(3):
    extra_data.insert((_ + 4), str(_ + 11), "No Data")

for _ in range(3):
    extra_data.insert((_ + 34), str(_ + 31), "No Data")

# columns names of 'StressLoads.xlsx' data come in as "Unnamed". fix them to be empty
load_data.rename(columns=lambda x: x.replace(': ', '_'), inplace=True)
load_data.columns = [column_names.replace('_', '') for column_names in load_data.columns]
load_data.columns = load_data.columns.str.replace('Unnamed.*[0-9999]$', '')

# columns names of 'extraStressLoads.xlsx' data come in as "Unnamed". fix them to be empty
extra_data.rename(columns=lambda x: x.replace(': ', '_'), inplace=True)
extra_data.columns = [column_names.replace('_', '') for column_names in extra_data.columns]
extra_data.columns = extra_data.columns.str.replace('Unnamed.*[0-9999]$', '')
extra_data.columns = extra_data.columns.str.replace('\d', '')

# move a column name to a row cell location below it
extra_data.iloc[0, 3] = extra_data.columns[3]

# let the column names of 'extraStressLoads.xlsx' be the same as that of 'StressLoads.xlsx'
extra_data.columns = load_data.columns

# now that the column names, numbers and indices are the same, concatenate the two sets of data
load_data = pd.concat([load_data, extra_data], join="outer")

# data is mutable. safe to make a copy so that earlier created referencing are avoided
load_data = copy.deepcopy((load_data.reset_index(drop=True)).fillna(0))

# the column name of 'extraStressLoads.xlsx' comes in as another line in the merged data.
# find the index and delete that row from the complete data
index_of_row_to_delete = load_data.loc[load_data.isin(['Node Number']).any(axis=1)].index.tolist()[1]
load_data = load_data.drop([index_of_row_to_delete])

# build up a new Pandas DataFrame with the desired output format using Lists
condition_names = [load_data.columns[_]
                    for _ in range(len(load_data.columns) - 5)
                    if load_data.columns[_] != ''
                    ]

header_names = [_ for _ in load_data.iloc[0, 0:7]]
header_names.append(load_data.columns[-1])
header_names.append('Condition')

for _ in  load_data.iloc[0, [8,7,9]]:
    header_names.append(_)

for _ in load_data.iloc[0, [35,34,36]]:
    header_names.append(_)

header_names_clean = [" ".join(_.split('\n')) for _ in header_names]

# data is mutable. safe to make a copy so that earlier created referencing are avoided
new_headers = header_names_clean.copy()

# Lists created above now used to make a new Pandas DataFrame
output_data = pd.DataFrame(columns=new_headers)

# Lists created above now used to make a new Pandas DataFrame
output_data_conditions = pd.DataFrame([condition_names[i] for i in range(len(condition_names))], columns=['Condition'])

# merging the two DataFrames created above. this DataFrame is now the desired output template to be filled
output_data = pd.merge(output_data, 
                        output_data_conditions, 
                        how='outer'
                        )

# the template's column indices are all over the place so re-index them and fill the cell data with zeros
# rename some of them as well, if needed
output_data = output_data.reindex(columns=['Support Name',
                                    'Support Status',
                                    'Diameter (mm)',
                                    'Node Number',
                                    'Location',
                                    'Pipe Name',
                                    'Pipe Status',
                                    'Date of Caesar Load Data',
                                    'East (mm)',
                                    'North (mm)',
                                    'Elevation (mm)',
                                    'Condition',
                                    'E‐W (kN)',
                                    'N‐S (kN)',
                                    'Vert (kN)']
                                    ).fillna(0)

output_data_column_names = (output_data.columns).values.tolist()

first_batch_columns = ['Support Name', 'Support Status', 'Diameter (mm)', 'Node Number', 'Location', 'Pipe Name', 'Pipe Status']

# data is mutable. safe to make a copy so that earlier created referencing are avoided
sub_output_data = copy.deepcopy(output_data)

# function takes one argument and uses it to fill up nine rows of the same data into the columns listed above
# of first_batch_columns. data goes into specific areas based on certain criterion. function will be called later
# the nine Condition names are used as a counter for the iteration
# the returned data is a collection of values populated into a sub_output_data DataFrame
def nine_block_fill(cell_value):
    column_counter = row_cell_values.index(cell_value)
    if column_counter < len(first_batch_columns):
        for _ in range(len(condition_names)):                
            sub_output_data.loc[_, output_data_column_names[column_counter]] = cell_value
    
    elif row_cell_values.index(cell_value) == 37:
        for _ in range(len(condition_names)):               
            sub_output_data.loc[_, 'Date of Caesar Load Data'] = cell_value
    
    elif row_cell_values.index(cell_value) == 35:
        for _ in range(len(condition_names)):               
            sub_output_data.loc[_, 'East (mm)'] = cell_value

    elif row_cell_values.index(cell_value) == 34:
        for _ in range(len(condition_names)):               
            sub_output_data.loc[_, 'North (mm)'] = cell_value
    
    elif row_cell_values.index(cell_value) == 36:
        for _ in range(len(condition_names)):               
            sub_output_data.loc[_, 'Elevation (mm)'] = cell_value
    
    return sub_output_data

# row indices are all over the place so re-index them
load_data = load_data.reset_index(drop=True)

# make a List collection of data per row and iterate through them
# call the Function above to populate the cells of sub_output_data DataFrame
# each iteration then merges the sub_output_data together with the previous one building a block on top of each other
# then the sub_output_data is concatenated with the out_data DataFrame
# pd.merge is used to merge the blocks of new columns together
# pd.concat is used to merge the block of new rows together
for _ in range(1, len(load_data.iloc[:, 0])):
    row_cell_values = load_data.loc[_, :].values.tolist()
    for _ in row_cell_values:
        sub_output_data = pd.merge(sub_output_data, nine_block_fill(_))
    output_data = (pd.concat([output_data, sub_output_data]))

# row indices are all over the place so re-index them
output_data.reset_index(drop=True)

# remove the first block of zero data set at the top of the DataFrame (first 9 rows)
output_data = output_data.iloc[9:]

# row indices are all over the place so re-index them
output_data.reset_index(drop=True)

# data is mutable. safe to make a copy so that earlier created referencing are avoided
final_output = copy.deepcopy((output_data.reset_index(drop=True)).fillna(0))
 
# now populate the forces for each condition on each row of the final_output DataFrame created above
# different counters required for the different iterations
# use the length of data present in the final_output DataFrame
# a Dictionary of data is created for each Condition for each row traversed
number_of_data_counter = 0
condition_names_counter = 0
for i_index in range(1, len(load_data.iloc[:, 0])):
    
    dict_of_forces = {
                final_output.loc[0,'Condition']: load_data.iloc[i_index, [8,7,9]].values.tolist(),
                final_output.loc[1,'Condition']: load_data.iloc[i_index, [11,10,12]].values.tolist(),
                final_output.loc[2,'Condition']: load_data.iloc[i_index, [14,13,15]].values.tolist(),
                final_output.loc[3,'Condition']: load_data.iloc[i_index, [17,16,18]].values.tolist(),
                final_output.loc[4,'Condition']: load_data.iloc[i_index, [20,19,21]].values.tolist(),
                final_output.loc[5,'Condition']: load_data.iloc[i_index, [23,22,24]].values.tolist(),
                final_output.loc[6,'Condition']: load_data.iloc[i_index, [26,25,27]].values.tolist(),
                final_output.loc[7,'Condition']: load_data.iloc[i_index, [29,28,30]].values.tolist(),
                final_output.loc[8,'Condition']: load_data.iloc[i_index, [32,31,33]].values.tolist()
                }

    forces_list_index = 0
    for _ in range(9):
        final_output.loc[(condition_names_counter), 'E‐W (kN)'] = dict_of_forces[condition_names[forces_list_index]][0]
        final_output.loc[(condition_names_counter), 'N‐S (kN)'] = dict_of_forces[condition_names[forces_list_index]][1]
        final_output.loc[(condition_names_counter), 'Vert (kN)'] = dict_of_forces[condition_names[forces_list_index]][2]
        forces_list_index += 1
        condition_names_counter += 1
        if condition_names_counter < len(condition_names):
            continue

    number_of_data_counter += 1
    if number_of_data_counter < len(load_data.iloc[:, 0]):
        continue
            
final_output = final_output.fillna(0)

# function takes one argument
# negative values in the final_output are objects at this stage
# convert them to float numbers by spliting, rejoining and negating them
def cconvert_to_numbers(value):
    if type(value) == int:
        return value
    if type(value) != float:
        value_into_list = list(value)
    else:
        return value
    if value_into_list[0] != '‐':
        return float(''.join(value_into_list))

    del value_into_list[0]
    return float(''.join(value_into_list)) * -1

# apply the function above to all the columns with numbers
final_output['East (mm)'] = final_output['East (mm)'].apply(cconvert_to_numbers)
final_output['North (mm)'] = final_output['North (mm)'].apply(cconvert_to_numbers)
final_output['Elevation (mm)'] = final_output['Elevation (mm)'].apply(cconvert_to_numbers)
final_output['E‐W (kN)'] = final_output['E‐W (kN)'].apply(cconvert_to_numbers) / 1000
final_output['N‐S (kN)'] = final_output['N‐S (kN)'].apply(cconvert_to_numbers) / 1000
final_output['Vert (kN)'] = final_output['Vert (kN)'].apply(cconvert_to_numbers) / 1000

# output format for certain columns variable created to be used later
number_formats = {
        'East (mm)': "{:.3f}",
        'North (mm)': "{:.3f}",
        'Elevation (mm)': "{:.3f}",
        'E‐W (kN)': "{:.3f}",
        'N‐S (kN)': "{:.3f}",
        'Vert (kN)': "{:.3f}"
        }

# apply the format the columns
for columns, item in number_formats.items():
    final_output[columns] = final_output[columns].map(lambda x: item.format(float(x)))

# some data has '=' infront of them. delete the '=' 
final_output.loc[:, 'Support Name'] = final_output.loc[:, 'Support Name'].str.replace("=", "")

# save this summary data to a named sheet
final_output.to_excel(writer, 
                        sheet_name='Extracted Useable Data', 
                        index=False
                        )
# use the writer object to pull out the workbook. assign to a variable
workbook = writer.book 

# access this named worksheet within the workbook. assign to a variable
worksheet_1 = writer.sheets['Extracted Useable Data'] 

# variables created for each Excel format required
# for the entire workbook variable, set up the preferred number format. assign to variables
text_data = workbook.add_format({'num_format': '@'})
other_entries = workbook.add_format({'align': 'right'})
coordinate_values = workbook.add_format({'num_format': '0.000', 'align': 'right'}) 
force_values = workbook.add_format({'num_format': '0.000', 'align': 'right'}) 

# set the width and style for these columns in this worksheet using the format variables created above
worksheet_1.set_column('A:A', 15, text_data) 
worksheet_1.set_column('B:B', 15, other_entries) 
worksheet_1.set_column('C:C', 14, other_entries) 
worksheet_1.set_column('D:E', 15, other_entries)
worksheet_1.set_column('F:F', 21, other_entries)
worksheet_1.set_column('G:G', 10, other_entries)
worksheet_1.set_column('H:H', 22, other_entries)
worksheet_1.set_column('I:K', 15, coordinate_values) 
worksheet_1.set_column('L:L', 12, other_entries)
worksheet_1.set_column('M:O', 10, force_values)

# save and close the writer object after all the creations 
# and after the writer object is completed
writer.save() 




######################################################### WORKING TEST BENCH AREA #############################################################################
