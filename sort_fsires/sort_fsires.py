'//////////////////////////////////////////////////////////////'
# Collect data from .fsires files in Microsoft Excel sheets
# L. Gil Martín. 2022
'//////////////////////////////////////////////////////////////'

import os
import sys
import xlsxwriter as xlsx


'------FUNCTIONS------'

# Function that removes "\n" from sys.stdin.readline() inputs.
def remove_lineskip(command):
    return command[:len(command)-1]


# Redefined input function
def new_input():
    input = sys.stdin.readline()
    return remove_lineskip(input)

# Function that converts string to float only if it is possible.
def new_float(element):
    try:
        float(element)
        return float(element)
    except ValueError:
        return element

# Function that converts all data from AMS .fsires files to Excel.
# Creates one worksheet per .fsires file in the selected folder.
def fsires2sheet(file_dir, file_num, workbook):

    # Creates new worksheet
    worksheet = workbook.add_worksheet(str(file_num))
    
    # Reads file
    with open(file_dir, 'r') as file:

        # Split file in lines
        lines = file.readlines()

    # Removes lineskip from every element in the array
    for i in range(len(lines)):
        lines[i] = lines[i].replace('\n', '')

    # Splits the previous array in 3 sections
    RESULTS_index = lines.index('[RESULTS]')
    DATA_index = lines.index('[BLOCK DATA]')

    specs_raw = lines[1 : RESULTS_index] # Skips the first line
    results_raw = lines[RESULTS_index : DATA_index]
    data_raw = lines[DATA_index : ]

    # Splits elements in columns
    specs = [item.split(' : ') for item in specs_raw]         # [['label1', '1'], ['label2', '2'], ...]

    results = [item.split(' : ') for item in results_raw]

    data = [item.split() for item in data_raw]

    # Writes specs section in worksheet
    row = 0

    for elem in specs:
        if len(elem) == 1:
            worksheet.write(row, 0, new_float(elem[0]))
        else:
            worksheet.write(row, 0, new_float(elem[0]))
            worksheet.write(row, 1, new_float(elem[1]))
        row += 1

    # Writes results section in worksheet
    for elem in results:
        if len(elem) == 1:
            worksheet.write(row, 0, new_float(elem[0]))
        else:
            worksheet.write(row, 0, new_float(elem[0]))
            worksheet.write(row, 1, new_float(elem[1]))
        row += 1

    # Writes data section in worksheet
    for elem in data:
        for i in range(len(elem)):
            worksheet.write(row, i, new_float(elem[i]))
        row +=1
    
    return

# Function that converts data block from AMS .fsires files to Excel.
# Creates one worksheet per sample in the files in the selected folder.
def fsiresdata2sheets(dirs_list, workbook):

    id_list = []
    worksheets_list = []
    row_counter = []
    count_arr = []

    for dir in dirs_list:

        filename = dir.split('\\')[-1] # Gets the name of the current .fsires file

        # Reads file
        with open(dir, 'r') as file:
            # Split file in lines
            lines = file.readlines()

        # Removes lineskip from every element in the array
        for i in range(len(lines)):
            lines[i] = lines[i].replace('\n', '')

        # Saves the data section
        DATA_index = lines.index('[BLOCK DATA]')

        data_raw = lines[DATA_index + 1: ]

        # Identifies the Sample Id and opens its worksheet
        id = lines[3].split(' : ')[1]

        if id not in id_list:
            id_list.append(id)
            id_index = id_list.index(id)

            worksheet = workbook.add_worksheet(id)
            worksheets_list.append(worksheet)

            row_counter.append(0)
            count_arr.append(0)
        else:
            id_index = id_list.index(id)
            worksheet = worksheets_list[id_index]
            data_raw = data_raw[1 : ]       # Only write the headers the first time a batch is read          

        # Splits elements in columns
        data = [item.split() for item in data_raw]

        # Writes data section in worksheet following the previous set     

        for elem in data:
            for i in range(len(elem)):
                worksheet.write(row_counter[id_index], 0, filename)               # Write filename column
                worksheet.write(row_counter[id_index], 1, count_arr[id_index])    # Write counter
                worksheet.write(row_counter[id_index], i+2, new_float(elem[i]))     # Write data

            row_counter[id_index] += 1
            count_arr[id_index] += 1
        
        worksheet.write(0, 0, "File of origin")      # Write filename column header


    return
        
        
'------MAIN------'

title = '//////////////////////////////////////////////////////////////\n Collect data from .fsires files in Microsoft Excel sheets\n L. Gil Martín. 2022\n////////////////////////////////////////////////////////////// \n'
print(title)

print('Please enter the name of the directory (absolute path of folder): ')
dir = new_input()

folder_name = dir.split('\\')[-1] # Gets the folder name from the directory

# Gathers all files with .fsires extension in a list
files_list = []

for file in os.listdir(dir):
    if file.endswith('.fsires'):
        files_list.append(file)

files_sorted = sorted(files_list, key=lambda x: int(x.split('.')[0]))
dirs_sorted = [dir + '\\' + fname for fname in files_sorted]

# Creates an Excel file
workbook = xlsx.Workbook(folder_name + '_' + 'data_sorted_by_file.xlsx')

# Reads data from .fsires files and sorts it in worksheets
for i in range(len(files_list)):
    fsires2sheet(dirs_sorted[i], i+1, workbook)

workbook.close()

# Creates another Excel file
workbook2 = xlsx.Workbook(folder_name  + '_' + 'data_sorted_by_batch.xlsx')

# Reads data and sorts it by batch code
fsiresdata2sheets(dirs_sorted, workbook2)

workbook2.close()

print("The files have been created successfully in the location of the executable.")
print("Press Enter to close this window.")
foo = new_input()