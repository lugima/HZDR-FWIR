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


# Function that converts all data from SuperSIMS .mpa files to Excel.
# Creates one workbook per .mpa file in the selected folder.
# One worksheet per datablock.
def mpa2sheets(file_dir, workbook):
    
    # Reads file
    with open(file_dir, 'r') as file:

        # Split file in lines
        lines = file.readlines()


    # Create first worksheet for parameters
    worksheet = workbook.add_worksheet('Parameters')

    row = 0
    for line in lines:

        # Removes lineskips from lines
        line = line.replace('\n', '')

        if '=' in line:     # Writes 'A=B' format lines                                 
            in_data = False

            splitline = line.split('=')
            worksheet.write(row, 0, splitline[0])
            worksheet.write(row, 1, new_float(splitline[1]))

        elif line == '[DATA]':  # Start writing 'A    B' format lines
            in_data = True
            worksheet.write(row, 0, line)

        elif line.startswith('[DATA') or line.startswith('[CDAT'): # Creates new worksheet for new datablock
            in_data = False
            # The name of each worksheet is the NAME of the corresponding graph
            name_index = lines.index(line + '\n') + 9          # The NAME=ABC line is 9 lines below [DATAX, 1234]
                                                               # We have to readd the lineskip we removed before to find the element in the lines list

            ws_name = lines[name_index].split('=')[1]
            worksheet = workbook.add_worksheet(ws_name)
            row = 0

            worksheet.write(row, 0, line)

        elif line.startswith('[') or line == '':   # Writes any other format lines
            in_data = False
            worksheet.write(row, 0, line)

        elif in_data == True:   # Writes 'A    B' format lines
            splitline = line.split('\t')
            for i in range(len(splitline)):
                worksheet.write(row, i, new_float(splitline[i]))

        row += 1
    
    return
        

        
        
'------MAIN------'

title = '//////////////////////////////////////////////////////////////\n Collect data from .mpa files in Microsoft Excel sheets\n L. Gil Martín. 2022\n////////////////////////////////////////////////////////////// \n'
print(title)

print('Please enter the name of the directory (absolute path of folder): ')
dir = new_input()

folder_name = dir.split('\\')[-1] # Gets the folder name from the directory

# Gathers all files with .mpa extension in a list
# Creates a workbook in a list for every .mpa file
files_list = []
workbook_list = []

for file in os.listdir(dir):
    if file.endswith('.mpa'):
        files_list.append(file)

        filename = file.split('.')[0] # Gets filename without extension

        workbook_list.append(xlsx.Workbook(filename + '.xlsx')) 

dirs_list = [dir + '\\' + fname for fname in files_list]

# Reads data from .mpa files and sorts it in worksheets
for i in range(len(dirs_list)):
    mpa2sheets(dirs_list[i], workbook_list[i])
    workbook_list[i].close()

print("The files have been created successfully in the location of the executable.")
print("Press Enter to close this window.")
foo = new_input()