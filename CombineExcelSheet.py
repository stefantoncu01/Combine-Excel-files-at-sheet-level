# import pandas as pd
# import glob
#
# path = 'C:\\Users\\Desktop\\ExcelPython\\*.xls' #choose data only from files with extension xls from a specific folder / location
# excel_files = glob.glob(location)
#
# df1 = pd.DataFrame()
#
# for excel_file in excel_files:
#     df2 = pd.read_excel(excel_file)
#     df1 = pd.concat([df1, df2], ignore_index=True)
#
# df1.to_excel("C:\\Users\\Desktop\\ExcelPythonResult\\varianta_finala.xls", index = False) #the path where you want the combined file to be created

# This part is just your own code, I've added it here because you
# couldn't figure out where `excel_files` came from
#################################################################

import os
import pandas as pd
#Combines data from multiple Excel files at sheet level

import os
import pandas as pd
from collections import defaultdict

EXCEL_FILES_LOCATION = r'C:\\Users\\Desktop\\ExcelPython\\'  # modify only between '' with the location of the excel files
OUTPUT_LOCATION = r'C:\\Users\\Desktop\\ExcelPythonResult\\combined.xls' # modify only between '' with the location of the output folder

path = os.chdir(EXCEL_FILES_LOCATION)
files = os.listdir(path)

# pull files with '.xls' or `.xlsx` extension
excel_files = [file for file in files if '.xls' in file]


worksheet_lists = defaultdict(list)
for file_name in excel_files:
    workbook = pd.ExcelFile(file_name)
    for sheet_name in workbook.sheet_names:
        worksheet = workbook.parse(sheet_name)
        # worksheet['source'] = file_name
        worksheet_lists[sheet_name].append(worksheet)

worksheets = {
    sheet_name: pd.concat(sheet_list)
    for (sheet_name, sheet_list) in worksheet_lists.items()
}
writer = pd.ExcelWriter(OUTPUT_LOCATION)

for sheet_name, df in worksheets.items():
    df.to_excel(writer, sheet_name=sheet_name, index=False)
writer.save()
