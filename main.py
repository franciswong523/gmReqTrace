# This is created by Francis to generate GM Requirement Traceability Report.
import pandas as pd
from openpyxl import load_workbook
import glob
import os.path
import shutil
import xlsxwriter

src_folder_path = r'\\cw01.contiwan.com\SMT\didr2537\Reports\GM\Traceability_reports'
des_folder_path = r'C:\Francis\Conti\GM\Reports\Requirement_Allocation'
file_type = r'\G*xlsx'

def FilterReqTraceReport(file_name, path_len, des_directory):

    # Create new file from template
    new_file_name = f'{des_directory}\\FILTER_{file_name[int(path_len) + 1:]}'
    shutil.copyfile(f'{des_directory}\\template.xlsx', new_file_name)
    print(f'Preparing the file - {new_file_name}')

    # Create a Panda dataframe by reading in a particular sheet
    data_read_in = pd.read_excel(file_name, sheet_name='SyRD Full Traceability', engine='openpyxl')

    # Remove all the duplicate items
    data_read_in.drop_duplicates(subset=['SyRD ID', 'SyRD Responsible'], keep='first', inplace=True)

    # remove rows base on conditions
    data_read_in.drop(data_read_in[data_read_in['SyRD State'] != 'Released'].index, inplace=True)
    data_read_in.drop(data_read_in[data_read_in['ProdApp - GM Gen 12'] != 'Accepted'].index, inplace=True)

    data_read_in.reset_index(drop=True)
#   data_rows = len(data_read_in.index)

    # Write the filtered data to a new xls file
    # Is there a way to trigger data resync?
    with pd.ExcelWriter(new_file_name, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        data_read_in.to_excel(writer, sheet_name="RAW Data", index=False)

# Find the latest file in the folder \\fr05245vma\Reports\GM\Traceability_reports\
files = glob.glob(src_folder_path + file_type)
proces_this_file = max(files, key=os.path.getctime)
print('Processing the file - ' + proces_this_file)

# start filtering the trace report
FilterReqTraceReport(proces_this_file, len(src_folder_path), des_folder_path)




