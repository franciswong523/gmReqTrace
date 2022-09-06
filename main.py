# This is created by Francis to generate GM Requirement Traceability Report.
import pandas as pd
from openpyxl import load_workbook
import glob
import os.path
import shutil
import xlsxwriter

src_folder_path = r'\\fr05245vma\Reports\GM\Traceability_reports'
des_folder_path = r'C:\Francis\Conti\GM\Reports\Requirement_Allocation'
file_type = r'\*xlsx'

def FilterReqTraceReport(file_name, path_len, des_directory):

    # Create a Panda dataframe by reading in a particular sheet
    data = pd.read_excel(file_name, sheet_name='SyRD Full Traceability')

    # Create new file from template
    new_file = f'{des_directory}\\FILTER_{file_name[int(path_len) + 1:]}'
    shutil.copyfile(f'{des_directory}\\template.xlsx', new_file)
    print(f'Preparing the file - {new_file}')

    # New method:
    # append a new sheet to a template file
    #book = load_workbook(new_file)
    #writer = pd.ExcelWriter(new_file, engine='openpyxl')
    #writer.book = book
    #writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    # Original method:
    # create a new file or overwrite it
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(new_file, engine='xlsxwriter')

    # Remove all the duplicate items
    data.drop_duplicates(subset=['SyRD ID', 'SyRD Responsible'], keep='first', inplace=True)

    # remove rows base on conditions
    data.drop(data[data['SyRD State'] != 'Released'].index, inplace=True)
    data.drop(data[data['ProdApp - GM Gen 12'] != 'Accepted'].index, inplace=True)

    # Convert the dataframe to an XlsxWriter Excel object. We also turn off the
    # index column at the left of the output dataframe.
    data.to_excel(writer, sheet_name='RAW Data', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()



# Find the latest file in the folder \\fr05245vma\Reports\GM\Traceability_reports\
files = glob.glob(src_folder_path + file_type)
proces_this_file = max(files, key=os.path.getctime)
print('Processing the file - ' + proces_this_file)

# start filtering the trace report
FilterReqTraceReport(proces_this_file, len(src_folder_path), des_folder_path)




