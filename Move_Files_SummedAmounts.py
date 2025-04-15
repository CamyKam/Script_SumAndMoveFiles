import os
import glob
import shutil
from datetime import datetime
import pandas as pd
import pandas.io.formats.excel

#Will need to make sure xlsxwriter is installed in python project

source_file_path = "C:/Users/camyk/PycharmProjects/InitialFolder"
destination_file_path = "C:/Users/camyk/PycharmProjects/MovedFilesFolder"

#Looking only at the files that start with the file name 'marriage'
files = glob.glob(os.path.join(source_file_path,'marriage*.xlsx' ), recursive=True)

for f in files:
    try:
        df = pd.read_excel(f)
        print(f)
    except IOError as error:
        print(error)

    summed_df = df.groupby(['Year'])['Total'].sum()   #sum the totals column by year"

    substring = f[45:65]

    date_time= datetime.now()
    datetime_file_name = f"New_{str(substring)}_{str(date_time.month)}-{str(date_time.day)}-{str(date_time.hour)}-{str(date_time.minute)}-{str(date_time.second)}-{str(date_time.microsecond)}.xlsx"

    new_file_path_loc = f"{source_file_path}/{datetime_file_name}"
    try:
        writer = pd.ExcelWriter(new_file_path_loc, engine='xlsxwriter')
        pd.io.formats.excel.ExcelFormatter.header_style = None    #this to remove the automatic borders that pandas has for headers
        summed_df.to_excel(writer)
        writer.close()
    except Exception as ex:
        print("An error occurred during writing: ", ex)

    try:
        shutil.move(new_file_path_loc, destination_file_path)
        print(f"Moved {datetime_file_name} to {destination_file_path}")
    except Exception as ex:
        print("An error occurred: ", ex)
