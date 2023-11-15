# Import packages

import pandas as pd
import numpy as np
import subprocess
from zipfile import ZipFile
from pathlib import Path
import shutil
import glob
import os
from datetime import datetime
from credentials import pword_bytes, pword
import re
import getpass
from tkinter import messagebox
import time



class Main:

    def __init__(self, days=""):
        self.days = days
        # To be assigned today's date in desired format
        self.today = datetime.now()
        # To contain formatted dates
        self.date_range = []
        # To contain formatted dates as a string
        self.date_range_str = []

        # No input will default to 1
        if self.days == "":
            self.days = int(1)
        else:
            self.days = int(self.days)

    def main_processor(self):

        # Sets daterange in format required for file names

        self.date_processing()

        # Initialise File_manipulation class instance
        
        FM = File_manipulation()

        # Delete previous files in the extract_files folder

        FM.delete_old_files()

        # Copy the files by name (dynamically using the date format setup above)

        FM.copy_files(self.date_range, self.date_range_str)

        # Unlock and extract the contents of the zip file

        FM.unzip_files()
        
        # Initialise Data manipulation class

        DM = Data_manipulation(self.date_range)

        # Union the tables to make master df

        self.master_df = DM.union()

        # Print output to console

        print(self.master_df)

        # Export the master_df

        self.master_df.to_csv(r"inputs\extract_files\Daily_customers_union.csv", index=False)

        # Create OneDrive path

        OneDrive_path = r"C:\Users\{}\OneDrive\folder\Daily_customers_union.csv".format(getpass.getuser())

        # Message Box - protect from duplication

        check_mod_time(OneDrive_path)

        # Delete the old file once checked
        '''if os.path.isfile(OneDrive_path):
            os.remove(OneDrive_path)''' # To remove if it runs successfully

        # Export to one drive / sharepoint

        self.master_df.to_csv(OneDrive_path, index=False)

        print("Upload completed")

        # Launch excel

        os.system('start EXCEL.EXE "inputs\extract_files\Daily_customers_union.csv"')



    def date_processing(self):
        for i in range(self.days):
            # Create an array of dates formatted Ymd
            # In date format
            self.date_range.append(self.today - (pd.Timedelta(f'{i} days')))
            # In string format
            self.date_range_str.append(self.date_range[i].strftime("%Y%m%d"))



class File_manipulation:
    
    def __init__(self):
        # To avoid repeated instances of the folder location

        self.in_folder = "inputs\\"
        self.archive_folder = "C:\\Users\\{}\\OneDrive\\folder_2\\2023\\Daily Reporting\\ customers With Active Tickets\\Output".format(getpass.getuser())
        self.copied_filepaths = []

    def delete_old_files(self):
        # Look in "extract_files\", remove all files
        try:
            files = glob.glob('{}extract_files\\*'.format(self.in_folder))
            for file in files:
                os.remove(file)
        except BaseException:
            raise BaseException("An error occured when trying to delete files from \\inputs\\extract_files")

    def copy_files(self, date_range, date_range_str):
        # Copy the zip file over to local folders
        '''Copy all the original files to protect originals from corruption / error
        Program processes these copies only'''
        self.month_folders = []
        try:
            for i in range(len(date_range)):
                # get all file locs for every file we want to copy
                copy_loc = f"\\\edi-fileserver\\DataServices\\Private\\SSIS_Automations\\Outputs_Zipped\\001__With_Active\\Data_{date_range_str[i]}*.zip"
                # get month folder names for each file loc
                self.month_folders.append("{}.{}\\".format(date_range[i].month, date_range[i].strftime("%B")))

                for file in glob.glob(copy_loc):
                    # Remove each file if they already exist
                    try:
                        # below formatted string simplified format: input\\4.April\\Data_04042023
                        os.remove("{}{}Data_{}*.Zip".format(self.in_folder, self.month_folders[i], date_range_str[i]))
                    except OSError:
                        pass
                    # Proceed to copy the files from source destination once removing existing files.
                    print(file)
                    file_path = "\\{}Data_{}.Zip".format(self.month_folders[i], date_range_str[i])
                    shutil.copy(file, f"{self.in_folder}{file_path}")
                    shutil.copy(file, f"{self.archive_folder}{file_path}")
                    self.copied_filepaths.append(file_path)

                    
            print("File copies made...")
        except PermissionError:
            file_permissions_error = '''***This application does not have permission to read the original zip files ('Data_yyyymmdd.zip').
            Please check you can access the files manually.'''
            raise PermissionError(f"{file_permissions_error}")


    def unzip_files(self):
        # Unzip the files once their copied over.
        try:
            for file in self.copied_filepaths:
                    print(file)
                    file = self.rename_file(file, "Zip")
                    print(file)
                    # Extract all to the "extract_files" folder
                    with ZipFile(f"{self.in_folder}{file}", "r") as zippy:
                        zippy.extractall(path = "{}extract_files\\".format(self.in_folder), pwd = pword_bytes)

            for file in os.listdir("{}extract_files\\".format(self.in_folder)):
                # Once the files are unzipped, give them new names to tidy them up
                if file.endswith(r".zip"):
                    self.rename_file(file, "Zip")
                
            print("File successfully unzipped")
            
        except BaseException:
            raise BaseException("Unable to unzip zip file...")



    """def csv_to_xl(self):
        # Depreciated function but kept in as it may have some use in the future
        try:
            for file in os.listdir(self.in_folder):
                if file.endswith(".csv"):
                    df = pd.read_csv(f"{self.in_folder}{file}")
                    file = file.replace(".csv", ".xlsx")
                    df.to_excel(f"{self.in_folder}{file}", engine="openpyxl")
                    print("Excel file created...")
        except BaseException:
            raise BaseException("An error has occured converting the CSV files to XLSX")"""

    def rename_file(self, file, extension):
        new_name = file
        new_name = re.sub(f"_[0-9]{{6}}.{extension}", fr".{extension}", new_name)
        os.rename(f"{self.archive_folder}{file}", f"{self.archive_folder}{new_name}")
        return new_name


class Data_manipulation:

    def __init__(self, date_range):
        self.date_range = date_range
        # To become a list of dataframes
        self.dataframes = []
        # Initialise file manipulation class to access attributes
        FM = File_manipulation()

        # if  a file is a csv, read it in as a dataframe in the init method
        for file in os.listdir("{}extract_files\\".format(FM.in_folder)):
            if file.endswith(".csv"):
                path = f"{FM.in_folder}extract_files\\{file}"
                self.dataframes.append(pd.read_csv(path))


    def union(self):
        # Unifies data into one table to make copying over a simpler and faster process.
        try:
            unioned = pd.concat(self.dataframes)
            return unioned
        except:
            messagebox.showerror("No Files within day range.", "There are no  customers files that exist for within the number of days specified. Please try again.")
            raise BaseException("No Files within day range.", "There are no  customers files that exist for within the number of days specified. Please try again.")



def check_mod_time(OneDrive_path):
    if os.path.isfile(OneDrive_path):
        mod_time = os.path.getmtime(OneDrive_path)
        date_last_run = datetime.fromtimestamp(mod_time).date()
        date_now_running = datetime.now().date()

        if date_last_run < date_now_running:
            return
        else:
            messagebox.showerror("Already ran today", r'''The  customers Automation has already been run today.
    If you think there may have been an error adding the new records, please add them manually using the union file in
    the the following folder:
    C:\Users\{}\Documents\Projects\Python\_customers_Automation\inputs\extract_files'''.format(getpass.getuser()))
            os.system('start EXCEL.EXE "inputs\extract_files\Daily_customers_union.csv"')
            exit()
        

