import os
import pandas as pd
import pathlib
import re

# global variables
SPEC_PATH = input("path: ")


def format_date(extracted_string):
    new_s = ''
    wip_s = re.sub(r"[^a-zA-Z0-9 ]", "", extracted_string)
    wip_s = ''.join(c for c in wip_s if c.isnumeric())
    wip_s = wip_s[:len(wip_s)//2]
    # will slice at 4th index and the 6th and will concatenate the remaining 3 before 4th and the rest after 5th
    extracted_string = wip_s[:4]+wip_s[6:]
    return extracted_string


def extract_date(path):  # needs to get the full path of the file from the folder
    try:
        df = pd.read_excel(path, sheet_name='Synthèse')  # df=dataframe
        return format_date(df['Unnamed: 1'].values[3])

    except OSError as e:
        print(e)


# main function
SPEC_PATH = pathlib.Path(SPEC_PATH).resolve()  # will access the given path
#SPEC_PATH = SPEC_PATH.replace("\\", r'/')

user_choice = input(
    "Would you like the new name to contain a string before the date?(y/n) ").lower()

while True:
    if user_choice.isalpha():
        if user_choice == 'y':
            head_string = input("Please enter the string: ")

            # renaming operation
            for x in os.listdir(SPEC_PATH):
                full_path = os.path.join(SPEC_PATH, x)
                extension = pathlib.Path(full_path).suffix
                custom_name = extract_date(full_path)
                new_name = f"{head_string}{custom_name}{extension}"
                try:
                    os.rename(os.path.join(SPEC_PATH, x),
                              os.path.join(SPEC_PATH, new_name))
                    print("Success!")
                except OSError as e:
                    print(e)

# need to work on the else condition and breaking loop
        # else:
        #     if user_choice == 'n':
        #         # will count all the files that contain a specified extension
        #         for files in os.listdir(SPEC_PATH):
        #             # might include all xl extensions just to be safe
        #             if files.endswith(".xls") or files.endswith(".xlsx"):
        #                 extract_date(files)
        #                 format_date(files)
        #                 if files.isfile():
        #                     custom_name = extract_date(files)
        #                     new_name = f"{files.suffix}"
        #                     if SPEC_PATH.isfile():
        #                         continue
        #                     else:
        #                         try:
        #                             x.rename(SPEC_PATH/new_name)
        #                             print("Success!")
        #                         except OSError as e:
        #                             print(e)
    # else:
    #     if user_choice.isnumeric() or user_choice != 'y' or user_choice != 'n':
    #         print("Please enter a valid choice!")
    #         break
    break
