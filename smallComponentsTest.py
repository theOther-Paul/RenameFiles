import os
from time import sleep
import pandas as pd
import pathlib
import re
from os import devnull

# global variables


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


def main():
    SPEC_PATH = input("path: ")
    SPEC_PATH = pathlib.Path(SPEC_PATH).resolve()  # will access the given path
    #SPEC_PATH = SPEC_PATH.replace("\\", r'/')
    file_count = sum(len(files) for _, _, files in os.walk(SPEC_PATH))
    counter = 0
    user_choice = input(
        "Would you like the new name to contain a string before the date?(y/n) ").lower()

    while counter <= file_count:
        if user_choice.isnumeric():
            print("Please enter a valid option!\nAvailable options are 'y' or 'n'.\nThe program will rerun to validate your data.")
            return main()
        else:
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
                            counter += 1
                        except OSError as e:
                            print(e)
                            counter += 1
                else:
                    if user_choice == 'n':
                        for x in os.listdir(SPEC_PATH):
                            full_path = os.path.join(SPEC_PATH, x)
                            extension = pathlib.Path(full_path).suffix
                            custom_name = extract_date(full_path)
                            head_string = ''
                            new_name = f"{head_string}{custom_name}{extension}"
                            try:
                                os.rename(os.path.join(SPEC_PATH, x),
                                          os.path.join(SPEC_PATH, new_name))
                                print("Success!")
                                counter += 1
                            except OSError as e:
                                print(e)
                                counter += 1

        if counter == file_count:
            break


if __name__ == '__main__':
    main()
