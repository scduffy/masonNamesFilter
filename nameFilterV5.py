import pandas as pd
from os import path
import sys

# name_filter list containing elements acceptable to have within last names (add to it if desired)
name_filter = ['jr', 'sr', 'von', 'van', 'mac', 'st', 'mc', 'de', 'la', 'du', 'le', '2nd', '3rd', 'ii', 'iii', 'lodge',
               'admission', 'admissions', 'des', 'di', '#', 'mc\'', 'o\'', 'd\'', 'lll', '4th', '5th', '6th',
               '#1', '#2', '#3', '#4', '#5', 'le', 'de', '111', '11', '3', '2']
error_dict = {}


def increment_dict(key):
    # if the Lodge is not counted in dictionary yet, add it with 1 error as its data
    if key not in error_dict.keys():
        error_dict[key] = 1
    # otherwise find the Lodge and increment the number of errors by 1
    else:
        error_dict[key] += 1


# writes file path previously used
def write_file_path(path_to_file):
    # open file in write mode, then write the file path used this iteration of the program
    if path_to_file != 'pandas_simple.xlsx':
        f = open("old_excel_file_path.txt", "w")
        f.write(path_to_file)
        f.close()  # close the file to save changes


# reads file path previously used
def read_file_path():
    # open file in read mode and parse old file path
    f = open("old_excel_file_path.txt", "r")
    path_to_file = f.read()
    f.close()  # close file to save
    return path_to_file  # return old file path to be opened again in other part of program


# function checks if a given name is valid
def is_name_valid(name_in):
    name_list = name_in.split()  # split the last name where there is whitespace and return as a list of sub-names

    # checks if last name is more than one word (check if the sub-name list has more than 1 element)
    if len(name_list) > 1:
        # iterate over parts of last name
        for string in name_list:
            # check parts of last name against filter of valid last name elements
            # strips off '.' if possible to account for elements like 'jr' and 'jr.' being the same thing
            # if name_filter.__contains__(string.strip().strip(',').strip('.').lower()):
            if '?' in string:
                return False
            if name_filter.__contains__(string.lower().replace(',.', '')):
                return True  # return True if the name has a valid part as specified by filter
        return False  # return False if the name does not have a part described in filter, name invalid
    return True  # name length is < 2 so it is automatically considered valid


def first_name_valid(name_in):
    name_list = name_in.split()

    titles = {'jr', 'sr'}

    if len(name_list) > 1:
        # iterate over parts of last name
        for string in name_list:
            if titles.__contains__(string.lower().replace(',.', '')):
                return False
    return True


# gets the file path and filename from user and checks that it exists and
# is a valid excel file
def get_filename():
    # check if user wants to use previously used file path
    old_file_path = read_file_path()
    string = 'Do you want to use previous file path \'' + old_file_path + '\'? (y/n): '
    use_old_file = input(string)  # get input from user to determine if they want to use old file path

    # old file logic
    if use_old_file.lower() == 'y':
        return old_file_path  # if using old file path, return it and stop further input from user
    elif use_old_file.lower() != 'n':
        print("Invalid input. Assuming new file path. ")

    # loop until valid file path given or quit on 'q'
    while True:
        # get file path from user and strip off '"' from path if present
        filename = input('Enter file path to excel file or \'q\' to quit: ').strip("\"")
        # quit if input is q
        if filename.lower() == 'q':
            print('Exiting gracefully...')
            sys.exit(0)
        # check file path and file type
        elif path.exists(filename) and filename.endswith('.xlsx'):
            print('File path valid. Filtering names...')
            return filename  # return file path if it is valid
        else:
            # file path or file name was invalid so keep looping until valid path found
            print('Invalid file path or file name. ')


# finds row index of the given first and last name in the data frame and returns it
def get_row_index(data_frame, fist_name, lst_name):
    matching = data_frame.loc[data_frame['Last Name'] == lst_name].index.values

    for x in matching:
        if data_frame.loc[data_frame.index[x], 'First Name'] == fist_name:
            return x
    raise IndexError('No matching values found.')


def mark_names(file_name):
    data = pd.read_excel(file_name)
    internal_tally = 0

    if 'Highlight' not in data:
        data['Highlight'] = ''

    for index, row in data.iterrows():
        # try/except in case of exceptions being thrown during runtime
        try:
            if not is_name_valid(row['Last Name']) or not first_name_valid(row['First Name']):
                last_name = row['Last Name']
                first_name = row['First Name']
                idx = get_row_index(data, first_name, last_name)  # getting row number so it can swap elements
                print('Marking', first_name, last_name)
                internal_tally += 1
                data.at[idx, 'Highlight'] = 'Error'
        # catch AttributeError exception thrown when a cell is NaN or empty
        except AttributeError:
            continue  # if there is an AttributeError, ignore the row and keep going
        # catch KeyError exception thrown when 'Last Name' or 'Lodge' column cannot be found
        except KeyError:
            print('Error: \'Last Name\' or \'Lodge\' column could not be found. Verify spreadsheet column formatting. ')
            print('Exiting...')
            sys.exit(0)
        except IndexError:
            continue  # no matching values found in get index just continue for now TODO: fix this
        # general exception catch of all unknown exceptions
        except Exception as e:
            print('\nError: Unexpected exception thrown.')
            print('Exception message:', str(e))
            print('Exiting...')
            sys.exit(-1)

    print('Saving excel file...')
    writer = pd.ExcelWriter('corrected.xlsx', engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    print('Updated file saved to corrected.xlsx')
    return internal_tally


def write_output_to_file(filename, output_str):
    filename = filename.split('\\')
    filename = filename[-1].strip('.xlsx') + '.txt'
    print('Saving to:', filename, '\n')
    f = open(filename, "w+")
    f.write(output_str)
    f.close()  # close the file to save changes


# Start of script:
# =====================================================================================================================
# opens and reads the excel file, puts file data into data frame obj
valid_filename = get_filename()
tally = mark_names(valid_filename)

print()
print(tally, 'names marked.\n')
print('Exiting gracefully...')
