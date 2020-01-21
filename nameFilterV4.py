import pandas as pd
from os import path
import sys

# name_filter list containing elements acceptable to have within last names (add to it if desired)
name_filter = ['jr', 'sr', 'von', 'van', 'mac', 'st', 'mc', 'de', 'la', 'du', 'le', '2nd', '3rd', 'ii', 'iii', 'iv',
               'admission', 'admissions', 'des', 'di', '#', 'mc\'', 'o\'', 'd\'', 'lll', '4th', '5th', '6th', 'lodge',
               '#1', '#2', '#3', '#4', '#5', 'le', 'de', '111', '11', '3', '2']
tally = 0
error_dict = {}


def increment_dict(key):
    # if the lodge is not counted in dictionary yet, add it with 1 error as its data
    if key not in error_dict.keys():
        error_dict[key] = 1
    # otherwise find the lodge and increment the number of errors by 1
    else:
        error_dict[key] += 1


def print_critical_sections():
    # print header for columns of critical sections
    output = ''
    print('Critical Lodge Sections: ')
    print('\n{0:4}    {1}'.format('Lodge', 'Errors'))
    print('================')

    output += 'Critical Lodge Sections: '
    output += '\n{0:4}    {1}'.format('Lodge', 'Errors')
    output += '\n================\n'

    # iterate over dictionary and print lodge num and number of errors in that lodge
    for key in error_dict:
        # only print if the # of errors exceeds a certain amount (can be changed if desired)
        if error_dict[key] >= 3:
            print('{0:4}       {1}'.format(key, error_dict[key]))
            output += '{0:4}       {1}'.format(key, error_dict[key]) + '\n'
    return output


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
    name_list = name_in.lower().replace('(', '').replace(')', '').replace('.', '').replace(',', '').split()

    # if there is a ? it is automatically wrong
    if '?' in name_list or '?' in name_in:
        return False

    # checks if last name is more than one word (check if the sub-name list has more than 1 element)
    if len(name_list) > 1:
        # case where name is something like 'o neil' or 'm donald' which is correct
        if name_list[0].lower() == 'o' or name_list[0].lower() == 'm':
            return True
        # now check for filter values
        for string in name_filter:
            if string in name_list:
                return True
        return False
    return True


def first_name_valid(name_in):
    name_list = name_in.split()  # split up the name into a list

    titles = {'jr', 'sr', 'mc'}

    # if there is a ? it is wrong
    if '?' in name_list or '?' in name_in:
        return False

    # if the list has more than one element then there is an error probably
    if len(name_list) > 1:
        # iterate over parts of last name
        for string in name_list:
            if titles.__contains__(string.lower().replace('.', '').replace(',', '')):
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
        print('Awaiting further input...')
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

    # iterate over rows till match is found
    for row in matching:
        if data_frame.loc[data_frame.index[row], 'First Name'] == fist_name:
            return row
    raise IndexError('No matching values found.')


def check_names(file_name):
    print('\n{1:15}  {0}  {2}'.format('Last Name', 'First Name', 'Lodge #'))
    print('================================================')

    output = '\n{1:15}  {0}  {2}'.format('Last Name', 'First Name', 'Lodge #')
    output += '\n================================================\n'

    error_dict.clear()
    num_invalid_names = 0

    # catch possible file error that was missed earlier because of a logic error :P
    try:
        data = pd.read_excel(file_name)
    except FileNotFoundError:
        print('FileNotFoundError: File path possibly changed or does no longer exist.')
        sys.exit(-1)  # quit program because it can't continue

    # iterate over rows collected from file in data frame
    for index, row in data.iterrows():
        # try/except in case of exceptions being thrown during runtime
        try:
            # check if first name and last name are the same
            if row['Last Name'] == row['First Name']:
                print('{1:15}  {0}   {2}  ID'.format(row['First Name'], row['First Name'], row['LODGE']))
                increment_dict(row['LODGE'])
                num_invalid_names += 1  # keep tally of invalid names for sanity reasons
            # check if the last name in current row is valid or not
            if not is_name_valid(row['Last Name']) or not first_name_valid(row['First Name']):
                # truncate last_name string if it is too long to preserve proper output formatting
                last_name = (row['Last Name'][0:12] + '...') \
                    if len(str(row['Last Name'])) >= 15 else row['Last Name']
                first_name = (row['First Name'][0:12] + '...') \
                    if len(str(row['First Name'])) >= 15 else row['First Name']

                # earlier excel books have strange formatting and some use character wrapping, this fixes that
                first_name = first_name.replace('\r\n', '').replace(chr(10), '')
                last_name = last_name.replace('\r\n', '').replace(chr(10), '')

                # last name is invalid so print last name and lodge number to console
                print('{1:15}  {0:15}   {2}'.format(last_name, first_name, row['LODGE']))
                output += '{1:15}  {0:15}  {2}'.format(last_name, first_name, row['LODGE']) + '\n'
                increment_dict(row['LODGE'])
                num_invalid_names += 1  # keep tally of invalid names for sanity reasons
        # catch AttributeError exception thrown when a cell is NaN or empty
        except AttributeError:
            continue  # if there is an AttributeError, ignore the row and keep going
        # catch KeyError exception thrown when 'Last Name' or 'LODGE' column cannot be found
        except KeyError:
            print('Error: \'Last Name\' or \'LODGE\' column could not be found. Verify spreadsheet column formatting. ')
            print('Exiting...')
            sys.exit(-1)
        except IndexError:
            continue  # no matching values found in get index just continue for now
        # general exception catch of all unknown exceptions
        except Exception as e:
            print('\nError: Unexpected exception thrown.')
            print('Exception message:', str(e))
            print('Exiting...')
            sys.exit(-1)

    output += '\n' + print_critical_sections()
    output += '\n\n' + str(num_invalid_names) + ' total invalid names found.'
    print(str(num_invalid_names) + ' total invalid names found.')

    write_file_path(valid_filename)

    return output


def swap_names(file_name):
    print('\nSearching for invalid names...')

    num_swapped_name = 0

    data = pd.read_excel(file_name)

    for index, row in data.iterrows():
        # try/except in case of exceptions being thrown during runtime
        try:
            if not is_name_valid(row['Last Name']) or not first_name_valid(row['First Name']):
                last_name = row['Last Name']
                first_name = row['First Name']
                copy = last_name

                idx = get_row_index(data, first_name, last_name)  # getting row number so it can swap elements
                data.at[idx, 'Last Name'] = first_name
                data.at[idx, 'First Name'] = copy

                num_swapped_name += 1  # keep tally of invalid names for sanity reasons

                print('Swapping', first_name, 'and', last_name)
        # catch AttributeError exception thrown when a cell is NaN or empty
        except AttributeError:
            continue  # if there is an AttributeError, ignore the row and keep going
        # catch KeyError exception thrown when 'Last Name' or 'LODGE' column cannot be found
        except KeyError:
            print('Error: \'Last Name\' or \'LODGE\' column could not be found. Verify spreadsheet column formatting. ')
            print('Exiting...')
            sys.exit(0)
        except IndexError:
            continue  # no matching values found in get index just continue for now
        # general exception catch of all unknown exceptions
        except Exception as e:
            print('\nError: Unexpected exception thrown.')
            print('Exception message:', str(e))
            print('Exiting...')
            sys.exit(-1)

    print('Saving excel file...')
    writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()
    print('Updated file saved to pandas_simple.xlsx')
    print(num_swapped_name, 'total names swapped.\n')


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

# loop for input
while True:
    selection = input('f - filter names and print invalid\n'
                      's - swap invalid names\n'
                      'q - quit program\n'
                      'selection: ')

    if selection.lower() == 'f':  # filter case
        out = check_names(valid_filename)
        write_output = input('Do you want to write output to file? (y/n):')
        if write_output.lower() == 'y' or write_output.lower() == 'yes':
            write_output_to_file(read_file_path(), out)
    elif selection.lower() == 's':  # swapping case
        swap_names(valid_filename)
        print('Switching focus to output file.\n')
        valid_filename = 'pandas_simple.xlsx'
    elif selection.lower() == 'q':  # quitting case
        break
    else:  # user done messed up the input somehow
        print('Invalid input. Try again.')

print('Exiting gracefully...')
