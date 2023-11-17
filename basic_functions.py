#!/usr/bin/env python3

#*************************************************************#
#                                                             #
#  Written by Yuri H. Galvao <yuri@galvao.ca>, November 2022  #
#                                                             #
#*************************************************************#

import os, json, sys, google.cloud.logging, logging, traceback

#client = google.cloud.logging.Client()

#client.setup_logging()

logging.basicConfig(
    filename='log.log',
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
    )
console_handler = logging.StreamHandler(sys.stdout)
logging.getLogger().addHandler(console_handler)

args = sys.argv[1:] # List of arguments that were passed

yes_for_all = True if '--yes-for-all' in args else False

def confirm(text:str)->bool:
    """
    Asks the user for confirmation using a provided text prompt.

    Arguments:
        text (str): the text prompt to display to the user.

    Returns:
        bool: True if the user confirms, False otherwise.
    """

    confirm = 'y' if yes_for_all else input(text)
    
    if confirm.lower() not in ('n', 'no'):
        return True
    else:
        return False

def check_file(file:str)->bool:
    """
    Checks if a file with the given name exists in the current directory.

    Arguments:
        file (str): the name of the file to check.

    Returns:
        bool: True if the file exists, False otherwise.
    """

    if os.path.isfile('./' + file):
        return True
    else:
        return False

def ask_for_data(required_data:tuple, file_name_no_extension:str, ask:bool=True)->dict:
    """
    Asks the user for the required data and saves it to a configuration file.

    Arguments:
        required_data (tuple): a tuple containing the names of the required data.
        file_name_no_extension (str): the name of the configuration file without the file extension.
        ask (bool, optional): if True, prompts the user for the required data. If False, uses default values.

    Returns:
        dict: a dictionary containing the provided data.
    """

    data_dict = {}
    if ask:
        for data in required_data:
            data_dict[data] = input(f'Enter the {data}: ')

        open(file_name_no_extension+'.json', 'w').write(json.dumps(data_dict))
    else:
        for data in required_data:
            data_dict[data[0]] = data[1]

        open(file_name_no_extension+'.json', 'w').write(json.dumps(data_dict))
    
    print()

    return data_dict

def list_from_input(text:str)->list:
    """
    Converts a user's input, containing comma-separated values, into a list.

    Arguments:
        text (str): the text prompt to display to the user.

    Returns:
        list: a list containing the user's input values as integers or strings.
    """
    raw_list = input(text)

    try:
        final_list = [int(n) for n in raw_list.replace(' ', '').split(',')]
    except:
        final_list = [str(s) for s in raw_list.replace(' ', '').split(',')]

    return final_list
