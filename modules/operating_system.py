import platform
from enum import Enum
import os

class System(Enum):
    Linux = 0
    Mac = 1
    Windows = 2

def detect_os():
    sys = platform.system()
    if sys == 'Windows':
        response = System.Windows
    elif sys == 'Darwin':
        response = System.Darwin
    else:
        response = System.Linux
    return response


def detect_user():
    return os.getlogin()


def set_globals():

    user_list = find_user_settings_files()
    current_user = detect_user()
    current_os = detect_os()

    txt_file_list = []
    if current_user in user_list:
        file_name = f"files/{current_user}.txt"
        with open(file_name) as f:
            lines = f.readlines()

        for line in lines:
            txt_file_list.append(line.replace('\n', ''))
    else:
        print('Tuntematon käyttäjä')
        print('Luo tiedosto käyttäjänimelläsi projektin kansioon files, katso mallia muilta käyttäjiltä')
        print('Käyttöjärjetelmän alla on 2 polkua, joista 1. on alkuperäinen ja 2. uuden muokatun.')
        print('.json tiedoston luodaan aina projektin kansioon "files", mutta sen nimi ja polku on annettava.')

    if current_os == System.Windows:
        pos1 = txt_file_list.index('[Windows]')
        original_file_path = txt_file_list[pos1 + 1]
        modified_file_path = txt_file_list[pos1 + 2]
        pos2 = txt_file_list.index('[json]')
        json_copy_path = txt_file_list[pos2 + 1]
    else:
        # must be Linux then
        pos3 = txt_file_list.index('[Linux]')
        original_file_path = txt_file_list[pos3 + 1]
        modified_file_path = txt_file_list[pos3 + 2]
        pos4 = txt_file_list.index('[json]')
        json_copy_path = txt_file_list[pos4 + 1]

    return original_file_path, modified_file_path, json_copy_path


def find_user_settings_files():
    users = []
    files_list = os.listdir('files')
    for file in files_list:
        if '.txt' in file:
            users.append(file.rstrip('.txt'))
    return users
