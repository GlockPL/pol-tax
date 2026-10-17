import re
from pathlib import Path
from Other.Version import Version


def str_to_list(string):
    string = string.replace("(", "")
    string = string.replace(")", "")
    string = string.replace(",", " ")
    string_list = string.split()
    return [int(x) for x in string_list]


def num_list_to_str(num_list):
    str_list = [str(x) for x in num_list]
    string = ', '.join(str_list)
    string = f'({string[:len(string)]})'
    return string


def num_list_to_str_dot(num_list):
    str_list = [str(x) for x in num_list]
    string = '.'.join(str_list)
    string = f'{string[:len(string)]}'
    return string

def build_up(type_name, file_content):
    # filevers
    ver = re.findall(type_name + r'=(\([0-9]{1,4}, [0-9]{1,4}, [0-9]{1,4}, [0-9]{1,6}\))', file_content)
    ver_tuple = ver[0]
    initial_ver = f'{type_name}={ver_tuple}'

    ver_list = str_to_list(ver_tuple)

    ver_list[3] += 1

    ver_str = num_list_to_str(ver_list)
    new_ver = f'{type_name}={ver_str}'

    file_content = file_content.replace(initial_ver, new_ver)

    return file_content


def build_up_version(ver_str: str, type_name: str, file_content: str) -> str:

    new_ver = f'{type_name}={ver_str}'

    ver = re.findall(type_name + r'=(\([0-9]{1,4}, [0-9]{1,4}, [0-9]{1,4}, [0-9]{1,6}\))', file_content)
    ver_tuple = ver[0]
    initial_ver = f'{type_name}={ver_tuple}'

    file_content = file_content.replace(initial_ver, new_ver)

    return file_content


def SringTable_version(ver_str: str, type_name: str, file_content: str) -> str:
    new_ver = f'StringStruct(\'{type_name}\', \'{ver_str}\')'

    regex_string = r'StringStruct\(\'' + type_name + r'\', \'([0-9]{1,4}\.[0-9]{1,4}\.[0-9]{1,4}\.{0,1}[0-9]{0,6})\'\)'
    ver = re.findall(regex_string, file_content)
    ver_tuple = ver[0]
    initial_ver = f'StringStruct(\'{type_name}\', \'{ver_tuple}\')'

    file_content = file_content.replace(initial_ver, new_ver)

    return file_content


if __name__ == "__main__":
    vers = Version()
    verinfo_file = 'file_version_info.txt'
    # Open a file: file
    with open(verinfo_file, mode='r') as file:
        verfile_content = file.read()

    current_version_l = list(vers.up_build())
    vers.save()
    ver_str = num_list_to_str(current_version_l)
    verfile_content = build_up_version(ver_str, 'filevers', verfile_content)
    verfile_content = build_up_version(ver_str, 'prodvers', verfile_content)

    ver_str = num_list_to_str_dot(current_version_l)
    verfile_content = SringTable_version(ver_str, 'FileVersion', verfile_content)
    verfile_content = SringTable_version(ver_str, 'ProductVersion', verfile_content)

    with open(verinfo_file, 'w') as f:
        f.write(verfile_content)

    # Keep pyproject.toml in sync with resources/version.txt
    pyproject = Path('pyproject.toml')
    pyproject_content = re.sub(r'^version = "[^"]*"', f'version = "{vers}"', pyproject.read_text(encoding='utf8'),
                               count=1, flags=re.M)
    pyproject.write_text(pyproject_content, encoding='utf8')





