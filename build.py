import re
\
def str_to_list(string):
    string = string.replace("(", "")
    string = string.replace(")", "")
    string = string.replace(",", " ")
    string_list = string.split()
    return [int(x) for x in string_list]

def num_list_to_str(num_list):
    str_list = [str(x) for x in num_list]
    string = ','.join(str_list)
    string = f'({string[:len(string)]})'
    return string

def build_up(type_name, file_content):
    # filevers
    ver = re.findall(type_name + '=(\([0-9]{1,4}, [0-9]{1,4}, [0-9]{1,4}, [0-9]{1,6}\))', file_content)
    ver_tuple = ver[0]
    initial_ver = f'{type_name}={ver_tuple}'

    ver_list = str_to_list(ver_tuple)

    ver_list[3] += 1

    ver_str = num_list_to_str(ver_list)
    new_ver = f'{type_name}={ver_str}'

    file_content = file_content.replace(initial_ver, new_ver)

    return file_content


if __name__ == "__main__":
    verinfo_file = 'file_version_info.txt'
    # Open a file: file
    with open(verinfo_file, mode='r') as file:
        verfile_content = file.read()

    verfile_content = build_up('filevers', verfile_content)
    verfile_content = build_up('prodvers', verfile_content)

    with open(verinfo_file, 'w') as f:
        f.write(verfile_content)





