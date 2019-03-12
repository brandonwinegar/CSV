import os
import re
import inspect


def get_io_type(line):
    return line[0].rstrip().strip('\'"')


def get_channel_name(line):
    return line[1].rstrip().strip('\'"')


def get_channel_key(line):
    return line[2].rstrip().strip('\'"')


def get_channel_val(line):
    return line[3].rstrip().strip('\'"')


def assemble_prog_id(module_number, channel_number):
    if (module_number == "") or (channel_number == ""):
        return ""
    else:
        return '{}{}'.format(module_number.rstrip().strip('\'"'), channel_number.strip('\'"'))


def process_file(file):
    pattern = re.compile('(.*)\.(.*)')
    result = pattern.match(file)
    return result.group(1)


settings = {}


def process_settings(settings_file):
    # Create a regex pattern to match all text to the left and right of '='s
    pattern = re.compile('(.*)=(.*)')
    with open(settings_file) as file:
        for line in file:
            result = pattern.match(line)
            settings[result.group(1)] = result.group(2)


def validate_name(file_name):
    pattern = re.compile(r'(.*)[_Copy(\d+)]?.xls')
    result = pattern.match(file_name)
    copy_num = 1
    while os.path.exists(os.getcwd() + '\\Converted\\{}'.format(file_name)):
        file_name = '{}_Copy{}.xls'.format(result.group(1), copy_num)
        copy_num += 1
    return file_name
