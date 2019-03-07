import os
import re


def get_io_type(line):
    return line[0]


def get_channel_name(line):
    return line[1]


def get_channel_key(line):
    return line[2]


def get_channel_val(line):
    return line[3]


settings = {}


def process_settings(settings_file):
    # Create a regex pattern to match all text to the left and right of '='s
    pattern = re.compile(r'(.*)=(.*)\.[A-Za-z]*')
    with open(settings_file) as file:
        for line in file:
            result = pattern.match(line)
            try:
                settings[result.group(1)] = result.group(2)
            except:
                pass


def validate_name(file_name):
    pattern = re.compile(r'(.*)[_Copy(\d+)]?.xls')
    result = pattern.match(file_name)
    copy_num = 1
    while os.path.exists(os.getcwd() + '\\Converted\\{}'.format(file_name)):
        file_name = '{}_Copy{}.xls'.format(result.group(1), copy_num)
        copy_num += 1
    return file_name
