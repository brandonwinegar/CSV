from Helper import *
import xlsxwriter
import re
import shutil
from time import *
import time
import os


def main(file_to_convert):
    base_path = os.getcwd()
    file_path = '{}\\To Convert\\{}'.format(base_path, file_to_convert)
    file_name = process_file(file_to_convert) + '.xls'
    # Opens the selected file and splits elements on a '.' or ',' and organizes the data into lists
    with open(file_path) as f:
        channel_list = list(map(lambda n: re.split('[,.\t]', n), f.readlines()))
        current_channel = get_channel_name(channel_list[0])
    #############################################################
    # Creates a master dictionary where each key will           #
    # be a channel name, each value will be a Dict containing   #
    # information  on that specific channel                     #
    #############################################################
    channel_data_dict = {current_channel: {}}
    # This loop will assemble data from the old format of CSV into dictionaries to be referenced later
    for currentLine in channel_list:
        try:

            # Execute if currentLine is from a new channel
            if current_channel != get_channel_name(currentLine):
                # Execute before moving on to next channel
                try:
                    module_number = channel_data_dict[current_channel]['Module']
                    channel_number = channel_data_dict[current_channel]['Channel']
                    channel_data_dict[current_channel]['ProgID'] = assemble_prog_id(module_number, channel_number)
                except:
                    pass
                # Moving on to next channel
                current_channel = get_channel_name(currentLine)
                channel_data_dict[current_channel] = {}
                channel_data_dict[current_channel]['Type'] = get_io_type(currentLine)
            channel_data_dict[current_channel][get_channel_key(currentLine)] = get_channel_val(currentLine)
        except:
            pass
    # Create the workbook for the data to be written to
    workbook = xlsxwriter.Workbook(process_file(file_to_convert) + '.xls')
    worksheet = workbook.add_worksheet()
    # Standard Headers
    group_names = ['ProgID', 'Tag', 'Detail', 'Type', 'Offset', 'Scaling', 'TaskName', 'Group', 'CalDate', 'Cab Connector']
    # Write Headers to worksheet    (row, col, val)
    for i in range(len(group_names)):
        worksheet.write(0, i, group_names[i])
    # Write Date to worksheet
    line_number: int = 1
    for key in channel_data_dict:
        for i in range(len(group_names)):
            try:
                current_tag = group_names[i]
                info = channel_data_dict[key][current_tag].strip()
                worksheet.write(line_number, i, info)
            except KeyError:
                # Error will be thrown if no data is present, will write nothing to cell.
                worksheet.write(line_number, i, '')
        line_number += 1
    workbook.close()
    # Place the file into the converted folder and account for any duplicates
    validated_name = validate_name(file_name)
    shutil.move('{}\\{}'.format(base_path, file_name),
                '{}\\{}\\{}'.format(base_path, 'Converted', validated_name))
    print("{} is now in the 'Converted' folder.".format(validated_name))


if __name__ == '__main__':
    begin_millis = int(round(time.time() * 1000))
    settingsFile = os.getcwd() + '\\settings.txt'.format()
    process_settings(settingsFile)
    print()
    if settings['DoAll']:
        file_list = settings['FileList'].split(',')
        for file in file_list:
            main(file)
    else:
        main(settings['SingleFile'])
    end_millis = int(round(time.time() * 1000))
    print('\nConversion took {} seconds'.format((end_millis-begin_millis)/1000))
    print(strftime('%Y-%m-%d %H:%M:%S', gmtime()))

