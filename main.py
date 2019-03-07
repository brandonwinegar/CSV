from Helper import *
import xlsxwriter
import re
import shutil
from time import gmtime, strftime
import os


def main(file):
    base_path = os.getcwd()
    file_path = '{}\\To Convert\\{}.csv'.format(base_path, file)
    file_name = file + '_generated.xls'
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
        # Execute if reading data from a new channel
        try:
            if current_channel != get_channel_name(currentLine):
                current_channel = get_channel_name(currentLine)
                channel_data_dict[current_channel] = {}
            channel_data_dict[current_channel][get_channel_key(currentLine)] = get_channel_val(currentLine)
        except:
            pass

    # Create the workbook for the data to be written to
    workbook = xlsxwriter.Workbook(file + '_generated.xls')
    worksheet = workbook.add_worksheet()

    # Standard Headers
    group_names = ["ProgID", "Tag", "Detail", "Offset", "Scaling", "TaskName", "Group", "CalDate", "Cab Connector"]

    # Write Headers to worksheet
    for i in range(len(group_names)):
        worksheet.write(0, i, group_names[i])

    # Write Date to worksheet
    line_number: int = 1
    for key in channel_data_dict:
        for i in range(len(group_names) - 1):
            try:
                current_tag = group_names[i+1]
                info = channel_data_dict[key][current_tag]
                worksheet.write(line_number, i + 1, info)
            except KeyError:
                # Error will be thrown if no data is present, will write nothing to cell.
                worksheet.write(line_number, i + 1, "N/A")
        line_number += 1
    workbook.close()

    # Place the file into the converted folder and account for any duplicates
    validated_name = validate_name(file_name)
    shutil.move('{}\\{}'.format(base_path, file_name),
                '{}\\{}\\{}'.format(base_path, 'Converted', validated_name))

    print("\nConversion complete. {} is now in the 'Converted' folder.".format(validated_name))
    print(strftime("%Y-%m-%d %H:%M:%S", gmtime()))


if __name__ == "__main__":
        settingsFile = os.getcwd() + '\\settings.txt'.format()
        process_settings(settingsFile)
        main(settings['file'])

