#!/usr/bin/env/ python
from pathlib import Path
from Helper import *
import re
import xlsxwriter
import sys
import os

channelDataDict = {}    # Highest level Dict


def main(file_path):
    # Opens the selected file and splits elements on a '.' or ',' and organizes the data into lists
    with open(file_path) as f:
        channel_list = list(map(lambda n: re.split('[,.]', n), f.readlines()))
        current_channel = get_channel_name(channel_list[0])

    # Creates a new Key:Value pair in the master dictionary where
    # Key is the current channel name and the value is an empty dict
    channelDataDict[current_channel] = {}

    # This loop will assemble data from the old format of CSV into dictionaries to be referenced later
    for currentLine in channel_list:
        # Execute if reading data from a new channel
        if current_channel != get_channel_name(currentLine):
            current_channel = get_channel_name(currentLine)
            channelDataDict[current_channel] = {}
        channelDataDict[current_channel][get_channel_key(currentLine)] = get_channel_val(currentLine)

    # Create the Workbook
    workbook = xlsxwriter.Workbook('TestCSV.xls')
    worksheet = workbook.add_worksheet()
    group_names = ["ProgID", "Tag", "Detail", "Offset", "Scaling", "TaskName", "Group", "CalDate", "Cab Connector"]

    # Write Headers
    for i in range(len(group_names)):
        worksheet.write(0, i, group_names[i])

    # Write Data
    line_number = 1
    for key in channelDataDict:
            for i in range(len(group_names)):
                try:
                    worksheet.write(line_number, i + 1, channelDataDict[key][group_names[i+1]])
                except:
                    worksheet.write(line_number, i + 1, "N/A")
            line_number += 1

    workbook.close()
    print("Conversion complete.")


if __name__ == "__main__":
    if len(sys.argv) <= 1:
        print("Please provide path to file to convert as a command line argument")
        print("Use 'help' as an argument for more information")
        exit()
    if sys.argv[1] == "Help" or sys.argv[1] == "help":
        print("Provide an exact path to the file you want to convert.")
        print("If you want to convert a file in the same location as this file ")
        print("pass 'Here' and then the name of the file")
        exit()
    if sys.argv[1] == "Here" or sys.argv[1] == "here":
        try:
            filePath = os.getcwd() + '/' + sys.argv[2]
        except IndexError:
            print("Please provide path to file to convert as a command line argument after 'Here'")
    else:
        filePath = sys.argv[1]

    if os.path.exists(filePath):
        main(filePath)
    else:
        print("Invalid path")
        exit()


