import os
import json
import school_sorting_section as sss
from datetime import datetime
from shutil import copy


def execute_folder(old_folder_path):
    if not os.path.exists(old_folder_path):
        raise NotADirectoryError
    new_folder = old_folder_path + "_" + datetime.now().strftime("%Y%m%d%H%M%S")
    os.mkdir(new_folder)
    return new_folder


def classify(section_list, boys_per_section, girls_per_section, sorted_list_boys, sorted_list_girls, new_folder_path):
    counter_boys = 0
    counter_girls = 0
    boys_list = []
    girls_list = []

    for section in section_list:
        # Section level
        print("Section level: ", section)
        section_folder_name = create_section_folder(section, new_folder_path)

        # Insert boys into the current section in the loop
        for boy in sorted_list_boys[counter_boys:counter_boys+boys_per_section]:
            print(boy[4])
            copy_student_excel_file_old_to_new(boy[4], section_folder_name + "\\" + "BOYS")
            boys_list.append((boy[0], get_old_section(boy[4], "BOYS"), section))
        counter_boys += boys_per_section  # move on to the next batch sequence for the next section

        # Insert girls into the current section in the loop
        for girl in sorted_list_girls[counter_girls:counter_girls+girls_per_section]:
            print(girl[4])
            copy_student_excel_file_old_to_new(girl[4], section_folder_name + "\\" + "GIRLS")
            girls_list.append((girl[0], get_old_section(girl[4], "GIRLS"), section))
        counter_girls += girls_per_section  # move on to the next batch sequence for the next section

        # Call the generate new section list method
        sss.section_list_file_creator(boys_list, girls_list, section_folder_name)

        girls_list = []
        boys_list = []


def get_old_section(old_section_file_path, gender):
    index = old_section_file_path.split("\\").index(gender) - 1
    old_section_name = old_section_file_path.split("\\").pop(index)

    return old_section_name


def create_section_folder(section_name, new_folder_path):

    dir_name = new_folder_path + "\\" + section_name

    try:
        # Create target Directory
        os.mkdir(dir_name)
        # print("Directory ", dir_name, " Created ")
        create_boys_girls_folder(dir_name)

        return dir_name
    except FileExistsError:
        print("Directory ", dir_name, " already exists")


def create_boys_girls_folder(section_folder_path):

    section_folder_path_boys = section_folder_path + "\\" + "BOYS"
    section_folder_path_girls = section_folder_path + "\\" + "GIRLS"

    # BOYS
    try:
        # Create target Directory
        os.mkdir(section_folder_path_boys)
        print("Directory ", section_folder_path_boys, " Created ")
    except FileExistsError:
        print("Directory ", section_folder_path_boys, " already exists")

    # GIRLS
    try:
        # Create target Directory
        os.mkdir(section_folder_path_girls)
        print("Directory ", section_folder_path_girls, " Created ")
    except FileExistsError:
        print("Directory ", section_folder_path_girls, " already exists")


def copy_student_excel_file_old_to_new(old_path, new_path):
    copy(old_path, new_path)
