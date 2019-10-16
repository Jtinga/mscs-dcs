# -*- coding: utf-8 -*-
"""
Created on Sat Oct 12 15:17:40 2019

@author: Steven
"""

from school_sorting_section import SchoolSortingSection

input_directory = "C:\EdwardT"
number_of_students = 10
boys_ = []
girls_ = []
section_list_ = []
sections = ['Honesty']

sss = SchoolSortingSection()
try:
    boys_, girls_, section_list_ = sss.get_directory_data(input_directory)
except Exception as e:
    print(e)
    exit

number_of_students_per_section = number_of_students

sorted_boys = sss.sorting(boys_)
sorted_girls = sss.sorting(girls_)

section_list = sections

new_folder_path_ = sss.execute_folder(input_directory)

len_boys = len(sorted_boys)
len_girls = len(sorted_girls)

boys_per_section_, girls_per_section_ = sss.ratio(len_boys, len_girls, number_of_students_per_section)

sss.classify(section_list, boys_per_section_, girls_per_section_, sorted_boys, sorted_girls, new_folder_path_)
    