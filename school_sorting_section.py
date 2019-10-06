import os
import openpyxl
from openpyxl import Workbook
from math import ceil
import datetime


class SchoolSortingSection:

    def __init__(self):
        pass

    # Input: Excel Directory Path (string)
    # Output: boys (tuple), girls (tuple)
    def get_directory_data(self, excel_directory_path):
        section_path_list = os.listdir(excel_directory_path)  # Sections
        print("SECTION", section_path_list)

        section_list = []
        error_list = []
        # (Full Name, Average, Average of 3, Old Section, File Path)
        boys = []
        girls = []
        for section_folder in section_path_list:
            section_folder_path = os.path.join(excel_directory_path, section_folder)
            section_list.append(section_folder)
            print(os.path.join(excel_directory_path, section_folder_path))
            gender_list = os.listdir(section_folder_path)  # Genders
            for gender_folder in gender_list:
                gender_path = os.path.join(section_folder_path, gender_folder)
                student_list = os.listdir(gender_path)  # Students
                for student in student_list:
                    # st = datetime.datetime.now()
                    full_name = os.path.splitext(student)[0]  # Filename
                    average, average_of_three = self.get_file_data(os.path.join(gender_path, student))
                    if average is not None or average_of_three is not None:
                        if gender_folder.lower() == "boys":
                            boys.append([full_name, average, average_of_three, section_folder, os.path.join(gender_path, student)])
                        else:
                            girls.append([full_name, average, average_of_three, section_folder, os.path.join(gender_path, student)])
                    else:
                        error_list.append([full_name, section_folder, os.path.join(gender_path, student)])
                    # et = datetime.datetime.now()
                    # print("TIME", et - st)

        return boys, girls, section_list

    # Input: Excel File Path (string)
    # Output: Average (float), Average of Math, English and Science (float)
    def get_file_data(self, excel_file_path):
        print(excel_file_path)
        if not os.path.isfile(excel_file_path):
            # print("Error: File Not Found")
            raise FileNotFoundError

        book = openpyxl.load_workbook(excel_file_path, data_only=True)
        sheet = book.active
        average = float(sheet[self.find_cell_by_text(sheet, 'General Average', 'FINAL')].value)
        math = float(sheet[self.find_cell_by_text(sheet, 'Mathematics', 'FINAL')].value)
        english = float(sheet[self.find_cell_by_text(sheet, 'English', 'FINAL')].value)
        science = float(sheet[self.find_cell_by_text(sheet, 'Science ', 'FINAL')].value)
        average_of_three = (math + english + science) / 3
        return average, average_of_three

    # Input: Spreadsheet (sheet object - openpyxl)
    # Output: Intersection of row cell and column cell (string)
    def find_cell_by_text(self, sheet, row_text, column_text):
        row_name = None
        column_name = None
        for row in sheet.iter_rows():
            for entry in row:
                try:
                    if row_text == entry.value and row_name is None:
                        row_name = entry.row
                        # print(entry.value)
                    elif column_text == entry.value and column_name is None:
                        column_name = entry.column_letter
                        # print(entry.value)
                    elif row_name is not None and column_name is not None:
                        print(row_text, "+", column_text)
                        print("Cell:", column_name + str(row_name))
                        return column_name + str(row_name)
                except Exception as e:
                    raise e
        print("Cell Not Found!")
        return None


    # Jomar
    def sorting(self, student_list):
        sorted_list = sorted(student_list, key=lambda x: (x[1], x[2]), reverse=True)
        print('Sorted', sorted_list)
        return sorted_list


    def ratio(self, boys_count, girls_count, estimation):

        if boys_count == 0:
            boys_section = 0
            girls_section = estimation
        elif girls_count == 0:
            boys_section = estimation
            girls_section = 0
        elif boys_count == 0 and girls_count == 0:
            boys_section = 0
            girls_section = 0
        elif boys_count == girls_count:
            girls_section = round(estimation / 2)
            boys_section = estimation - girls_section
        elif boys_count > girls_count:
            avg = boys_count / girls_count
            girls_section = round(estimation / ceil(avg))
            boys_section = estimation - girls_section
        else:
            avg = girls_count / boys_count
            boys_section = round(estimation / ceil(avg))
            girls_section = estimation - boys_section

        return boys_section, girls_section


    # Steven
    def section_list_file_creator(self, boys_list, girls_list, new_folder_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Section List"

        start_row = 1
        start_column = 1

        ws.cell(row=1, column=1).value = "BOYS"
        ws.cell(row=1, column=2).value = "OLD SECTION"
        ws.cell(row=1, column=3).value = "NEW SECTION"

        for boy in boys_list:
            ws.cell(row=start_row + 1, column=start_column).value = boy[0]
            ws.cell(row=start_row + 1, column=start_column + 1).value = boy[1]
            ws.cell(row=start_row + 1, column=start_column + 2).value = boy[2]
            start_row += 1

        ws.cell(row=start_row + 2, column=1).value = "GIRLS"
        ws.cell(row=start_row + 2, column=start_column + 1).value = "OLD SECTION"
        ws.cell(row=start_row + 2, column=start_column + 2).value = "NEW SECTION"

        for girl in girls_list:
            ws.cell(row=start_row + 3, column=start_column).value = girl[0]
            ws.cell(row=start_row + 3, column=start_column + 1).value = girl[1]
            ws.cell(row=start_row + 3, column=start_column + 2).value = girl[2]
            start_row += 1

        wb.save(filename=new_folder_path+"\\"+"SectionList.xlsx")
