import os
import classify_method as classify
from flask import Flask, render_template, request
from school_sorting_section import SchoolSortingSection as ScSoSe
from datetime import datetime

app=Flask(__name__)

@app.route("/", methods=['GET'])
def index():
    return render_template("index.html")


@app.route("/success", methods=['POST'])
def success():
    number_of_students = request.form["number_of_students"]
    input_directory = request.form["path"]
    sections = request.form["section_list"]
    
    if not os.path.exists(input_directory):
        print("File does not exist!")
        return False, "File does not exist!"

    sss = ScSoSe()
    try:
        boys_, girls_, section_list_ = sss.get_directory_data(input_directory)
    except Exception as e:
        print(e)
        return False, e

    number_of_students_per_section = number_of_students

    sorted_boys = sss.sorting(boys_)
    sorted_girls = sss.sorting(girls_)

    section_list = sections

    new_folder_path_ = classify.execute_folder(input_directory)

    len_boys = len(sorted_boys)
    len_girls = len(sorted_girls)

    boys_per_section_, girls_per_section_ = sss.ratio(len_boys, len_girls, number_of_students_per_section)

    classify.classify(section_list, boys_per_section_, girls_per_section_, sorted_boys, sorted_girls, new_folder_path_)
    # return True, new_folder_path_
    return render_template("success.html")


# # Trial Run
# # cell_average = 'AB40'
# # cell_math = 'AB27'
# # cell_english = 'AB26'
# # cell_science = 'AB28'


st = datetime.now()
# automated_classification_of_students("C:\\Edward")
et = datetime.now()
print("EXECUTION TIME: ", et - st)

if __name__ == '__main__':
    app.debug=True
    app.run(port=3000)
