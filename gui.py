# -*- coding: utf-8 -*-
"""
Created on Wed Oct  2 13:37:27 2019

@author: Steven
"""

from flask import Flask, render_template, flash, request
from wtforms import Form, TextField, TextAreaField, validators, StringField, SubmitField
import classify_method as classify
import school_sorting_section as sss

# App config.
DEBUG = True
app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = '7d441f27d441f27567d441f2b6176a'

class ReusableForm(Form):
    name = TextField('Name:', validators=[validators.required()])
    
    @app.route("/", methods=['GET', 'POST'])
    def hello():
        form = ReusableForm(request.form)
        
        print (form.errors)
        if request.method == 'POST':
            path=request.form['path']
            studentPerSection=request.form['studentPerSection']
            ave=request.form['ave']
            math=request.form['math']
            eng=request.form['eng']
            sci=request.form['sci']
           
        
        if form.validate():
        # Save the comment here.
            boys_, girls_, section_list_ = sss.get_directory_data(path)

            print(len(boys_))
            print(len(girls_))

            sorted_boys = sss.sorting(boys_)
            sorted_girls = sss.sorting(girls_)

            print(sorted_boys)
            print(sorted_girls)

            section_list_manual = ["Amity", "Benevolence", "Bravery", "Charity", "Courage"]
            new_folder_path_ = classify.execute_folder(path)
            print(len(sorted_boys))
            print(len(sorted_girls))
            len_boys = len(sorted_boys)
            len_girls = len(sorted_girls)
            boys_per_section_, girls_per_section_ = sss.ratio(len_boys, len_girls, studentPerSection)

            print("boys_per_section_", boys_per_section_)
            print("girls_per_section_", girls_per_section_)

            classify.classify(section_list_manual, boys_per_section_, girls_per_section_, sorted_boys, sorted_girls, new_folder_path_)
        else:
            flash('Error: All the form fields are required. ')
        
        return render_template('hello.html', form=form)

if __name__ == "__main__":
    app.run()