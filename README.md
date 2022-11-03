# Python Categorizer

[![GitHub license](https://badgen.net/github/license/Naereen/Strapdown.js)](https://github.com/Naereen/StrapDown.js/blob/master/LICENSE)
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)
[![build](https://img.shields.io/appveyor/build/gruntjs/grunt)](https://pub.dev/packages/flutter_bounce#-analysis-tab-)

- This project is a scripting project which is made to automate the process of categorising different sets of feedbacks we receive from people.
- This solves the basic needs and in no way a very advanced version, which can work like Machine Learning project. 
- This project will keep on improving with given more data in respect of keywords and categories.

## Motivation

This project was a side project which was made to made the life of operation team easier, who were dealing with more than `800 feedbacks` per day, and had to manually read all of it, and categorise, which is a tiring process.

## Pre-requisites

Please be assured of the following things installed before you try running the project

- Language Dependency
    - [x] Python3.x

- Packages Dependencies
    - [x] xlrd
    - [x] pandas
    - [x] openpyxl
    - [x] xlsxwriter

- Make sure your sheet has a column **English**, or if you do want to look out for some other column, then change the code in the `main.py` file by replacing the word of the column name with **English**.

## Run Project

- To install the packages in python, we use `pip/pip3`. Make sure it is installed and then you just use the below command to install the packages

```python
pip3 install <package-name>
```

- To run the project simply get inside the directory via terminal and run the following command

```python
python3 main.py
```

## Links

- [Pandas](https://pandas.pydata.org/docs/)
- [Install Python on Mac](https://www.youtube.com/watch?v=M323OL6K5vs)
- [Install Python on Windows](https://youtu.be/8cAEH1i_5s0)
- [XLXSWrite Official Documentation](https://xlsxwriter.readthedocs.io/worksheet.html)