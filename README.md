# Python Categorizer

[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)
[![GitHub license](https://badgen.net/github/license/Naereen/Strapdown.js)](https://github.com/Naereen/StrapDown.js/blob/master/LICENSE)
[![build](https://img.shields.io/appveyor/build/gruntjs/grunt)](https://pub.dev/packages/flutter_bounce#-analysis-tab-)

- This project is a scripting project which is made to automate the process of categorising different sets of feedbacks we receive from people.
- This solves the basic needs and in no way a very advanced version, which can work like Machine Learning project. 
- This project will keep on improving with given more data in respect of keywords and categories.

## Motivation

This project was a side project which was made to made the life of operation team easier, who were dealing with more than `800 feedbacks` per day, and had to manually read all of it, and categorise, which is indeed a very tiring process.

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

- If your expected file is not there in the same directory as `main.py`, then please get the full path of the file using `pwd` for **MacOS/Linux** terminal. For **Windows** please use `cd`.
- Run the command by providing the full path of the file, ending with the file name

```python
python3 main.py <full-path>/<file_name>
```

- If you have the expected file in the same directory, then simply add your file name, and it will work with the following command:

```python
python3 main.py <file-name>
```

## Output

- On succesfull completion of the task, you will have the file named **Final-Sheet.xlsx** stored in the same directory.
- You will also see set of interactive messages coming in your terminal like below:

```
Work in progress. Have some coffee ☕
Process success ✓
Please check the Final-Sheet.xlsx file in your same directory 📁
```

## Links

- [Pandas](https://pandas.pydata.org/docs/)
- [Install Python on Mac](https://www.youtube.com/watch?v=M323OL6K5vs)
- [Install Python on Windows](https://youtu.be/8cAEH1i_5s0)
- [XLXSWrite Official Documentation](https://xlsxwriter.readthedocs.io/worksheet.html)