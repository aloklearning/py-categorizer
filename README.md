# Python Categorizer

[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)
[![GitHub license](https://badgen.net/github/license/Naereen/Strapdown.js)](https://github.com/Naereen/StrapDown.js/blob/master/LICENSE)
[![build](https://img.shields.io/appveyor/build/gruntjs/grunt)](https://pub.dev/packages/flutter_bounce#-analysis-tab-)

- This project is a scripting project which is made to automate the process of categorising different sets of feedbacks we receive from people.
- This solves the basic needs and in no way a very advanced version, which can work like Machine Learning project. 
- This project will keep on improving with given more data in respect of keywords and categories.

## Motivation

This project was a side project which was made to made the life of operation team easier, who were dealing `800+ feedbacks` per day, and had to manually read all of it, and categorise, which is indeed a very tiring process.

## Pre-requisites

Please be assured of the following things installed before you try running the project

- Language Dependency
    - [x] Python3.x
    - [x] Pyinstaller (To make the file executable)

- Packages Dependencies
    - [x] xlrd
    - [x] pandas
    - [x] openpyxl
    - [x] xlsxwriter

- Make sure your sheet has a column **English**, or if you do want to look out for some other column, then change the code in the `main.py` file by replacing the word of the column name with **English**. 

## To make the project executable

- In order to make your main file executable, please install the package `pyinstaller`. There would be an absolute chances, that even after the command `pip3 install pyinstaller`, you will be seeing a problem as `pyinstaller command not found`. No worries, in order make things work, just do the following:

```python
# If you don't have pyinstaller installed, then run the pyinstaller installation command, else skip to second
pip3 install pyinstaller
python3 -m PyInstaller --onefile main.py
```

- This will generate an executable file for the **MacOS**, and I think it can work on **Windows** as well, but **not tested**
- Just go to the `dist` directory, and double click the executable file, and it will run as usual

> Please keep on running the above command with each update in the `main.py` file.

## Run Project

When you will the run the project, irrespective of the below cases, you will be asked the below questions. Please provide the right details, failing which, the program will exit it abruptly.

```terminal
Please enter the full path of the file along with an xlsx extension: <your_input>
Please enter the full path where you want to save the result sheet: <your_input>
```

### Normal Run Via Main.Py
- To install the packages in python, we use `pip/pip3`. Make sure it is installed and then you just use the below command to install the packages

```python
pip3 install <package-name>
```

- To run the project simply get inside the directory via terminal and run the following command

- If your expected file is not there in the same directory as `main.py`, then please get the full path of the file using `pwd` for **MacOS/Linux** terminal. For **Windows** please use `cd`.
- Run the command by providing the full path of the file, ending with the file name

```python
python3 main.py
```

### Run Via Executable File

- Double click on the provided executable file, and provide the details required to perform the operation

## Output

### Running Normally python3 main.py
- On succesfull completion of the task, you will have the file named **Final-Sheet.xlsx** stored in the provided path
- You will also see set of interactive messages coming in your terminal like below:

```
Work in progress. Have some coffee ☕
Processing 129 feedbacks...
Process success ✅
We've successfully categorised 129 feedbacks from the provided sheet 🙂
Please check the Final-Sheet.xlsx file in <your_provided_path> directory 📁
```

### Running Executable File
- This file will take few seconds to run, as it installs everything at your disposal, and help you get the results

> This output will come in your terminal/console, when the machine was not able to find any words which was matching a category, hence, the category was not added. This is irrespetive of which path you choose to run the project

```
We were unable to add these categories due to other categories priority precedence: 
Generic Negative
Supply Chain
```

## Future Prospects

Should you require to add more categories, for the machine to give a more robust output. You just need go to your `main.py` file, and do the following:
- If new words need to be added to an existing category:
    - Go to the `categories` word, and inside the existing category, add another word like below:
    
    ```python
    categories = {
        'existing_category': [
            'your_new_word', 
            'some_existing_words'
            ...
        ]
        ...
    }
    ```

- If a new category with new words has to be added:
    - It is something like above only, just add the category with word on top, or may be down. 
    > **Please Note:** The sequence does matter, so if you're category is important, and it's words has to be checked first, add it on top, otherwise, at the bottom

    ```python
    categories = {
        'your_new_category': [
            'your_new_word', 
            'your_new_word_2'
            ...
        ],
        'existing_category': [
            'your_new_word', 
            'some_existing_words'
            ...
        ]
        ...
    }
    ```

Run the project, if everything goes well, you will be able to see the results in a form of final sheet in your directory. Enjoy!


## Links

- [Pandas](https://pandas.pydata.org/docs/)
- [Install Python on Mac](https://www.youtube.com/watch?v=M323OL6K5vs)
- [Install Python on Windows](https://youtu.be/8cAEH1i_5s0)
- [XLXSWrite Official Documentation](https://xlsxwriter.readthedocs.io/worksheet.html)