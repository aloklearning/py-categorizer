import sys
import time
import xlsxwriter
import pandas as pd

dataframe = pd.read_excel('assets/Sample-Sheet.xlsx')

categories = { 
    'Random': ['random'],
    'Help': ['impact', 'better']
}
ignore_words = ['and', 'or', 'not', 'which', 'to', 'a', 'hence', 'is']

# Business Logic
final_results = {}

print("Work in progress. Have some coffee", "\N{hot beverage}")
for index, sentence in enumerate(dataframe['English']):
    for word in sentence.split():
        # Skip and move back to line 15 with the next word
        if word.lower() in ignore_words: continue

        for key in categories:
            for keyword in categories[key]: 
                '''
                This is for checking substring. For example impacted could be ignored
                Keyword has to be in present tense which can be found in any grammatical word
                For example impact (keyword) can be found in impacted (word from excel) and not
                the other way round. 
                '''
                if keyword in word.lower():
                    if key in final_results: final_results[key].append(sentence)
                    else: final_results[key] = [sentence]

# Final Process
workbook = xlsxwriter.Workbook('Final-Sheet.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

for key in final_results:
    worksheet.write(row, col, key)
    for item in final_results[key]:
        worksheet.write(row+1, col, item)
        row += 1
    row = 0
    col += 1

print("Process succss", "\N{check mark}")
print("Please check the Final-Sheet.xlsx file in your same directory", "\N{file folder}")
workbook.close()

