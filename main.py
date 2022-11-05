import sys
import xlsxwriter
import pandas as pd

user_input_file = sys.argv[1]
dataframe = pd.read_excel(user_input_file)

categories = { 
    'Bugs': [
        'terminal', 
        'error', 
        'connection', 
        'pos', 
        'e-pos', 
        'card', 
        'payment', 
        'website'
    ],
    'Customer Ops': [
        'translation', 
        'english', 
        'emails', 
        'documentation'
    ],
    'Customer Support': [
        'contract', 
        'communication', 
        'inactivity', 
        'service', 
        'phone', 
        'email', 
        'chat', 
        'customer', 
        'subscription', 
        'telphone',
        'assistance'
    ],
    'Generic Negative': ['bad'],
    'Generic Positive': [
        'good',
        'great',
        'impressive',
        'reliable',
        'satisfied',
        'outstanding',
        'well',
        'top',
        'service',
        'quick',
        'excellent',
        'perfect',
        'quick',
        'love',
        'comfortable',
        'like',
        'accessible'
    ],
    'New Feature Request': [
        'rental',
        'possibility',
        'offer',
        'possible', 
        'elimiate',
        'installement', 
        'fractional',
        'increase',
        'accept',
        'payment',
        'welcome'
    ],
    'onboarding': [
        'password',
        'terminal', 
        'interview', 
        'sales',
        'representative',
        'change'
    ],
    'Pricing/Billing': [
        'transaction',
        'settlement', 
        'high',
        'fee',
        'business', 
        'rate',
        'summary',
        'statement',
        'report',
        'price',
        'service',
        'percentage',
        'financial',
        'money',
        'commission', 
        'billing',
        'delay',
        'payment',
        'receipt',
        'b-online'
    ],
    'Usability Issue': [
        'interface',
        'slow',
        'android',
        'system',
        'interest',
        'troubleshoot',
        'miss',
        'poor', 
        'pop-up'
    ],
    'Supply Chain': ['order']
}

ignore_words = ['and', 'or', 'not', 'which', 'to', 'a', 'hence', 'is', 'was', 'it', 'with', 'for']

# Business Logic
final_results = {}

print("Work in progress. Have some coffee", "\N{hot beverage}")
for sentence in dataframe['English']:
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
                else:
                    if 'Junk' in final_results: final_results['Junk'].append(sentence)
                    else: final_results['Junk'] = [sentence]


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

print("Process success", "\N{check mark}")
print("Please check the Final-Sheet.xlsx file in your same directory", "\N{file folder}")
workbook.close()

