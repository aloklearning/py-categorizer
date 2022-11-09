import xlsxwriter
import pandas as pd

user_input_file = input("Please enter the full path of the file along with an xlsx extension: ")
save_file_path = input("Please enter the full path where you want to save the result sheet: ")
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
        'email', 
        'document'
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
    'Onboarding': [
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

ignore_words = [
    'and', 
    'or', 
    'not', 
    'which', 
    'to', 
    'a',
    'an', 
    'hence', 
    'is', 
    'was', 
    'it', 
    'with', 
    'for',
    'of',
    'it',
    'have'
]

# Business Logic
keyJunk = 'Junk'
final_results = {}

print("Work in progress. Have some coffee", "\N{hot beverage}")
print(f"Processing {len(dataframe['English'])} feedbacks...")
for feedback in dataframe['English']:
    found = False
    feedback = feedback.strip()

    for word in feedback.split():
        # Skip and move back to line 15 with the next word
        if word.lower() in ignore_words: continue

        for key in categories:
            for keyword in categories[key]:
                '''
                For substring match. A word in a sentence can be of any type.
                Key word is in the present, so it has to be in the word. For
                example: Slow (substring) should be there in Slowly
                '''
                if keyword in word.lower():
                    found = True
                    if key in final_results:
                        final_results[key].append(feedback)
                    else:
                        final_results[key] = [feedback]
                    break
            # Main idea is to push the sentence and move to next sentence immediately
            if found: break
        # Breaks word loop
        if found: break

    # Pushing to junk. No word in sentence found in any category
    if not found:
        if keyJunk in final_results:
            final_results[keyJunk].append(feedback)
        else:
            final_results[keyJunk] = [feedback]

processed_count = 0                  
for key in final_results:
    processed_count += len(final_results[key])

missing_categories = []
for category_key in categories:
    if category_key not in final_results:
         if category_key not in missing_categories: 
            missing_categories.append(category_key)

# Final Process
workbook = xlsxwriter.Workbook(f'{save_file_path}/Final-Sheet.xlsx')
worksheet = workbook.add_worksheet()

text_wrap_format = workbook.add_format({'text_wrap': True})
worksheet.set_column('A:Z', 50, text_wrap_format)

row = 0
col = 0

for key in final_results:
    worksheet.write(row, col, key)
    for item in final_results[key]:
        worksheet.write(row+1, col, item)
        row += 1
    row = 0
    col += 1

print("Process success ✅")
print(f"We've successfully categorised {processed_count} feedbacks from the provided sheet 🙂")

if len(missing_categories) > 1:
    missing_categories = '\n'.join(missing_categories)
    print(f'\nWe were unable to add these categories due to other categories priority precedence: \n{missing_categories}\n')

print(f"Please check the Final-Sheet.xlsx file in {save_file_path} directory 📁")
workbook.close()

