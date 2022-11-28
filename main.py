import xlsxwriter
import categories
import pandas as pd

# Global Variables
keyJunk = 'Junk'

def writing_result_excel(final_results, missing_categories, processed_count, save_file_path):
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

    if len(missing_categories) > 0:
        missing_categories = '\n'.join(missing_categories)
        print(f'\nWe were unable to add these categories due to other categories priority precedence: \n{missing_categories}\n')

    print(f"Please check the Final-Sheet.xlsx file in {save_file_path} directory 📁")
    workbook.close()

def compute_missing_categories(final_results, processed_count, save_file_path):
    missing_categories = []
    for category_key in categories.categories:
        if category_key not in final_results:
            if category_key not in missing_categories: 
                missing_categories.append(category_key)
    
    writing_result_excel(final_results, missing_categories, processed_count, save_file_path)

def count_processed_data(final_results, save_file_path):
    processed_count = 0                  
    for key in final_results:
        processed_count += len(final_results[key])

    compute_missing_categories(final_results, processed_count, save_file_path)

def unique_entry_category(dataframe, save_file_path):
    unique_final_results = {}
    for feedback in dataframe['English']:
        found = False

        # Project fails when we have something other than string 
        feedback = str(feedback)
        feedback = feedback.strip()

        for word in feedback.split():
            # Skip and move back to line 15 with the next word
            if word.lower() in categories.ignore_words: continue

            for key in categories.categories:
                for keyword in categories.categories[key]:
                    '''
                    keyword.lower() and word.lower() removes the hassle of comparing issues

                    For substring match. A word in a sentence can be of any type.
                    Key word is in the present, so it has to be in the word. For
                    example: Slow (substring) should be there in Slowly
                    '''
                    if keyword.lower() in word.lower():
                        found = True
                        if key in unique_final_results:
                            unique_final_results[key].append(feedback)
                        else:
                            unique_final_results[key] = [feedback]
                        break
                # Main idea is to push the sentence and move to next sentence immediately
                if found: break
            # Breaks word loop
            if found: break

        # Pushing to junk. No word in sentence found in any category
        if not found:
            if keyJunk in unique_final_results:
                unique_final_results[keyJunk].append(feedback)
            else:
                unique_final_results[keyJunk] = [feedback]
    
    count_processed_data(unique_final_results, save_file_path)

def duplicate_entry_category(dataframe, save_file_path):
    duplicate_final_results = {}

    for feedback in dataframe['English']:
        found = False

        feedback = str(feedback)
        feedback = feedback.strip()

        for word in feedback.split():
            if word.lower() in categories.ignore_words: continue

            for key in categories.categories:
                for keyword in categories.categories[key]:
                    if keyword.lower() in word.lower():
                        found = True
                        if key in duplicate_final_results:
                            # Removing duplicate value entry in a category
                            if feedback not in duplicate_final_results[key]:
                                duplicate_final_results[key].append(feedback)
                        else:
                            duplicate_final_results[key] = [feedback]

        if not found:
            if keyJunk in duplicate_final_results:
                duplicate_final_results[keyJunk].append(feedback)
            else:
                duplicate_final_results[keyJunk] = [feedback]

    count_processed_data(duplicate_final_results, save_file_path)

def process_input_data(trigger):
    user_input_file = input("Please enter the full path of the file along with an xlsx extension: ")
    save_file_path = input("Please enter the full path where you want to save the result sheet: ")
    dataframe = pd.read_excel(user_input_file)

    print("Work in progress. Have some coffee", "\N{hot beverage}")
    print(f"Processing {len(dataframe['English'])} feedbacks...")

    if trigger == "trigger_unique":
        unique_entry_category(dataframe, save_file_path)
    else:
        duplicate_entry_category(dataframe, save_file_path)

# Program starts here
user_input_option = input("Do you need the resultant sheet with duplicate records solution (Y/N): ")

# We immediately want to exit the system, without asking any further question, if the answer is other than Y or N
if user_input_option.lower().strip() == "y":
    process_input_data("trigger_duplicate")
elif user_input_option.lower().strip() == "n":
    process_input_data("trigger_unique")
else:
    print("You have entered an invalid input, please enter either Y or N to get the results")

