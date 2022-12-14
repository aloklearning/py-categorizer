'''
This is an advanced version of the main.py project. It will have the results 
stored in the JSON file with a format like this

{
    "category_one": {
        "promoters": [
            {
                "data_1": "",
                "data_2": "",
                ...
            }
        ],
        "passive": [
            {
                "data_1": "",
                "data_2": "",
                ...
            }
        ],
        "detractors": [
            {
                "data_1": "",
                "data_2": "",
                ...
            }
        ]
    },
    "category_two": {
         "promoters": [
            {
                "data_1": "",
                "data_2": "",
                ...
            }
        ]
        ...
    }
    ...
}
'''
import json
import categories
import pandas as pd

key_junk = "Junk"

def write_data_to_json(final_data_object, processed_count, save_file_path):
    with open(f'{save_file_path}/Final-Data.json', 'w', encoding='utf-8') as f:
        json.dump(final_data_object, f, ensure_ascii=False, indent=4)
    
    print("Process success ✅")
    print(f"We've successfully categorised {processed_count} feedbacks from the provided sheet 🙂")

    print("Here is the summary of the items we have per each category:\n")
    for index, categoryKey in enumerate(final_data_object):
        total_count = 0

        for type_values in final_data_object[categoryKey].values():
            total_count += len(type_values)
        print(f"{index+1}. {categoryKey} has {total_count} items", end="\n")

    print(f"\nPlease check the Final-Data.json file in {save_file_path} directory 📁")

def count_processed_data(final_data, save_file_path):
    processed_count = 0                  
    for key in final_data:
        for values in final_data[key].values():
            processed_count += len(values)

    write_data_to_json(final_data, processed_count, save_file_path)
    
def processing_final_object(key_name, category_type, category_item, current_row_data, categories_object):
    # Filling the object data and then adding it to the final object
    category_item["rating"] = current_row_data[2]
    category_item["feedback"] = current_row_data[4].replace("\n", " ")
    category_item["method"] = current_row_data[5]
    category_item["user_country_code"] = current_row_data[11]
    category_item["user_name"] = current_row_data[15]
    category_item["user_email"] = current_row_data[16]

    if category_item["rating"] >= 9 and category_item["rating"] <= 10: 
        category_type = "promoters"
    elif category_item["rating"] >= 7 and category_item["rating"] <= 8: 
        category_type = "passive"
    else:
        category_type = "detractors"
    
    if key_name in categories_object:
        if category_type in categories_object[key_name]:
            if category_item not in categories_object[key_name][category_type]:
                categories_object[key_name][category_type].append(category_item)
        else:
            categories_object[key_name][category_type] = [category_item]
    else:
        categories_object[key_name] = {category_type: [category_item]}

def process_unique_entry(array_data, save_file_path):
    categories_object = {}

    for index, _ in enumerate(array_data):
        found = False
        category_type = ""
        category_item = {}

        # 1 -> Project fails when we have something other than string
        # 2 -> 0th position cosists of the column name, so we don't need it
        # 3 -> Make sure to have the sheet having English column at 4th pos
        current_row_data = array_data[index] 
        feedback = str(current_row_data[4])
        feedback = feedback.strip()

        for word in feedback.split():
            # Skip and move back to line 15 with the next word
            if word.lower() in categories.ignore_words: continue

            for key in categories.categories:
                for keyword in categories.categories[key]:
                    if keyword.lower() in word.lower():
                        found = True
                        processing_final_object(key, category_type, category_item, current_row_data, categories_object)
                        
                        break
                # Main idea is to push the sentence and move to next sentence immediately
                if found: break
            # Breaks word loop
            if found: break

        # Pushing to junk. No word in sentence found in any category
        if not found:
            processing_final_object(key_junk, category_type, category_item, current_row_data, categories_object)

    count_processed_data(categories_object, save_file_path)

def process_input_data():
    user_input_file = input("Please enter the full path of the file along with an xlsx extension: ")
    save_file_path = input("Please enter the full path where you want to save the result sheet: ")
    data_frame = pd.read_excel(user_input_file, engine='openpyxl', dtype=object, header=None)
    data_frame = data_frame.fillna("") # Does the job for filling out " " for NaN item (which was a pain)
    feedbacks_array_data = data_frame.values.tolist()
    feedbacks_array_data = feedbacks_array_data[1:len(feedbacks_array_data)]

    print("Work in progress. Have some coffee", "\N{hot beverage}")
    print(f"Processing {len(feedbacks_array_data)} feedbacks...")

    process_unique_entry(feedbacks_array_data, save_file_path)

# Process starts here
process_input_data()