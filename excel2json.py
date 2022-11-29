'''
This is an advanced version of the main.py project. It will have the results 
stored in the JSON file with a format like this

{
    "category_one": [
        {
            "data_1": "",
            "data_2": ""
        }
    ],
    "category_two": [
        {
            "data_1": "",
            "data_2": ""
        }
    ]
    ...
}
'''
import json
import categories
import pandas as pd

data_frame = pd.read_excel('assets/Sample-Sheet.xlsx', engine='openpyxl', dtype=object, header=None)
data_frame = data_frame.fillna("") # Does the job for filling out " " for NaN item (which was a pain)
feedbacks_array_data = data_frame.values.tolist()
feedbacks_array_data = feedbacks_array_data[1:len(feedbacks_array_data)]

key_junk = "Junk"
categories_object = {}

# unique entry logic
for index, _ in enumerate(feedbacks_array_data):
    found = False
    category_item = {}

    # 1 -> Project fails when we have something other than string
    # 2 -> 0th position cosists of the column name, so we don't need it
    # 3 -> Make sure to have the sheet having English column at 4th pos
    current_row_data = feedbacks_array_data[index] 
    feedback = str(current_row_data[4])
    feedback = feedback.strip()

    for word in feedback.split():
        # Skip and move back to line 15 with the next word
        if word.lower() in categories.ignore_words: continue

        for key in categories.categories:
            for keyword in categories.categories[key]:
                if keyword.lower() in word.lower():
                    found = True

                    # Filling the object data and then adding it to the final object
                    category_item["rating"] = current_row_data[2]
                    category_item["feedback"] = current_row_data[4]
                    category_item["method"] = current_row_data[5]
                    category_item["user_country_code"] = current_row_data[11]
                    category_item["user_name"] = current_row_data[15]
                    category_item["user_email"] = current_row_data[16]

                    if key in categories_object:
                        if category_item not in categories_object[key]:
                            categories_object[key].append(category_item)
                    else:
                        categories_object[key] = [category_item]
                    break
            # Main idea is to push the sentence and move to next sentence immediately
            if found: break
        # Breaks word loop
        if found: break

        # Pushing to junk. No word in sentence found in any category
        if not found:
            # Filling the object data and then adding it to the final object
            category_item["rating"] = current_row_data[2]
            category_item["feedback"] = current_row_data[4]
            category_item["method"] = current_row_data[5]
            category_item["user_country_code"] = current_row_data[11]
            category_item["user_name"] = current_row_data[15]
            category_item["user_email"] = current_row_data[16]
            
            if key in categories_object:
                if category_item not in categories_object[key]:
                    categories_object[key].append(category_item)
            else:
                categories_object[key] = [category_item]

# Writing the data to JSON
with open('Final-Data.json', 'w', encoding='utf-8') as f:
    json.dump(categories_object, f, ensure_ascii=False, indent=4)
