# python3 -m pip install openpyxl
import json
import pandas as pd
from pathlib import Path
import sys

def main():

    filepath = input('Enter file path: ')
    absolute_path = str(Path(filepath).resolve())

    content = None
    try:
        with open(absolute_path, 'r') as f:
            content = json.load(f)
    except:
        pass

    if content == None:
        print('File not found. Try again.')
        sys.exit(1)   

    lists = []
    try_again = True
    while try_again:
        goal_text = 'continuing' if any(lists) else 'exporting ''all lists'
        list_title = input(f'Enter target list title (enter empty for {goal_text}): ')
        try_again = len(list_title) != 0
        if try_again:
            lists.append(list_title)

    print('Processing...')

    no_filter = not any(lists)
    extracted = [extract_list(l) for l in content['items'] if no_filter or l['title'] in lists]
    flat_extracted = [item for sublist in extracted for item in sublist]

    df = pd.DataFrame(flat_extracted)
    df.to_excel('extraction.xlsx')

    print('Tasks extracted to file "extraction.xlsx".')

def extract_list(list_json):
    return [ [list_json['title'], item['title'], item['notes'] if 'notes' in item else '', item['updated']] for item in list_json['items'] ]

if __name__ == "__main__": 
    main()