import json
import re
import csv

headers = ['Page'
           , 'Element_id'
           , 'Source']

with open('report.json', 'r', encoding='utf-8') as input_file:
    data = json.load(input_file)

with open('elements_relations.csv', 'w', newline = '') as output_file:
    writer = csv.writer(output_file, delimiter=';')
    writer.writerow(headers)
    for page in data['sections']:
        page_name = page['displayName']
        print (page_name)
        containers = page['visualContainers']
        for container in containers:
            element = json.loads(container['config'])
            element_name = element['name']
            element_string = json.dumps(element, ensure_ascii = False, indent = 0)
            sources = re.findall(r'\"queryRef\": \"(.*)\"', element_string)
            if not sources:
                continue
            print ('\t' + element['name'])
            for page, source in enumerate(sources):
                # if ('(' or ')') in source:
                #     sources[i] = re.search(r'\(.*\)', source).group(0)[1:-1]
                # sources[i] = sources[i].replace('.', '[', 1)+']'
                print('\t\t' + source)
                writer.writerow([page_name, element_name, source])
            print()
        print ()