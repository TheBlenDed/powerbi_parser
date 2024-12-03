import json
import csv

headers = ['Page'
           , 'Element_id'
           , 'Source']

report_path = r'c:\Users\d.kovalev\Desktop\БЦК\Торговый эквайринг\Торговый эквайринг.Report\report.json'
save_file_path = r'c:\Users\d.kovalev\Desktop\БЦК\Торговый эквайринг\elements_relations_utf-8.csv'

with open(report_path, 'r', encoding='utf-8') as input_file:
    data = json.load(input_file)

with open(save_file_path, 'w', newline = '', encoding='utf-8') as output_file:
    writer = csv.writer(output_file, delimiter=';')
    writer.writerow(headers)
    for page in data['sections']:
        page_name = page['displayName']
        print (page_name)
        containers = page['visualContainers']
        page_name_written = False

        for container in containers:
            element = json.loads(container['config'])
            element_id = element['name']
            element_string = json.dumps(element, ensure_ascii = False, indent = 0)
            try:
                sources = element['singleVisual']['prototypeQuery']
            except:
                continue
            sources_names_dict = {s['Name']:s['Entity'] for s in sources['From']}
            print ('\t' + element['name'])
            element_id_written = False

            for source in sources['Select']:
                dir = source
                while 'Source' not in dir:
                    dir = dir[list(dir.keys())[0]]
                source_table = sources_names_dict[dir['Source']]

                dir = source
                while 'Property' not in dir:
                    dir = dir[list(dir.keys())[0]]
                source_table_field = dir['Property']

                source_str = f'{source_table}[{source_table_field}]'
                print('\t\t' + source_str)
                writer.writerow([page_name if not page_name_written else ""
                                , element_id if not element_id_written else ""
                                , source_str])
                element_id_written = True
                page_name_written = True

            print()
        print ()