import re
import json
import csv

json_from_tabular_editor_input_file = r'json_from_tabular_editor.json'
tables_formulas_output_file = r'tables_formulas.csv'

report_json_input_file = r'report.json'
element_sources_output_file = r'element_sources.csv'

def main():
    #elements_sources_parcer(report_json_input_file, element_sources_output_file)
    parser_tables_formulas(json_from_tabular_editor_input_file, tables_formulas_output_file)

def elements_sources_parcer(input_file_name, output_file_name):

    headers = ['Номер страницы'   # page_number
            , 'Страница'          # page_name
            , 'ID элемента'       # element_id
            , 'Таблица источник'] # source_str

    with open(folder_path + "\\" + folder_path.split("\\")[-1]+".Report" + "\\" + input_file_name, 'r', encoding='utf-8') as input_file:
        data = json.load(input_file)

    with open(folder_path + "\\" + output_file_name, 'w', newline = '', encoding='utf-8') as output_file:
        writer = csv.writer(output_file, delimiter=';')
        writer.writerow(headers)
        for page in data['sections']:
            if 'ordinal' in page:
                page_number = page['ordinal'] + 1
            else:
                page_number = 1
            page_name = page['displayName']
            print (page_name)
            containers = page['visualContainers']
            for container in containers:
                element = json.loads(container['config'])
                element_id = element['name']
                try:
                    sources = element['singleVisual']['prototypeQuery']
                except:
                    continue
                sources_names_dict = {s['Name']:s['Entity'] for s in sources['From']}
                print ('\t' + element['name'])
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
                    writer.writerow([page_number
                                    , page_name
                                    , element_id
                                    , source_str])
                print()
            print ()

def get_sql(table):
    try:
        pquery = table["partitions"][0]["source"]["expression"]
        pquery = re.sub(r'#\(.*?\)', ' ', pquery)
        pquery = fr'{pquery}'
        if 'Query=' in pquery:
            sql_ = re.search("Query=\"\"(.*?)\"\"", pquery).group(1)
            print ("sql_")

        else:
            return pquery
    except:
        return ""

def parser_tables_formulas(input_file_name, output_file_name):

    headers = ['Field Name'                            # col_name
            , 'Используемые в элементе атрибуты'       # f"{table_name}[{col_name}]"
            , 'Таблица источник'                       # table_name
            , 'Power Query создания таблицы источника' # table_formula
            , 'Формула атрибута'                       # col_formula
            , 'Таблица источник атрибута'              # col_table_sources
            , 'Атрибуты из формулы'                    # col_col_sources
            ]
    
    with open (folder_path + "\\" + input_file_name, 'r', encoding = 'utf-8') as input_file:
        data = json.load(input_file)

    with open (folder_path + "\\" + output_file_name, 'w', newline='', encoding = 'utf-8') as output_file:
         writer = csv.writer(output_file, delimiter=";")
         writer.writerow(headers)
#         for table in (data["tables"] + data["calculatedtables"]):
#             table_name = table["name"]
#             table_formula = get_sql(table)
#             table_cols = (table["measures"] if "measures" in table else []) + (table["columns"] if "columns" in table else [])
#             if not table_cols:
#                 continue
#             for col in table_cols:
#                 col_name = col["name"]
#                 if "expression" in col:
#                     mahinatsii = lambda pattern, text: "\n".join(sorted(set(re.findall(pattern, text)))).replace("'", "")
#                     col_formula = col["expression"].replace("\n", "")
#                     col_col_sources = mahinatsii(r"'[^']*?'\[.*?\]", col_formula)
#                     col_table_sources = mahinatsii(r"'.*?'", col_formula)
#                 else:
#                     col_formula = ""
#                     col_col_sources = ""
#                     col_table_sources = ""
#                 row = [col_name
#                     , f"{table_name}[{col_name}]"
#                     , table_name
#                     , table_formula
#                     , col_formula
#                     , col_table_sources
#                     , col_col_sources
#                     ]
#                 #print (row)
#                 writer.writerow(row)

if __name__ == "__main__":
    main()