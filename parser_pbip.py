import re
import json
import os
import csv

def main():
    listdir = os.listdir()
    r = re.compile(r'.*\.pbip') # Регулярка для поиска *.pbip файлов
    matches = list(filter(r.match, listdir))
    if len(matches) == 0:
        print ('No .pbip found in directory. 1 should be')
        return False
    elif len(matches) > 1:
        print ('More than 2 .pbip files. Only 1 should be in directory')
        return False
    pbip_file_name = matches[0].replace('.pbip', '')
    model_path = pbip_file_name + '.SemanticModel\\model.bim' # У папки такое же название, как у файла, только в конце .SemanticModel
    report_path = pbip_file_name + '.Report\\report.json' # У папки такое же название, как у файла, только в конце .Report
    output_file_name = pbip_file_name + ' СВЯЗИ.csv' # Название выходного .csv файла
    
    if not (model_path and report_path):
        return
    
    columns_dictionary = parser_tables_formulas(model_path)
    elements_sources_parcer(report_path, output_file_name, columns_dictionary)

def elements_sources_parcer(report_path, output_file_name, columns_dictionary):
    enable_print = 0

    headers = ['Номер страницы'                         # page_number
            , 'Страница'                                # page_name
            , 'ID элемента'                             # element_id
            , 'Используемые в элементе атрибуты'        # source_str
            , 'Таблица источник'                        # columns_dictionary[source_str][table_name]
            , 'Power Query создания таблицы источника'  # columns_dictionary[source_str][table_formula]
            , 'Формула атрибута'                        # columns_dictionary[source_str][col_formula]
            , 'Таблица источник атрибута'               # columns_dictionary[source_str][col_table_sources]
            , 'Атрибуты из формулы'                     # columns_dictionary[source_str][col_col_sources]
            ]

    with open(report_path, 'r', encoding='utf-8') as input_file:
        data = json.load(input_file)

    with open(output_file_name, 'w', newline = '', encoding='utf-8') as output_file:
        writer = csv.writer(output_file, delimiter=';')
        writer.writerow(headers)
        for page in data['sections']:
            if 'ordinal' in page:
                page_number = page['ordinal'] + 1
            else:
                page_number = 1
            page_name = page['displayName']
            if enable_print: print (page_name)
            containers = page['visualContainers']
            for container in containers:
                element = json.loads(container['config'])
                element_id = element['name']
                try:
                    sources = element['singleVisual']['prototypeQuery']
                except:
                    continue
                sources_names_dict = {s['Name']:s['Entity'] for s in sources['From']}
                if enable_print: print ('\t' + element['name'])
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
                    if enable_print: print('\t\t' + source_str)

                    row_to_write = ([
                                    page_number
                                  , page_name
                                  , element_id
                                  , source_str
                                  ])

                    if source_str in columns_dictionary:
                        row_to_write.append(columns_dictionary[source_str]['table_name'])
                        row_to_write.append(columns_dictionary[source_str]['table_formula'])
                        row_to_write.append(columns_dictionary[source_str]['col_formula'])
                        row_to_write.append(columns_dictionary[source_str]['col_table_sources'])
                        row_to_write.append(columns_dictionary[source_str]['col_col_sources'])

                    writer.writerow (row_to_write)

                if enable_print: print()
            if enable_print: print ()

def parser_tables_formulas(model_path):
    enable_print = 0

    with open (model_path, 'r', encoding = 'utf-8') as input_file:
        data = json.load(input_file)

    tables_and_columns_dictionary = {}
    for table in (data['model']['tables']):
        if 'isHidden' in table:
            if table['isHidden']:
                continue
        table_name = table['name']

        table_formula = ''
        if 'partitions' in table:
            partitions = table['partitions'][0]
            if 'source' in partitions:
                source = partitions['source']
                if 'expression' in source:
                    expression = source['expression']
                    if source['type'] == 'calculated':
                        table_formula = expression # Если таблица Calculated, то в expression лежит Excel формула
                    else: # Если таблица не calculated, но в expression список действий. На 0 строке 'let', на 1-й 'Источник = ...'
                        table_formula = expression[1].strip().replace('Источник = ', '')
        table_cols = (table['measures'] if 'measures' in table else []) + (table['columns'] if 'columns' in table else [])
        if not table_cols:
            continue
        for col in table_cols:
            if enable_print: print('–'*50)
            col_name = col['name']
            if 'expression' in col:
                col_formula = col['expression'].replace('\n', '')
                # Лямбда-функция для применения регулярки на выражение. Возвращает отсортированный список без дубликатов
                mahinatsii = lambda pattern, text: '\n'.join(sorted(set(re.findall(pattern, text)))).replace('\'', '')
                # ------------------------------------------
                col_table_sources = mahinatsii(r'\'.*?\'', col_formula) # Регулярка для поиска 'Таблица', ищем все таблицы-источники
                col_col_sources = mahinatsii(r'\'[^\']*?\'\[.*?\]', col_formula) # регулярка для поиска 'Таблица'[Значение], ищем все атрибуты-источники
            else:
                col_formula = ''
                col_col_sources = ''
                col_table_sources = ''
            column_dict = { 'table_name':        table_name
                          , 'table_formula':     table_formula.replace('#(lf)', ' ')
                          , 'col_formula':       col_formula
                          , 'col_table_sources': col_table_sources
                          , 'col_col_sources':   col_col_sources
                          }
            tables_and_columns_dictionary[f'{table_name}[{col_name}]'] = column_dict
    # тут кончается проход по таблицам
    if enable_print:
        for key, value in tables_and_columns_dictionary.items():
            print (f'{key}: {value}')
            print()
    return tables_and_columns_dictionary

if __name__ == '__main__':
    main()
