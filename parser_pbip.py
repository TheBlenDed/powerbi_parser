import pip._internal as pip

def import_lib(name):
    try:
        return __import__(name) # пытаемся импортировать
    except ImportError:
        pip.main(['install', name]) # ставим библиотеку если её нет
    return __import__(name) # возвращаем библиотеку

re = import_lib('re')
json = import_lib('json')
os = import_lib('os')
pandas = import_lib('pandas')
openpyxl = import_lib('openpyxl')

def main():
    current_directory = os.path.dirname(os.path.abspath(__file__))
    listdir = os.listdir(current_directory)
    r = re.compile(r'.*\.pbip') # Регулярка для поиска *.pbip файлов
    matches = list(filter(r.match, listdir))
    if len(matches) == 0:
        print ('No .pbip found in directory. 1 should be')
        return
    elif len(matches) > 1:
        print ('More than 2 .pbip files. Only 1 should be in directory')
        return
    pbip_file_name = matches[0].replace('.pbip', '')
    model_path = current_directory + '\\' + pbip_file_name + '.SemanticModel\\model.bim' # У папки такое же название, как у файла, только в конце .SemanticModel
    report_path = current_directory + '\\' + pbip_file_name + '.Report\\report.json' # У папки такое же название, как у файла, только в конце .Report
    output_file_name = pbip_file_name + ' СВЯЗИ.xlsx' # Название выходного .csv файла
    
    if not (model_path and report_path):
        return
    try:
        columns_dictionary = parser_tables_formulas(model_path)
        print ('model.bim parsed successfully')
    except:
        print ('Error parsing model.bim')
        return
    try:
        output_table = elements_sources_parcer(report_path, columns_dictionary)
        print ('report.json parsed successfully')
    except:
        print ('Error parsing report.json')
        return
    try:
        output_table.to_excel(current_directory + '\\' + output_file_name
                            , sheet_name = 'Источники'
                            , index = False)
        print (f'{output_file_name} saved successfully')
    except:
        print (f'Unable to save {output_file_name}')
        return

def parser_tables_formulas(model_path):

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
                        if type(table_formula) == list:
                            table_formula = '\n'.join(table_formula) # Если она в несколько строк, объединяем в одну
                    else: # Если таблица не calculated, то в expression список действий. На 0 строке 'let', на 1-й 'Источник = ...'
                        table_formula = expression[1].strip().replace('Источник = ', '')
        table_cols = (table['measures'] if 'measures' in table else []) + (table['columns'] if 'columns' in table else [])
        if not table_cols:
            continue
        for col in table_cols:
            col_name = col['name']
            if 'expression' in col:
                col_formula = col['expression']
                if type(col_formula) == list:
                    col_formula = '\n'.join(col_formula) # Формула может быть записана построчно и являться list. На этот случай джойним в одну строку через \n
                source_matches = re.findall(r'(\'[^\']*?\')?(\[.*?\])', col_formula)
                col_col_sources = []
                col_table_sources = []
                for match in source_matches:
                    match = list(match)
                    if not match[0]:
                        match[0] = table_name
                    match[0] = match[0].replace('\'', '')
                    col_col_sources.append(''.join(match))
                    col_table_sources.append(match[0])
                col_col_sources = '\n'.join(sorted(set(col_col_sources)))
                col_table_sources = '\n'.join(sorted(set(col_table_sources)))
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
    return tables_and_columns_dictionary

def elements_sources_parcer(report_path, columns_dictionary): # returns DataFrame

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
    sort_order = [headers[0]
                , headers[2]
                , headers[3]]
    
    df = [] # мы сначала всё сохраним как двумерный список, в конце переделаем в Dataframe
    with open(report_path, 'r', encoding='utf-8') as input_file:
        data = json.load(input_file)

    for page in data['sections']:
        if 'ordinal' in page:
            page_number = page['ordinal'] + 1
        else:
            page_number = 1
        page_name = page['displayName']
        containers = page['visualContainers']
        for container in containers:
            element = json.loads(container['config'])
            element_id = element['name']
            try:
                sources = element['singleVisual']['prototypeQuery']
            except:
                continue
            sources_names_dict = {s['Name']:s['Entity'] for s in sources['From']}
            for source in sources['Select']:
                dir = source
                while 'Source' not in dir:
                    dir = dir[list(dir.keys())[0]]
                source_table = sources_names_dict[dir['Source']]

                dir = source
                while 'Property' not in dir:
                    dir = dir[list(dir.keys())[0]]
                source_table_field = dir['Property'].replace('+virtual', '')

                source_str = f'{source_table}[{source_table_field}]'
                row_to_write = ([page_number
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

                df.append(row_to_write)
    df = pandas.DataFrame(df, columns = headers)
    return df.sort_values(sort_order)

if __name__ == '__main__':
    main()
    input("Press any key to exit")