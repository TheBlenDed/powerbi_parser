import re
import json
import csv

def get_sql(table):
    try:
        pquery = table["partitions"][0]["source"]["expression"]
        pquery = re.sub(r'#\(.*?\)', ' ', pquery)
        pquery = fr'{pquery}'
        #try:
        #    return re.search("Query=\"(.*?)\"", pquery).group(1)
        #except:
        return pquery
    except:
        return ""

def main():
    with open ('Торговый эквайринг.json', 'r', encoding = 'utf-8') as input_file:
        data = json.load(input_file)

    headers = ['Field Name'
                , 'Table Name'
                , 'Table Formula'
                , 'Table + Field Names'
                , 'Field Formula'
                , 'Field Table Sources'
                , 'Field Field Sources'
                ]

    with open ('output.csv', 'w', newline='') as output_file:
        writer = csv.writer(output_file, delimiter=";")
        writer.writerow(headers)
        for table in (data["tables"] + data["calculatedtables"]):
            table_name = table["name"]
            table_formula = get_sql(table)
            table_cols = (table["measures"] if "measures" in table else []) + (table["columns"] if "columns" in table else [])
            if not table_cols:
                continue
            for col in table_cols:
                col_name = col["name"]
                if "expression" in col:
                    mahinatsii = lambda pattern, text: "\n".join(sorted(set(re.findall(pattern, text)))).replace("'", "")
                    col_formula = col["expression"].replace("\n", "")
                    col_col_sources = mahinatsii(r"'[^']*?'\[.*?\]", col_formula)
                    col_table_sources = mahinatsii(r"'.*?'", col_formula)
                else:
                    col_formula = ""
                    col_col_sources = ""
                    col_table_sources = ""
                writer.writerow([col_name
                                , table_name
                                , table_formula
                                , f"{table_name}[{col_name}]"
                                , col_formula
                                , col_table_sources
                                , col_col_sources
                                ])


if __name__ == "__main__":
    main()