def mamkynePodminky(header, split_options):
    if split_options is not None:
        for split in split_options:
            if header in split and split[header] == 1:
                rt = {k: v for k, v in split.items()}
                return rt

def fix_xlsx_with_excel(input_file, output_file):
    
    import win32com.client
    import os
    import pythoncom
    try: 
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(input_file))
        wb.SaveAs(os.path.abspath(output_file))
        wb.Close()
        excel.Quit()
    except: 
        wb.Close()
        excel.Quit()



def parse_monthly_consumption(file_path, header_mapping=None, splitOptions=None):

    import locale
    import openpyxl
    import pandas as pd
    import uuid
    if header_mapping is None:
        header_mapping = {}

    locale.setlocale(locale.LC_ALL, 'cs_CZ.UTF-8')
    df = pd.read_excel(file_path, header=[0, 1], engine='openpyxl')

    den_col = None
    for col in df.columns:
        if col[1] == 'Den':
            den_col = col
            break

    if den_col is None:
        raise ValueError("Sloupec 'Den' nebyl nalezen ve druhé úrovni hlavičky.")

    summary_row = df[df[den_col] == 'Sumář'].iloc[0]

    summary_row = summary_row.iloc[:-1]
    

    headers = list(set(col[0] for col in df.columns if col[0] != den_col[0]))
    
    result = {}
    
    header_to_key = {}
    for header in headers:
        found = False
        

        for target_header, aliases in header_mapping.items():
            if header in aliases:
                header_to_key[header] = target_header
                found = True
                break
        if not found:
            header_to_key[header] = header

    for header in headers:
        target_header = header_to_key[header]
        header_cols = [col for col in df.columns if col[0] == header]
        
        
        valid_header_cols = [col for col in header_cols if col in summary_row.index]
        if not valid_header_cols:
            continue  

        total = summary_row[valid_header_cols].sum()
        if len(header_cols) > 40:
            total = total / 2
        else:
            if len(valid_header_cols) > 1:
                valid_header_cols = valid_header_cols[:-1]
            total = summary_row[valid_header_cols].sum()

        neco = mamkynePodminky(header, split_options=splitOptions)
        
        
        if(type(neco) == dict):
            for tenant, ratio in neco.items():
                if tenant in result:
                    result[tenant] += total * ratio
                else:
                    result[tenant] = total * ratio
            continue  


        if target_header in result:
            result[target_header] += total
        else:
            result[target_header] = total

    if result:
        last_header = list(result.keys())[-1]
        result[last_header] = result[last_header] / 2
    
    return result



def RunParser():
    import json
    import urllib.request
    contents = urllib.request.urlopen("http://192.168.1.137:8080/getParsingOption/1").read()
    json_data = contents.decode("utf-8")
    parsed = json.loads(json_data)

    remapingOptions = parsed.get("remapingOptions", {})
    splitOptions = []

    for split in parsed.get("splitOptions", []):
        
        split_map = {k: v for k, v in split.items()}
        splitOptions.append(split_map)
    
   
    
    import pandas as pd
    import json
    inputFile = "temp_excel_downloads\Mesicni-data.xlsx"
    outputFile = "temp_excel_downloads\Mesicni-data_fixed.xlsx"
    
    try:
        fix_xlsx_with_excel(inputFile, outputFile)
        totals = parse_monthly_consumption(outputFile, remapingOptions, splitOptions)
        for header, total in totals.items():
            print(f"{header}: {total:.2f}")
        
        return totals.items()
    except Exception as e:
        print(f"Chyba: {e}")
        return f"Chyba: {e}"

