import pandas as pd
import json

excel_file = '账单.xlsx'
cleaned_json_file = '新账单（AI）.json'
rules_json_file = 'rules.json'

try:
    df = pd.read_excel(excel_file)
    columns_to_drop = ['Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10']
    existing_columns_to_drop = [col for col in columns_to_drop if col in df.columns]
    
    if existing_columns_to_drop:
        df_columns_dropped = df.drop(columns=existing_columns_to_drop)
    else:
        df_columns_dropped = df

    header_index = df_columns_dropped[df_columns_dropped['微信支付账单明细'] == '交易时间'].index[0]
    column_names = df_columns_dropped.iloc[header_index].values
    data_start_index = header_index + 1
    cleaned_df = df_columns_dropped.iloc[data_start_index:].copy()
    cleaned_df.columns = column_names
    cleaned_df.reset_index(drop=True, inplace=True)
    
    with open(rules_json_file, 'r', encoding='utf-8') as f:
        rules_data = json.load(f)
        
    cleaned_data = cleaned_df.to_dict(orient='records')
    
    final_data = {
        "rules": rules_data,
        "data": cleaned_data
    }

    with open(cleaned_json_file, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, indent=4, ensure_ascii=False)
        
except FileNotFoundError as e:
    print(f"错误：文件未找到 - {e.filename}。请确保文件名和路径正确。")
except Exception as e:
    print(f"发生错误: {e}")