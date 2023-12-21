import pandas as pd

def count_status(header_name, rows):
    header = {(header_name, 'Name'), (header_name, '○'), (header_name, '欠')}
    result = {(header_name, 'Name'): [], (header_name, '○'): [], (header_name, '欠'): []}
    for _, row in rows.iterrows():
        name = row[3]

        result[(header_name, 'Name')].append(name)
        result[(header_name, '○')].append(sum(1 for s in row[1:] if isinstance(s, str) and '○' in s.split('\n')))
        result[(header_name, '欠')].append(sum(1 for s in row[1:] if isinstance(s, str) and '欠' in s.split('\n')))
    return pd.DataFrame(result)

def read_excel_and_count(file_path):
    
    excel_data = pd.read_excel(file_path, sheet_name=None, header=None)

    
    # シートごとの結果を保存するリストを作成
    sheet_results = []
    
    for sheet_name, sheet_data in excel_data.items():
        # シートごとの結果をDataFrameに変換
        result_df = count_status(sheet_name, sheet_data)
        
        # シート名を結果に追加
        # result_df['Sheet'] = sheet_name
        
        # シートごとの結果をリストに追加
        sheet_results.append(result_df)
        print(sheet_results)
        
    # シートごとの結果を1つのDataFrameに結合
    final_result_df = pd.concat(sheet_results, axis=1)
    final_result_df.columns = final_result_df.columns.droplevel(level=1)
    
    # 結果を1つのCSVファイルに書き込み（UTF-8 エンコーディング）
    result_csv_path = 'c:/Users/lull-/Desktop/test.csv'
    final_result_df.to_csv(result_csv_path, index=False, encoding='utf-8-sig')

    
    # 結果CSVファイルのパスを返す
    return result_csv_path

# Excelファイルのパス
excel_file_path = 'c:/Users/lull-/Desktop/スタサポ/新しいフォルダー/手動出席簿/2018年度8月通常授業出席簿.xls'

# 関数呼び出し
result_path = read_excel_and_count(excel_file_path)

# 結果CSVファイルのパスを表示
print(f'各シートごとの結果はこちらに保存されました: {result_path}')
