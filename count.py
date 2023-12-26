import pandas as pd

def count_status(header_name, rows):
    header = [(header_name, 'Name'), (header_name, '○'), (header_name, '欠'),(header_name,'遅・早')]
    result = {(header_name, 'Name'): [], (header_name, '○'): [], (header_name, '欠'): [],(header_name,'遅・早'):[]}
    name_totals = {}

    for _, row in rows.iloc[3:].iterrows():
        name = row.iloc[3]  # 4行目から最初の要素を取得

        if name not in name_totals:
            name_totals[name] = {'○': 0, '欠': 0,'遅・早':0}
        print(name)
        result[(header_name, 'Name')].append(name)
        o_count = sum(1 for s in row[1:] if isinstance(s, str) and '○' in s.split('\n'))
        result[(header_name, '○')].append(o_count)
        name_totals[name]['○'] += o_count
        欠_count = sum(1 for s in row[1:] if isinstance(s, str) and '欠' in s.split('\n'))
        result[(header_name, '欠')].append(欠_count)
        name_totals[name]['欠'] += 欠_count
        t_count = sum(1 for s in row[1:] if isinstance(s, str) and ('遅刻' in s or '早退' in s.split('\n')))
        result[(header_name, '遅・早')].append(t_count)
        name_totals[name]['遅・早'] += t_count

    # Add total counts to the right of the corresponding sheet's result
    result[(header_name, 'Name')].extend([f'Total_{name}' for name in name_totals.keys()])
    result[(header_name, '○')].extend(['' for _ in name_totals.keys()])
    result[(header_name, '欠')].extend(['' for _ in name_totals.keys()])
    result[(header_name, '遅・早')].extend(['' for _ in name_totals.keys()])

    for name, totals in name_totals.items():
        result[(header_name, '○')][result[(header_name, 'Name')].index(f'Total_{name}')] = totals['○']
        result[(header_name, '欠')][result[(header_name, 'Name')].index(f'Total_{name}')] = totals['欠']
        result[(header_name, '遅・早')][result[(header_name, 'Name')].index(f'Total_{name}')] = totals['遅・早']

    return pd.DataFrame(result)

def read_excel_and_count(file_path):
    
    excel_data = pd.read_excel(file_path, sheet_name=None, header=0)
    
    # シートごとの結果を保存するリストを作成
    sheet_results = []
    
    for sheet_name, sheet_data in excel_data.items():
        # シートごとの結果をDataFrameに変換
        result_df = count_status(sheet_name, sheet_data)
        
        # シートごとの結果をリストに追加
        sheet_results.append(result_df)
       
        
    # シートごとの結果を1つのDataFrameに結合
    final_result_df = pd.concat(sheet_results, axis=1)
    #final_result_df.columns = final_result_df.columns.droplevel(level=1)
    #print(final_result_df)
    
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
