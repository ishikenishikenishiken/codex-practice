import pandas as pd
import os
import glob
import argparse

# コマンドライン引数セットアップ
def parse_args():
    parser = argparse.ArgumentParser(description='複数Excelから指定項目の値を抽出してCSV出力')
    parser.add_argument('--pattern', type=str, required=True, help='対象となるExcelファイルのパターン（例: ./data/*.xlsx）')
    parser.add_argument('--item', type=str, required=True, help='抽出したい項目名（例: 売上）')
    parser.add_argument('--output', type=str, default='output.csv', help='抽出結果を保存するCSVファイル名')
    return parser.parse_args()

# 項目名に沿って「列方向」または「行方向」のどちらかを自動判定して数値を返す
def extract_values_from_file(filepath, item_name):
    values = []
    try:
        df = pd.read_excel(filepath, header=None, dtype=object)
    except Exception as e:
        print(f'ファイル読込失敗: {filepath}: {e}')
        return values
    # 列方向探索（項目名がどのカラム行にも出現）
    for row_idx in range(df.shape[0]):
        for col_idx in range(df.shape[1]):
            cell = df.iat[row_idx, col_idx]
            if str(cell) == item_name:
                # まずは「カラムヘッダー」扱い（項目が1行に並ぶタイプ）
                # 例: | 名前 | 売上 | 利益 |
                if row_idx == 0:
                    # この列全て抽出（ヘッダ自体は除外）
                    column_values = df.iloc[1:, col_idx].dropna()
                    for val in column_values:
                        if isinstance(val, (int, float)):
                            values.append(val)
                # 逆に「行ヘッダー」扱い（項目がA列固定の下向きに並ぶタイプ）
                if col_idx == 0:
                    row_values = df.iloc[row_idx, 1:].dropna()
                    for val in row_values:
                        if isinstance(val, (int, float)):
                            values.append(val)
    return values

def main():
    args = parse_args()
    file_pattern = args.pattern
    item_name = args.item
    output_file = args.output

    files = glob.glob(file_pattern)
    if not files:
        print("指定パターンに一致するExcelファイルがありません。")
        return
    results = []
    for file in files:
        vals = extract_values_from_file(file, item_name)
        for v in vals:
            results.append({'file': os.path.basename(file), 'item': item_name, 'value': v})

    if results:
        df_out = pd.DataFrame(results)
        df_out.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f'抽出結果を{output_file}に保存しました')
    else:
        print('該当項目の抽出結果がありませんでした')

if __name__ == '__main__':
    main()





