import streamlit as st
import pandas as pd
from io import BytesIO

st.title('複数Excelから項目値自動抽出アプリ')

st.markdown('''
1. 複数のエクセルファイル（.xlsx）をアップロード
2. 項目名（例: 売上）を入力
3. "抽出" ボタンをクリック
4. 結果が一覧表示＆ダウンロードボタン
''')

uploaded_files = st.file_uploader(
    'Excelファイルを複数選択してアップロードしてください',
    type=['xlsx'],
    accept_multiple_files=True)

item_name = st.text_input('抽出したい項目名（例: 売上）')

extract_btn = st.button('抽出')


def extract_values_from_file(file, item_name):
    values = []
    try:
        df = pd.read_excel(file, header=None, dtype=object)
    except Exception as e:
        st.warning(f'ファイル読込失敗: {e}')
        return values
    for row_idx in range(df.shape[0]):
        for col_idx in range(df.shape[1]):
            cell = df.iat[row_idx, col_idx]
            if str(cell) == item_name:
                # 列方向ヘッダ
                if row_idx == 0:
                    column_values = df.iloc[1:, col_idx].dropna()
                    for val in column_values:
                        if isinstance(val, (int, float)):
                            values.append(val)
                # 行方向ヘッダ
                if col_idx == 0:
                    row_values = df.iloc[row_idx, 1:].dropna()
                    for val in row_values:
                        if isinstance(val, (int, float)):
                            values.append(val)
    return values

if extract_btn and uploaded_files and item_name:
    results = []
    for uploaded_file in uploaded_files:
        vals = extract_values_from_file(uploaded_file, item_name)
        for v in vals:
            results.append({'file': uploaded_file.name, 'item': item_name, 'value': v})
    if results:
        df_out = pd.DataFrame(results)
        st.success('抽出に成功しました！')
        st.dataframe(df_out)
        # CSVダウンロード
        csv_bytes = df_out.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button('抽出結果ダウンロード(CSV)',
            data=csv_bytes,
            file_name='extract_result.csv',
            mime='text/csv')
    else:
        st.warning('該当するデータが見つかりませんでした')

