import os
import re
import pandas as pd
import io
import unicodedata

# 全角 → 半角 の変換テーブル
ZEN2HAN_TABLE = str.maketrans({
    '，': ',', '．': '.', '　': ' ', '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9'
})

def sanitize_filename(name):
    """ファイル名として使えない文字を除去する"""
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def extract_content(file_path):
    """ファイルから<HeadLine>と<CsvData>タグの中身を正規表現で抽出"""
    encodings_to_try = ['sjis', 'cp932', 'utf-16', 'utf-8']
    for encoding in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content = f.read()
            headline_match = re.search(r'<HeadLine>(.*?)</HeadLine>', content, re.DOTALL)
            csv_data_match = re.search(r'<CsvData>(.*?)</CsvData>', content, re.DOTALL)
            if headline_match and csv_data_match:
                print(f"✓ '{encoding}' でデータを抽出しました。")
                return headline_match.group(1).strip(), csv_data_match.group(1).strip()
        except Exception:
            continue
    return None, None

def format_name_for_display(name):
    """氏名の文字数に応じて全角スペースを挿入して整形する"""
    name = str(name).strip().replace('　', '').replace(' ', '')
    ln = len(name)
    if ln == 2: return f"{name[0]}　　　{name[1]}"
    if ln == 3: return f"{name[:2]}　　{name[2:]}"
    if ln == 4: return f"{name[:2]}　{name[2:]}"
    return name

def write_df_to_excel_with_formatting(df, excel_path, sheet_name='Sheet1'):
    """DataFrameをExcelファイルとして書き出し、詳細な書式を設定する"""
    try:
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            workbook  = writer.book
            worksheet = workbook.add_worksheet(sheet_name)
            
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            base_num_props = {'num_format': '#,##0', 'align': 'right'}

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for row_idx, row_data in df.iterrows():
                row_num = row_idx + 1
                bg_props = {'bg_color': '#F0F0F0'} if row_num % 2 == 0 else {}
                
                for col_idx, col_name in enumerate(df.columns):
                    cell_value = row_data[col_name]
                    
                    if col_name == '政党名／候補者名':
                        rich_bold_format = workbook.add_format({'bold': True, **bg_props})
                        rich_normal_format = workbook.add_format({**bg_props})
                        
                        cand_name = format_name_for_display(row_data.get('政党名／候補者名', ''))
                        party = row_data.get('党派名', '')
                        status = row_data.get('身分', '')
                        
                        segments = [rich_bold_format, cand_name]
                        if pd.notna(party) and party: segments.extend([rich_normal_format, f" {party}"])
                        if pd.notna(status) and status: segments.extend([rich_normal_format, f" {status}"])
                        
                        worksheet.write_rich_string(row_num, col_idx, *segments)
                    
                    elif col_name in df.select_dtypes(include=['number']).columns:
                        num_format = workbook.add_format({**base_num_props, **bg_props})
                        if pd.notna(cell_value):
                            worksheet.write_number(row_num, col_idx, cell_value, num_format)
                        else:
                            worksheet.write(row_num, col_idx, '', workbook.add_format(bg_props))
                    
                    else:
                        text_format = workbook.add_format(bg_props)
                        worksheet.write(row_num, col_idx, cell_value if pd.notna(cell_value) else '', text_format)

            for idx, col in enumerate(df.columns):
                header_len = sum(1 + (unicodedata.east_asian_width(c) in 'FWA') for c in str(col))
                max_len = df[col].astype(str).map(lambda x: sum(1 + (unicodedata.east_asian_width(c) in 'FWA') for c in x)).max()
                worksheet.set_column(idx, idx, max(header_len, max_len if pd.notna(max_len) else 0) + 2)
            
        print(f"✓ {os.path.basename(excel_path)} を出力しました。")

    except Exception as e:
        print(f"エラー: {excel_path} の書き込み中にエラーが発生しました: {e}")

def process_xml_file(xml_path):
    print(f"\n--- 処理開始: {xml_path} ---")

    headline, csv_text = extract_content(xml_path)
    if not headline or not csv_text:
        print(f"警告: <HeadLine> または <CsvData> が見つかりませんでした。スキップします。")
        return

    csv_text_han = csv_text.translate(ZEN2HAN_TABLE)
    
    try:
        # ★★★【修正点】空白行を削除するためのデータ読み込みロジック ★★★
        f = io.StringIO(csv_text_han)
        reader = csv.reader(f)
        all_rows = list(reader)

        header, header_index = None, -1
        for i, row in enumerate(all_rows):
            if row and row[0].strip() == '順位':
                header = [c.strip() for c in row]
                header_index = i
                break

        if not header:
            print("警告: ヘッダー行が見つかりませんでした。スキップします。")
            return

        candidate_data = []
        for row in all_rows[header_index + 1:]:
            # 2列目(index 1)が4桁以上の数字で始まる行のみを候補者データとして抽出
            if len(row) > 1 and row[1] and re.match(r'^\s*\d{4,}', str(row[1]).strip()):
                candidate_data.append(row)
        
        if not candidate_data:
            print("警告: 処理対象の候補者データが見つかりませんでした。")
            return

        df_full = pd.DataFrame(candidate_data, columns=header)
        # ★★★ ここまで ★★★

        numeric_cols = [c for c in df_full.columns if c not in ['順位', '政党コード／人物番号', '政党名／候補者名', '当落マーク', '党派コード', '党派名', '身分', '候補者氏名', '特定枠']]
        for col in ['順位'] + numeric_cols:
            if col in df_full.columns:
                df_full[col] = pd.to_numeric(df_full[col], errors='coerce').fillna(0).astype(int)

        base_filename = sanitize_filename(headline)
        
        full_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}.xlsx")
        write_df_to_excel_with_formatting(df_full.fillna(''), full_excel_path, sheet_name=headline[:31])

        if '比例代表候補者得票順' in headline:
            df_tou = df_full[df_full['当落マーク'].str.strip().fillna('').astype(bool)].copy()
            df_tou.drop(columns=[c for c in ['政党コード／人物番号', '当落マーク', '党派コード'] if c in df_tou.columns], inplace=True)
            tou_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}当.xlsx")
            write_df_to_excel_with_formatting(df_tou.fillna(''), tou_excel_path, sheet_name='当選者リスト')

            df_raku = df_full[
                (df_full['当落マーク'].str.strip().fillna('') == '') &
                (df_full['党派コード'].str.strip().fillna('').astype(bool))
            ].copy()
            cols_to_keep = ['順位', '政党名／候補者名', '党派名', '身分', '合　計']
            df_raku = df_raku[[c for c in cols_to_keep if c in df_raku.columns]]
            raku_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}落.xlsx")
            write_df_to_excel_with_formatting(df_raku.fillna(''), raku_excel_path, sheet_name='落選者リスト')

    except Exception as e:
        print(f"エラー: ファイル '{xml_path}' の処理中に予期せぬエラーが発生しました: {e}")

# --- メイン処理 ---
if __name__ == '__main__':
    current_folder = '.'
    for filename in os.listdir(current_folder):
        if filename.lower().endswith('.xml'):
            process_xml_file(os.path.join(current_folder, filename))
    print("\nすべての処理が完了しました。")