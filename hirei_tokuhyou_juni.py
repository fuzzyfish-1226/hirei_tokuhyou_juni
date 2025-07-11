import os
import re
import pandas as pd
import io
import unicodedata
import csv
from typing import List, Tuple, Optional, Dict

# スクリプト全体の目的を記載
"""
選挙結果が記録された特定のXML形式のファイルを自動で処理し、
整形されたExcelファイルを出力するためのスクリプト。

主な機能:
- カレントディレクトリ内の全XMLファイルを自動検出して処理。
- 複数の文字エンコーディングに対応し、堅牢にデータを抽出。
- 全角文字を半角に変換。
- 候補者データのみを抽出し、不要な行（政党集計、空白行）を削除。
- 3種類のExcelファイルを出力:
  1. 全データリスト
  2. 当選者リスト
  3. 落選者リスト
- Excel出力時に詳細な書式設定を適用:
  - ヘッダー: 太字、中央揃え、罫線
  - データ行: 1行おきの網掛け
  - 数値: 3桁区切りカンマ付き、右寄せ
  - 氏名: 文字数に応じたスペース挿入、太字（党派・身分は通常フォント）、MSゴシック
"""

# 全角 → 半角 の変換テーブル
ZEN2HAN_TABLE = str.maketrans({
    '，': ',', '．': '.', '　': ' ', '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9'
})

def sanitize_filename(name: str) -> str:
    """ファイル名として使えない文字を除去または置換する。

    Windowsのファイル名として使用できない文字 (`\\/*?:"<>|`) を
    アンダースコア `_` に置き換えます。

    Args:
        name (str): 元のファイル名候補の文字列。

    Returns:
        str: 安全なファイル名に変換された文字列。
    """
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def extract_content(file_path: str) -> Optional[Tuple[str, str]]:
    """
    ファイルから<HeadLine>と<CsvData>タグの中身を正規表現で抽出する。

    複数の文字エンコーディング（utf-8, sjisなど）を順番に試行し、
    ファイル内容を読み込めたものから見出しとCSVデータを抽出します。
    どちらかが見つからない場合はNoneを返します。

    Args:
        file_path (str): 読み込むXMLファイルのパス。

    Returns:
        Optional[Tuple[str, str]]: (見出し, CSVデータ) のタプル。見つからない場合はNone。
    """
    encodings_to_try: List[str] = ['utf-8', 'sjis', 'cp932', 'utf-16']
    for encoding in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content: str = f.read()
            
            headline_match = re.search(r'<InHeadLine>(.*?)</InHeadLine>', content, re.DOTALL)
            if not headline_match: headline_match = re.search(r'<HeadLine>(.*?)</HeadLine>', content, re.DOTALL)
            if not headline_match: headline_match = re.search(r'<DeliveryHeadline1>(.*?)</DeliveryHeadline1>', content, re.DOTALL)

            csv_data_match = re.search(r'<CsvData>(.*?)</CsvData>', content, re.DOTALL)
            if not csv_data_match: csv_data_match = re.search(r'<Sentence>(.*?)</Sentence>', content, re.DOTALL)

            if headline_match and csv_data_match:
                print(f"✓ '{encoding}' でデータを抽出しました。")
                headline: str = headline_match.group(1).strip()
                csv_data: str = re.sub(r'</?InData>.*', '', csv_data_match.group(1).strip(), flags=re.DOTALL)
                return headline, csv_data
        except Exception:
            continue
    return None, None

def format_name_for_display(name: str) -> str:
    """
    氏名の文字数に応じて全角スペースを挿入し、見栄えを整える。

    - 2文字: 姓と名の間に全角スペース3つ
    - 3文字: 姓2・名1とみなし、間に全角スペース2つ
    - 4文字: 姓2・名2とみなし、間に全角スペース1つ
    - 5文字以上: スペースなし

    Args:
        name (str): 整形前の氏名。

    Returns:
        str: 整形後の氏名。
    """
    name = str(name).strip().replace('　', '').replace(' ', '')
    ln: int = len(name)
    if ln == 2: return f"{name[0]}　　　{name[1]}"
    if ln == 3: return f"{name[:2]}　　{name[2:]}"
    if ln == 4: return f"{name[:2]}　{name[2:]}"
    return name

def write_df_to_excel_with_formatting(df: pd.DataFrame, excel_path: str, sheet_name: str = 'Sheet1', combine_name_cols: bool = True) -> None:
    """
    DataFrameをExcelファイルとして書き出し、詳細な書式（網掛け、フォント、数値形式など）を設定する。

    Args:
        df (pd.DataFrame): 書き出すデータフレーム。
        excel_path (str): 出力するExcelファイルのパス。
        sheet_name (str, optional): Excelシート名。デフォルトは'Sheet1'。
        combine_name_cols (bool, optional): 氏名・党派・身分を結合するかどうかのフラグ。デフォルトはTrue。
    """
    try:
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            workbook  = writer.book
            worksheet = workbook.add_worksheet(sheet_name)
            
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            base_num_props: Dict[str, str] = {'num_format': '#,##0', 'align': 'right'}
            
            cols_to_hide: List[str] = ['党派名', '身分'] if combine_name_cols else []
            display_cols: List[str] = [c for c in df.columns if c not in cols_to_hide]
            
            for col_num, value in enumerate(display_cols):
                worksheet.write(0, col_num, value, header_format)

            for row_idx, row_data in df.iterrows():
                row_num: int = row_idx + 1
                bg_props: Dict[str, str] = {'bg_color': '#F0F0F0'} if row_num % 2 == 0 else {}
                
                for display_col_idx, col_name in enumerate(display_cols):
                    cell_value = row_data.get(col_name)
                    
                    if combine_name_cols and col_name == '政党名／候補者名':
                        rich_bold_format = workbook.add_format({'bold': True, 'font_name': 'MS Gothic', **bg_props})
                        rich_normal_format = workbook.add_format({'font_name': 'MS Gothic', **bg_props})
                        cand_name = format_name_for_display(row_data.get('政党名／候補者名', ''))
                        party = row_data.get('党派名', '')
                        status = row_data.get('身分', '')
                        segments: list = [rich_bold_format, cand_name]
                        if pd.notna(party) and party: segments.extend([rich_normal_format, f" {party}"])
                        if pd.notna(status) and status: segments.extend([rich_normal_format, f" {status}"])
                        worksheet.write_rich_string(row_num, display_col_idx, *segments)
                    
                    elif col_name in df.select_dtypes(include=['number']).columns:
                        num_format = workbook.add_format({**base_num_props, **bg_props})
                        if pd.notna(cell_value):
                            worksheet.write_number(row_num, display_col_idx, cell_value, num_format)
                    
                    else:
                        text_format = workbook.add_format(bg_props)
                        worksheet.write(row_num, display_col_idx, cell_value if pd.notna(cell_value) else '', text_format)

            for idx, col in enumerate(display_cols):
                header_len = sum(1 + (unicodedata.east_asian_width(c) in 'FWA') for c in str(col))
                max_len = df[col].astype(str).map(lambda x: sum(1 + (unicodedata.east_asian_width(c) in 'FWA') for c in x)).max()
                worksheet.set_column(idx, idx, max(header_len, max_len if pd.notna(max_len) else 0) + 3)
            
        print(f"✓ {os.path.basename(excel_path)} を出力しました。")

    except Exception as e:
        print(f"エラー: {excel_path} の書き込み中にエラーが発生しました: {e}")

def process_xml_file(xml_path: str) -> None:
    """指定されたXMLファイルを処理し、必要なExcelファイルを出力するメイン関数"""
    print(f"\n--- 処理開始: {xml_path} ---")
    headline, csv_text = extract_content(xml_path)
    if not headline or not csv_text:
        print(f"警告: 見出しまたはCSVデータが見つかりませんでした。スキップします。")
        return

    csv_text_han: str = csv_text.translate(ZEN2HAN_TABLE)
    
    try:
        all_rows = list(csv.reader(io.StringIO(csv_text_han)))
        header, header_index = next(((row, i) for i, row in enumerate(all_rows) if row and row[0].strip() == '順位'), (None, -1))
        if not header:
            print("警告: ヘッダー行が見つかりませんでした。スキップします。")
            return
        header = [c.strip() for c in header]

        name_col_idx: int = header.index('政党名／候補者名')
        candidate_data: List[List[str]] = [
            row for row in all_rows[header_index + 1:]
            if len(row) > name_col_idx and str(row[name_col_idx]).strip() and re.match(r'^\s*\d{4,}', str(row[1]).strip())
        ]
        
        if not candidate_data:
            print("警告: 処理対象の候補者データが見つかりませんでした。")
            return

        df_full = pd.DataFrame(candidate_data, columns=header)

        numeric_cols: List[str] = [c for c in df_full.columns if c not in ['順位', '政党コード／人物番号', '政党名／候補者名', '当落マーク', '党派コード', '党派名', '身分', '候補者氏名', '特定枠']]
        for col in ['順位'] + numeric_cols:
            if col in df_full.columns:
                df_full[col] = pd.to_numeric(df_full[col], errors='coerce').fillna(0).astype(int)

        base_filename: str = sanitize_filename(headline)
        
        full_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}.xlsx")
        write_df_to_excel_with_formatting(df_full.fillna(''), full_excel_path, sheet_name=headline[:31], combine_name_cols=True)

        if '比例代表候補者得票順' in headline:
            df_tou = df_full[df_full['当落マーク'].str.strip().fillna('').astype(bool)].copy().reset_index(drop=True)
            df_tou.drop(columns=[c for c in ['政党コード／人物番号', '当落マーク', '党派コード'] if c in df_tou.columns], inplace=True)
            tou_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}当.xlsx")
            write_df_to_excel_with_formatting(df_tou.fillna(''), tou_excel_path, sheet_name='当選者リスト', combine_name_cols=True)

            # ★★★【修正箇所】落選者リストのロジック ★★★
            df_raku_filtered = df_full[
                (df_full['当落マーク'].str.strip().fillna('') == '') &
                (df_full['党派コード'].str.strip().fillna('') != '')
            ].copy()
            
            cols_to_keep: List[str] = ['順位', '政党名／候補者名', '合 計']
            # df_raku_filtered から必要な列を抽出し、インデックスをリセット
            df_raku_output = df_raku_filtered[[c for c in cols_to_keep if c in df_raku_filtered.columns]].reset_index(drop=True)
            
            raku_excel_path = os.path.join(os.path.dirname(xml_path), f"{base_filename}落.xlsx")
            write_df_to_excel_with_formatting(df_raku_output.fillna(''), raku_excel_path, sheet_name='落選者リスト', combine_name_cols=False)
            # ★★★ ここまで ★★★

    except Exception as e:
        print(f"エラー: ファイル '{xml_path}' の処理中に予期せぬエラーが発生しました: {e}")

# --- メイン処理 ---
if __name__ == '__main__':
    current_folder: str = '.'
    for filename in os.listdir(current_folder):
        if filename.lower().endswith('.xml'):
            process_xml_file(os.path.join(current_folder, filename))
    print("\nすべての処理が完了しました。")