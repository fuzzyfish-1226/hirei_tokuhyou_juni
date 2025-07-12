# ==============================================================================
# 必要なライブラリをインポート
# ==============================================================================
import os
import re
import pandas as pd
import io
import unicodedata
import csv
from typing import List, Tuple, Optional, Dict

# ==============================================================================
# スクリプト全体の目的
# ==============================================================================
"""
このスクリプトは、選挙結果が記録された特定のXML形式のファイルを自動で処理し、
見やすく整形されたExcelファイルを出力することを目的としています。

■ 主な機能:
1. スクリプトと同じフォルダにある全てのXMLファイル(.xml)を自動で探し出して処理します。
2. ファイルがUTF-8やShift_JISなど、異なる文字コードで保存されていても、自動で対応して読み込みます。
3. ファイル内の全角の数字や記号を、Excelで扱いやすい半角に変換します。
4. データの中から候補者の情報だけを抽出し、集計行などの不要なデータは取り除きます。
5. 以下の3種類のExcelファイルを出力します。
   - 全候補者リスト: 全ての候補者のデータを含みます。
   - 当選者リスト: 当選した候補者のみのデータです。
   - 落選者リスト: 落選した候補者のみのデータです。（特定の選挙結果でのみ生成）
6. Excelファイルを見やすくするために、以下のような書式設定を自動で行います。
   - 表の見出し（ヘッダー）: 太字、中央揃え、枠線付きにします。
   - データ行: 1行おきに色を付けて（網掛け）、縞模様にします。
   - 数値データ: 3桁ごとにカンマを付けて、右揃えにします。
   - 氏名: 文字数に応じてスペースを挿入し、氏名部分のみを太字にして見やすくします。
"""

# ==============================================================================
# グローバル定数・設定
# ==============================================================================

# 全角文字を半角文字に変換するための対応表（辞書）を作成します。
# str.maketrans() を使うと、高速に置換処理ができます。
ZEN2HAN_TABLE = str.maketrans({
    '，': ',', '．': '.', '　': ' ', '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9'
})

# ==============================================================================
# ヘルパー関数群
# ==============================================================================

def sanitize_filename(name: str) -> str:
    """
    ファイル名として使えない特殊文字を安全な文字に置き換えます。

    Windowsのファイル名では `\\ / * ? : " < > |` といった文字は使えません。
    この関数は、これらの文字をすべてアンダースコア `_` に置き換えることで、
    プログラムがエラーを起こすのを防ぎます。

    Args:
        name (str): 元のファイル名候補となる文字列。

    Returns:
        str: 安全なファイル名に変換された文字列。
    """
    safe_name = re.sub(r'[\\/*?:"<>|]', '_', name)
    return safe_name

def extract_content_from_xml(file_path: str) -> Optional[Tuple[str, str]]:
    """
    XMLファイルから見出しとCSV形式のデータ部分を抽出します。

    日本のシステムでは、ファイルの文字コードが複数存在することがあります（例: UTF-8, Shift_JIS）。
    この関数は、いくつかの主要な文字コードを順番に試し、正しく読み込めるまで試行します。
    ファイルから`<HeadLine>`（見出し）と`<CsvData>`（データ本体）のタグで囲まれた
    テキストを抽出します。

    Args:
        file_path (str): 処理対象のXMLファイルのパス。

    Returns:
        Optional[Tuple[str, str]]:
            成功した場合は (見出しの文字列, CSVデータの文字列) のタプルを返します。
            見出しやデータが見つからなかった場合は (None, None) を返します。
    """
    encodings_to_try: List[str] = ['utf-8', 'sjis', 'cp932', 'utf-16']

    for encoding in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content: str = f.read()

            headline_match = re.search(r'<InHeadLine>(.*?)</InHeadLine>', content, re.DOTALL)
            if not headline_match:
                headline_match = re.search(r'<HeadLine>(.*?)</HeadLine>', content, re.DOTALL)
            if not headline_match:
                headline_match = re.search(r'<DeliveryHeadline1>(.*?)</DeliveryHeadline1>', content, re.DOTALL)

            csv_data_match = re.search(r'<CsvData>(.*?)</CsvData>', content, re.DOTALL)
            if not csv_data_match:
                csv_data_match = re.search(r'<Sentence>(.*?)</Sentence>', content, re.DOTALL)

            if headline_match and csv_data_match:
                print(f"✓ ファイルを '{encoding}' で読み込み、データを抽出しました。")
                headline: str = headline_match.group(1).strip()
                csv_data_text: str = csv_data_match.group(1).strip()
                csv_data_cleaned: str = re.sub(r'</?InData>.*', '', csv_data_text, flags=re.DOTALL)
                return headline, csv_data_cleaned

        except Exception as e:
            continue

    return None, None

def format_name_for_display(name: str) -> str:
    """
    氏名の文字数に応じて全角スペースを挿入し、Excelでの見栄えを整えます。

    - 2文字の氏名: 姓と名の間に全角スペースを3つ挿入 (例: "山田　　　花子")
    - 3文字の氏名: 姓2文字・名1文字とみなし、間に2つ挿入 (例: "佐々木　　小次郎")
    - 4文字の氏名: 姓2文字・名2文字とみなし、間に1つ挿入 (例: "徳川　家康")
    - 5文字以上: スペースは挿入しない

    Args:
        name (str): 整形前の氏名文字列。

    Returns:
        str: 見栄えが整えられた氏名文字列。
    """
    name_str = str(name)
    name_cleaned = name_str.strip().replace('　', '').replace(' ', '')
    name_length: int = len(name_cleaned)

    if name_length == 2:
        return f"{name_cleaned[0]}　　　{name_cleaned[1]}"
    if name_length == 3:
        return f"{name_cleaned[:2]}　　{name_cleaned[2:]}"
    if name_length == 4:
        return f"{name_cleaned[:2]}　{name_cleaned[2:]}"
    
    return name_cleaned

def _get_display_width(text: str) -> int:
    """
    文字列の表示上の幅を計算します（半角=1, 全角=2）。

    Args:
        text (str): 幅を計算する文字列。

    Returns:
        int: 計算された表示幅。
    """
    width = 0
    for char in str(text):
        if unicodedata.east_asian_width(char) in 'FWA':
            width += 2
        else:
            width += 1
    return width

# ★★★【ここからが修正箇所】★★★
# write_df_to_excel_with_formatting 関数を修正します。
def write_df_to_excel_with_formatting(df: pd.DataFrame, excel_path: str, sheet_name: str, combine_name_cols: bool) -> None:
    """
    データフレームを、見やすい書式設定を適用してExcelファイルに書き出します。

    Args:
        df (pd.DataFrame): Excelに書き出すデータが格納されたDataFrame。
        excel_path (str): 出力するExcelファイルのパス（例: "output.xlsx"）。
        sheet_name (str): Excelのシート名。
        combine_name_cols (bool): Trueの場合、「氏名」「党派」「身分」を1つのセルに結合して表示します。
    """
    try:
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)

            # --- Excelの書式（フォーマット）を事前に定義 ---
            header_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
            })
            number_format_base = {'num_format': '#,##0', 'align': 'right'}
            
            # --- 表示する列を決定 ---
            cols_to_hide: List[str] = []
            if combine_name_cols:
                cols_to_hide = ['党派名', '身分']

            display_cols: List[str] = []
            for col_name in df.columns:
                if col_name not in cols_to_hide:
                    display_cols.append(col_name)

            # --- 1. ヘッダー行の書き込み ---
            for col_idx, col_name in enumerate(display_cols):
                worksheet.write(0, col_idx, col_name, header_format)

            # --- 2. データ行の書き込み ---
            for row_idx, row_data in df.iterrows():
                excel_row_num = row_idx + 1

                if excel_row_num % 2 == 0:
                    bg_color_prop = {'bg_color': '#F0F0F0'}
                else:
                    bg_color_prop = {}

                for display_col_idx, col_name in enumerate(display_cols):
                    cell_value = row_data.get(col_name)

                    # a) 「政党名／候補者名」列の場合 (リッチテキスト)
                    if combine_name_cols and col_name == '政党名／候補者名':
                        rich_bold_format = workbook.add_format({'bold': True, 'font_name': 'MS Gothic', **bg_color_prop})
                        rich_normal_format = workbook.add_format({'font_name': 'MS Gothic', **bg_color_prop})

                        cand_name = format_name_for_display(row_data.get('政党名／候補者名', ''))
                        party = row_data.get('党派名', '')
                        status = row_data.get('身分', '')
                        
                        segments = [rich_bold_format, cand_name]
                        if pd.notna(party) and party:
                            segments.extend([rich_normal_format, f" {party}"])
                        if pd.notna(status) and status:
                            segments.extend([rich_normal_format, f" {status}"])
                        
                        worksheet.write_rich_string(excel_row_num, display_col_idx, *segments)

                    # b) 「氏名」列の場合 (セル全体を太字)
                    elif col_name == '氏名':
                        # 背景色に加えて 'bold': True を指定した書式を作成
                        bold_text_format = workbook.add_format({'bold': True, **bg_color_prop})
                        worksheet.write(excel_row_num, display_col_idx, cell_value if pd.notna(cell_value) else '', bold_text_format)

                    # c) 数値データを含む列の場合
                    elif col_name in df.select_dtypes(include=['number']).columns:
                        num_format = workbook.add_format({**number_format_base, **bg_color_prop})
                        if pd.notna(cell_value):
                            worksheet.write_number(excel_row_num, display_col_idx, cell_value, num_format)

                    # d) その他の通常のテキストデータの列の場合
                    else:
                        text_format = workbook.add_format(bg_color_prop)
                        worksheet.write(excel_row_num, display_col_idx, cell_value if pd.notna(cell_value) else '', text_format)

            # --- 3. 列幅の自動調整 ---
            for col_idx, col_name in enumerate(display_cols):
                header_width = _get_display_width(col_name)
                
                max_data_width = 0
                for text in df[col_name].astype(str):
                    width = _get_display_width(text)
                    if width > max_data_width:
                        max_data_width = width
                
                final_width = max(header_width, max_data_width) + 3
                worksheet.set_column(col_idx, col_idx, final_width)
            
        print(f"✓ Excelファイル '{os.path.basename(excel_path)}' を出力しました。")

    except Exception as e:
        print(f"エラー: Excelファイル '{excel_path}' の書き込み中にエラーが発生しました: {e}")
# ★★★【ここまでが修正箇所】★★★

# ==============================================================================
# メイン処理関数
# ==============================================================================

def process_xml_file(xml_path: str) -> None:
    """
    単一のXMLファイルを処理し、必要なExcelファイルをすべて出力する一連の流れを実行します。
    
    Args:
        xml_path (str): 処理対象のXMLファイルのパス。
    """
    print(f"\n--- 処理開始: {xml_path} ---")

    headline, csv_text = extract_content_from_xml(xml_path)
    if not headline or not csv_text:
        print(f"警告: '{xml_path}' から見出しまたはCSVデータを抽出できませんでした。このファイルはスキップします。")
        return

    csv_text_hankaku = csv_text.translate(ZEN2HAN_TABLE)
    
    try:
        all_rows = list(csv.reader(io.StringIO(csv_text_hankaku)))

        header_row = None
        header_row_index = -1
        for i, row in enumerate(all_rows):
            if row and row[0].strip() == '順位':
                header_row = [col_name.strip() for col_name in row]
                header_row_index = i
                break
        
        if not header_row:
            print(f"警告: '{xml_path}' 内でヘッダー行（'順位'から始まる行）が見つかりませんでした。スキップします。")
            return

        name_col_idx = header_row.index('政党名／候補者名')
        candidate_data_rows = []
        for row in all_rows[header_row_index + 1:]:
            is_valid_row = len(row) > name_col_idx
            has_name = str(row[name_col_idx]).strip() != ''
            is_candidate_code = re.match(r'^\s*\d{4,}', str(row[1]).strip()) is not None
            
            if is_valid_row and has_name and is_candidate_code:
                candidate_data_rows.append(row)
        
        if not candidate_data_rows:
            print(f"警告: '{xml_path}' 内で処理対象となる候補者データが見つかりませんでした。")
            return

        df_full = pd.DataFrame(candidate_data_rows, columns=header_row)

        cols_to_convert_to_numeric = []
        all_cols_except_text = ['順位', '政党コード／人物番号', '政党名／候補者名', '当落マーク', '党派コード', '党派名', '身分', '候補者氏名', '特定枠']
        for col in df_full.columns:
            if col not in all_cols_except_text:
                cols_to_convert_to_numeric.append(col)
        
        for col in ['順位'] + cols_to_convert_to_numeric:
            if col in df_full.columns:
                series_numeric = pd.to_numeric(df_full[col], errors='coerce')
                series_filled = series_numeric.fillna(0)
                df_full[col] = series_filled.astype(int)

        base_filename = sanitize_filename(headline)
        output_dir = os.path.dirname(xml_path)

        full_excel_path = os.path.join(output_dir, f"{base_filename}.xlsx")
        write_df_to_excel_with_formatting(df_full.fillna(''), full_excel_path, sheet_name=headline[:31], combine_name_cols=True)

        if '比例代表候補者得票順' in headline:
            df_tou = df_full[df_full['当落マーク'].str.strip().fillna('').astype(bool)].copy()
            df_tou.reset_index(drop=True, inplace=True)
            
            cols_to_drop_tou = ['政党コード／人物番号', '当落マーク', '党派コード']
            df_tou.drop(columns=[c for c in cols_to_drop_tou if c in df_tou.columns], inplace=True)
            
            # --- '氏名' 列の挿入（ここのロジックは変更なし） ---
            if '政党名／候補者名' in df_tou.columns:
                name_data_to_copy = df_tou['政党名／候補者名']
                insert_position = 22
                if insert_position > len(df_tou.columns):
                    insert_position = len(df_tou.columns)
                df_tou.insert(loc=insert_position, column='氏名', value=name_data_to_copy)
                print("情報: 当選者リストに '氏名' 列をコピーして追加しました。")
            
            tou_excel_path = os.path.join(output_dir, f"{base_filename}当.xlsx")
            write_df_to_excel_with_formatting(df_tou.fillna(''), tou_excel_path, sheet_name='当選者リスト', combine_name_cols=True)

            # --- 落選者リストの作成と出力 ---
            is_raku = (df_full['当落マーク'].str.strip().fillna('') == '')
            has_party_code = (df_full['党派コード'].str.strip().fillna('') != '')
            df_raku_filtered = df_full[is_raku & has_party_code].copy()

            cols_to_keep_raku = ['順位', '政党名／候補者名', '合 計']
            final_raku_cols = [c for c in cols_to_keep_raku if c in df_raku_filtered.columns]
            df_raku_output = df_raku_filtered[final_raku_cols].reset_index(drop=True)

            raku_excel_path = os.path.join(output_dir, f"{base_filename}落.xlsx")
            write_df_to_excel_with_formatting(df_raku_output.fillna(''), raku_excel_path, sheet_name='落選者リスト', combine_name_cols=False)

    except Exception as e:
        print(f"エラー: ファイル '{xml_path}' の処理中に予期せぬエラーが発生しました: {e}")

# ==============================================================================
# スクリプト実行のエントリーポイント
# ==============================================================================
if __name__ == '__main__':
    current_folder_path = '.'
    all_files_in_folder = os.listdir(current_folder_path)
    
    for filename in all_files_in_folder:
        if filename.lower().endswith('.xml'):
            xml_file_path = os.path.join(current_folder_path, filename)
            process_xml_file(xml_file_path)
            
    print("\nすべてのXMLファイルの処理が完了しました。")