import xlwings as xw
import pandas as pd
import os
import shutil
from datetime import datetime
import copy

def is_numeric(val):
    if pd.isnull(val): return False
    if isinstance(val, (int, float)): return True
    try:
        float(str(val).strip())
        return True
    except ValueError:
        return False

def is_code_conv(val):
    if pd.isnull(val): return False
    val_str = str(val).strip()
    return "変換" in val_str or "コード" in val_str

def get_mapping_message(item):
    if item.get('x_val') != "固定長":
        return "入力ファイルが固定長でないため検証不可"
    elif item.get('ap_val') != "固定長": 
        return "出力ファイルが可変長でないため検証不可"
    return "入力項目がないため検証不可" # Hoặc logic khác

def evaluate_rule(rule_name, item, data, has_code_conv=False):
    in_seq = str(data.get('in_seq', '')).strip() if pd.notnull(data.get('in_seq')) else ""
    out_seq = str(data.get('out_seq', '')).strip() if pd.notnull(data.get('out_seq')) else ""

    if rule_name == "HAS_IN":
        return "Inputがない" if not in_seq else None
    elif rule_name == "HAS_OUT":
        return "Outputがない" if not out_seq else None
    elif rule_name == "NOT_HAS_IN":
        return "Inputが存在する" if in_seq else None
    elif rule_name == "NOT_HAS_OUT":
        return "Outputが存在する" if out_seq else None
    elif rule_name == "IS_IN_FIXED_LENGTH":
        return "固定長である" if item.get('x_val') != "固定長" else None
    elif rule_name == "IS_IN_NOT_LENGTH":
        return "固定長である" if item.get('x_val') == "固定長" else None
    elif rule_name == "IS_OUT_FIXED_LENGTH":
        return "固定長ではない" if item.get('ap_val') != "固定長" else None
    elif rule_name == "IS_OUT_NOT_FIXED_LENGTH":
        return "固定長ではない" if item.get('ap_val') == "固定長" else None
    elif rule_name == "ALWAYS_NG":
        return "常にNG"
    elif rule_name == "IS_NUM_IN":
        return "Inputが数値型ではない" if not data.get('is_num') else None
    elif rule_name == "IS_TEXT_ZENKAKU_IN":
        return "Inputが文字型(全角)ではない" if not data.get('is_text_zenkaku') else None
    elif rule_name == "IS_NUM_OUT":
        return "Outputが数値型ではない" if not data.get('is_num_out') else None
    
    elif rule_name == "HAS_NEGATIVE_IN":
        in_l = str(data.get('in_minus', '')).upper() if pd.notnull(data.get('in_minus')) else ""

        return "Inputにマイナス項目がない" if "X" not in in_l else None

    elif rule_name == "HAS_NEGATIVE_OUT":
        out_ae = str(data.get('out_minus', '')).upper() if pd.notnull(data.get('out_minus')) else ""

        return "Outputにマイナス項目がない" if "X" not in out_ae else None
    
    elif rule_name == "HAS_DECIMAL_IN":
        in_k = str(data.get('in_dec', '')).upper() if pd.notnull(data.get('in_dec')) else ""
        try:
            ad_val = float(in_k)
        except ValueError:
            ad_val = 0.0
            
        if ad_val <= 0:
            return "inputに小数点がない"
        
        return None
    
    elif rule_name == "HAS_GT0_IN":
        in_l = str(data.get('in_minus', '')).strip() if pd.notnull(data.get('in_minus')) else ""
        return "Inputに小数点がない" if in_l in ["", "0", "-"] else None
    elif rule_name == "HAS_X_OUT":
        out_ad = str(data.get('out_dec', '')).upper() if pd.notnull(data.get('out_dec')) else ""
        return "Outputにマイナス項目がない" if "S" not in out_ad and "-" not in out_ad else None
    elif rule_name == "HAS_DECIMAL_OUT":
        out_ad = str(data.get('out_dec', '')).strip() if pd.notnull(data.get('out_dec')) else ""
       
        
        # 1. Check cột AD xem có > 0 không (ép kiểu float an toàn)
        try:
            ad_val = float(out_ad)
        except ValueError:
            ad_val = 0.0
            
        if ad_val <= 0:
            return "Outputに小数点がない"
        
        return None
    
    elif rule_name == "IS_CODE_CONV_OUT":
        return "コード変換処理ではない" if not is_code_conv(data.get('c_code_conv')) else None
    elif rule_name == "COMMON_CODE_CONV_GRAYOUT":
        return "コード変換処理がない" if not has_code_conv else None
    
    return None

def generate_testcase(input_ids=None, target_date=None, cols_input='', cols_output='', cols_check='', copy_sheets=''):
    # --- CẤU HÌNH ĐƯỜNG DẪN ---
    source_path = os.path.abspath('01.要件定義_インターフェース一覧（STEP3）.xlsx')
    design_docs_dir = r'D:\Project\151_ISA_AsteriaWrap\trunk\99_FromJP\10_プロジェクト資材\02.IFAgreement\03.確定'
    output_dir = os.path.join(os.getcwd(), 'testcases')

    # Cấu hình mapping Template
    TEMPLATE_MAP = {
        "FtoF": [
            "template_FtoF_単体テスト仕様書兼成績書.xlsx", # Dùng cho IF đầu tiên
            "template_FtoF_単体テスト仕様書兼成績書_同一集約No.のIF展開時.xlsx"  # Dùng cho các IF từ thứ 2 trở đi
        ],
        "DBtoF_1": "template_DBtoF_1_単体テスト仕様書兼成績書.xlsx",
        "DBtoF_2": "template_DBtoF_2_単体テスト仕様書兼成績書.xlsx",
        "FtoDB_1": "template_FtoDB_1_単体テスト仕様書兼成績書.xlsx",
        "FtoDB_2": "template_FtoDB_2_単体テスト仕様書兼成績書.xlsx"
    }

    if not os.path.exists(source_path):
        print(f"Lỗi: Không tìm thấy file {source_path}")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # --- CẤU HÌNH DÒNG VÀ LOGIC THEO TEMPLATE ---
    PATTERN_CONFIGS = {
        "FtoF": {
            "テスト計画書兼結果報告書(マッピング)": {
                10: {"name": "最大(指定)桁数", "checks": ["HAS_IN"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                11: {"name": "指定桁数(固定長)未満", "checks": ["HAS_IN"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                12: {"name": "桁あふれ", "checks": ["IS_IN_NOT_LENGTH", "HAS_IN", "HAS_OUT"], "skip_msg": lambda item, mapping: "固定長のため検証不可" if item.get('x_val') == "固定長" else "入力項目がないため検証不可"},
                13: {"name": "記号混在", "checks": ["HAS_IN"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                14: {"name": "データ型不整合混在", "checks": ["HAS_IN"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                15: {"name": "全角文字混在", "checks": ["IS_TEXT_ZENKAKU_IN"], "skip_msg": lambda item, mapping: "「文字型(全角)/Text」項目がないため検証不可"},
                16: {"name": "全項目値無", "checks": ["HAS_IN"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                17: {"name": "0の場合", "checks": ["IS_NUM_IN"], "skip_msg": lambda item, mapping: "数値型がある入力項目がないため検証不可"},
                18: {"name": "+の場合", "checks": ["IS_NUM_IN"], "skip_msg": lambda item, mapping: "数値型がある入力項目がないため検証不可"},
                19: {"name": "-の場合", "checks": ["IS_NUM_IN", "HAS_NEGATIVE_IN"], "skip_msg": lambda item, mapping: "マイナス項目がないため検証不可"},
                20: {"name": "小数点", "checks": ["IS_NUM_IN", "HAS_DECIMAL_IN"], "skip_msg": lambda item, mapping: "小数点項目がないため検証不可"},
                
                22: {"name": "InputのみNG", "checks": ["HAS_OUT"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                
                25: {"name": "Outputチェック", "checks": ["HAS_OUT"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"},
                26: {"name": "Output 0の場合", "checks": ["IS_NUM_OUT"], "skip_msg": lambda item, mapping: "数値型がある入力項目がないため検証不可"},
                27: {"name": "Output +の場合", "checks": ["IS_NUM_OUT"], "skip_msg": lambda item, mapping: "数値型がある入力項目がないため検証不可"},
                28: {"name": "Output -の場合", "checks": ["IS_NUM_OUT", "HAS_NEGATIVE_OUT"], "skip_msg": lambda item, mapping: "マイナス項目がないため検証不可"},
                29: {"name": "Output 小数点", "checks": ["IS_NUM_OUT", "HAS_DECIMAL_OUT"], "skip_msg": lambda item, mapping: "小数点項目がないため検証不可"},
                
                30: {"name": "固定長ファイル", "checks": ["IS_NUM_OUT", "IS_OUT_FIXED_LENGTH", "NOT_HAS_IN"], "skip_msg": lambda item, mapping: "出力ファイルが固定長でないため検証不可" if item.get('ap_val') != "固定長" else "入力項目がないため検証不可"},
                31: {"name": "可変長ファイル（入力ファイルが固定長の場合）", "checks": ["IS_IN_FIXED_LENGTH", "IS_NUM_OUT", "IS_OUT_NOT_FIXED_LENGTH"], "skip_msg": lambda item, mapping: get_mapping_message(item)},
                
                33: {"name": "固定値", "checks": ["HAS_OUT", "NOT_HAS_IN"], "skip_msg": lambda item, mapping: "固定値項目がないため検証不可"},
                
                35: {"name": "コード・区分変換1", "checks": ["HAS_OUT", "IS_CODE_CONV_OUT"], "skip_msg": lambda item, mapping: "コード変換処理がないため検証不可"},
                36: {"name": "コード・区分変換2", "checks": ["HAS_OUT", "IS_CODE_CONV_OUT"], "skip_msg": lambda item, mapping: "コード変換処理がないため検証不可"},
                37: {"name": "コード・区分変換3", "checks": ["HAS_OUT", "IS_CODE_CONV_OUT"], "skip_msg": lambda item, mapping: "コード変換処理がないため検証不可"},
                
                42: {"name": "Outputチェック", "checks": ["HAS_OUT"], "skip_msg": lambda item, mapping: "入力項目がないため検証不可"}
            },
            "テスト計画書兼結果報告書(共通)": {
                35: {"name": "コード・区分変換", "checks": ["COMMON_CODE_CONV_GRAYOUT"], "skip_msg": lambda item, mapping: "コード変換処理がないため検証不可"}
            },
            "テスト計画書兼結果報告書(個別)": {
                # Sử dụng hàm lambda để tự do format, ghép chuỗi, tính toán...
                # Biến truyền vào `item` chứa: 'if_id', 'if_name', 'master_row' (toàn bộ dữ liệu dòng)
                11: {"col": "F", "value": lambda item: f"{str(item['master_row'][23]).strip() if pd.notnull(item['master_row'][23]) else ''}ファイル。" + (f"{str(item['master_row'][32]).strip()}区切りであること。" if pd.notnull(item['master_row'][32]) and str(item['master_row'][32]).strip() in ['タブ', 'カンマ'] else "区切りがないこと。")},
                12: {"col": "F", "value": lambda item: "文字コードがないこと。" if pd.isnull(item['master_row'][30]) or str(item['master_row'][30]).strip() in ["", "対象外"] else f"文字コードが{str(item['master_row'][30]).strip()}であること。"},
                13: {"col": "F", "value": lambda item: "改行コードがないこと。" if pd.isnull(item['master_row'][31]) or str(item['master_row'][31]).strip() in ["", "対象外"] else f"改行コードが{str(item['master_row'][31]).strip()}であること。"},
                14: {"col": "F", "value": lambda item: "ヘッダー行があること。" if pd.notnull(item['master_row'][29]) and str(item['master_row'][29]).strip() == "有" else "ヘッダー行がないこと。"},
                15: {"col": "F", "value": lambda item: "フッター行があること。" if pd.notnull(item['master_row'][29]) and str(item['master_row'][29]).strip() == "有" else "フッター行がないこと。"},
                16: {"col": "F", "value": lambda item: "囲み文字があること。" if pd.notnull(item['master_row'][33]) and str(item['master_row'][33]).strip() == "有" else "囲み文字がないこと。"},
                18: {"col": "F", "value": lambda item: f"{str(item['master_row'][41]).strip() if pd.notnull(item['master_row'][41]) else ''}ファイル。" + (f"{str(item['master_row'][50]).strip()}区切りであること。" if pd.notnull(item['master_row'][50]) and str(item['master_row'][50]).strip() in ['タブ', 'カンマ'] else "区切りがないこと。")},
                19: {"col": "F", "value": lambda item: "文字コードがないこと。" if pd.isnull(item['master_row'][48]) or str(item['master_row'][48]).strip() in ["", "対象外"] else f"文字コードが{str(item['master_row'][48]).strip()}であること。"},
                20: {"col": "F", "value": lambda item: "改行コードがないこと。" if pd.isnull(item['master_row'][49]) or str(item['master_row'][49]).strip() in ["", "対象外"] else f"改行コードが{str(item['master_row'][49]).strip()}であること。"},
                21: {"col": "F", "value": lambda item: "ヘッダー行があること。" if pd.notnull(item['master_row'][47]) and str(item['master_row'][47]).strip() == "有" else "ヘッダー行がないこと。"},
                22: {"col": "F", "value": lambda item: "フッター行があること。" if pd.notnull(item['master_row'][47]) and str(item['master_row'][47]).strip() == "有" else "フッター行がないこと。"},
                23: {"col": "F", "value": lambda item: "囲み文字があること。" if pd.notnull(item['master_row'][51]) and str(item['master_row'][51]).strip() == "有" else "囲み文字がないこと。"},
            }
        },
        "DBtoF_1": {
            "テスト計画書兼結果報告書(マッピング)": {
                # TODO: Cập nhật số dòng và tên Rule cho DB to File (Template 1)
            },
            "テスト計画書兼結果報告書(共通)": {
            },
            "テスト計画書兼結果報告書(個別)": {
            }
        },
        "DBtoF_2": {
            "テスト計画書兼結果報告書(マッピング)": {
                # TODO: Cập nhật số dòng và tên Rule cho DB to File (Template 2)
            },
            "テスト計画書兼結果報告書(共通)": {
            },
            "テスト計画書兼結果報告書(個別)": {
            }
        },
        "FtoDB_1": {
            "テスト計画書兼結果報告書(マッピング)": {
                # TODO: Cập nhật số dòng và tên Rule cho File to DB (Template 1)
            },
            "テスト計画書兼結果報告書(共通)": {
            },
            "テスト計画書兼結果報告書(個別)": {
            }
        },
        "FtoDB_2": {
            "テスト計画書兼結果報告書(マッピング)": {
                # TODO: Cập nhật số dòng và tên Rule cho File to DB (Template 2)
            },
            "テスト計画書兼結果報告書(共通)": {
            },
            "テスト計画書兼結果報告書(個別)": {
            }
        }
    }

    processing_date = target_date if target_date else datetime.now().strftime('%Y/%m/%d')
    print(f"--- Đang đọc dữ liệu Master cho Testcase (Ngày: {processing_date}) ---")
    df = pd.read_excel(source_path, sheet_name='【STEP3】インターフェース一覧', skiprows=20, header=None)

    testcase_groups = {}
    for index, row in df.iterrows():
        if_ag_id = str(row[65]).strip() if pd.notnull(row[65]) else None
        bm_val = str(row[64]).strip() if pd.notnull(row[64]) else ""
        
        if not if_ag_id or if_ag_id == '-': continue
        if bm_val == '-': continue
        if input_ids and if_ag_id not in input_ids: continue

        # Cột H (index 7): IFA記載集約＃
        if_h = str(row[7]).strip() if pd.notnull(row[7]) else ""

        # Lấy loại Interface từ cột J (index 9)
        if_type_raw = str(row[9]).strip().upper() if pd.notnull(row[9]) else ""
        # Lấy giá trị cột X (index 23)
        x_val = str(row[23]).strip() if pd.notnull(row[23]) else ""

        ap_val = str(row[41]).strip() if pd.notnull(row[41]) else ""
        # Lấy giá trị cột AQ (index 42)
        aq_val = str(row[42]).strip() if pd.notnull(row[42]) else ""
        
        if "DB TO FILE" in if_type_raw: template_keys = ["DBtoF_1", "DBtoF_2"]
        elif "FILE TO DB" in if_type_raw: template_keys = ["FtoDB_1", "FtoDB_2"]
        else: template_keys = ["FtoF"] # Mặc định

        cluster_key = if_h if if_h else if_ag_id
        if cluster_key not in testcase_groups:
            testcase_groups[cluster_key] = []

        for t_key in template_keys:
            testcase_groups[cluster_key].append({
                'if_ag_id': if_ag_id, 'if_id': str(row[2]).strip(), 'if_name': str(row[3]).strip(),
                'template_key': t_key, 'aq_val': aq_val, 'x_val': x_val, 'ap_val': ap_val,
                'master_row': row  # Lưu toàn bộ dữ liệu dòng để mapping sau
            })

    testcase_items = []
    for cluster_key, items in testcase_groups.items():
        for idx, item in enumerate(items):
            item['idx_in_cluster'] = idx
            item['cluster_first_if_id'] = items[0]['if_id']
            testcase_items.append(item)

    if not testcase_items: return

    # --- SETUP CỘT MAPPING ĐỘNG ---
    # Chữ L đại diện cho phần bên Trái (Input), R là bên Phải (Output), C là Cột Chung
    cols_map = {
        'L_SEQ': 'B', 'L_NAME': 'C', 'L_TYPE': 'I', 'L_LEN': 'J', 'L_DEC': 'K', 'L_MINUS': 'L', 'L_SAMPLE': 'D',
        'R_SEQ': 'U', 'R_NAME': 'V', 'R_TYPE': 'AB', 'R_LEN': 'AC', 'R_DEC': 'AD', 'R_MINUS': 'AE', 'R_SAMPLE': 'W',
        'C_MAP1': 'AH', 'C_MAP2': 'AL', 'C_MAP3': 'AJ'
    }
    if cols_input:
        for k, v in zip(['L_SEQ', 'L_NAME', 'L_TYPE', 'L_LEN', 'L_DEC', 'L_MINUS', 'L_SAMPLE'], [x.strip() for x in cols_input.split(',') if x.strip()]):
            cols_map[k] = v.upper()
    if cols_output:
        for k, v in zip(['R_SEQ', 'R_NAME', 'R_TYPE', 'R_LEN', 'R_DEC', 'R_MINUS', 'R_SAMPLE'], [x.strip() for x in cols_output.split(',') if x.strip()]):
            cols_map[k] = v.upper()
    if cols_check:
        for k, v in zip(['C_MAP1', 'C_MAP2', 'C_MAP3'], [x.strip() for x in cols_check.split(',') if x.strip()]):
            cols_map[k] = v.upper()
    
    def col2idx(col_str):
        idx = 0
        for i, char in enumerate(reversed(col_str.strip().upper())):
            idx += (ord(char) - 64) * (26 ** i)
        return idx - 1
    c_idx = {k: col2idx(v) for k, v in cols_map.items()}

    print(f"--- Đang khởi động Excel ---")
    app = xw.App(visible=False)
    app.display_alerts = False
    
    try:
        existing_design_files = os.listdir(design_docs_dir) if os.path.exists(design_docs_dir) else []
        for idx_global, item in enumerate(testcase_items):
            if_ag_id, if_id, template_key = item['if_ag_id'], item['if_id'], item['template_key']
            idx = item['idx_in_cluster']
            
            template_mapped = TEMPLATE_MAP.get(template_key, TEMPLATE_MAP["FtoF"])
            if isinstance(template_mapped, list):
                # Item đầu tiên của cụm (index = 0) dùng template 1, từ item thứ 2 trở đi dùng template 2
                template_name = template_mapped[0] if idx == 0 else template_mapped[-1]
            else:
                template_name = template_mapped
                
            print(template_name)
            template_path = os.path.abspath(template_name)
            if not os.path.exists(template_path):
                print(f"Bỏ qua IF {if_id}: Không tìm thấy template {template_path}")
                continue

            # Đưa tên template_key vào tên file để không bị ghi đè lên nhau
            file_name = f"単体テスト仕様書兼成績書_{if_ag_id}_{if_id}.xlsx"
            
            # Tạo sub-folder mang tên IF (if_id) trong thư mục testcases
            if_dir = os.path.join(output_dir, if_id)
            if not os.path.exists(if_dir):
                os.makedirs(if_dir)
                
            out_p = os.path.join(if_dir, file_name)
            matches = [f for f in existing_design_files if if_ag_id in f and not f.startswith('~$')]
            found_design = sorted(matches, reverse=True)[0] if matches else None

            print(f"Đang tạo: {file_name} (Dùng template: {template_key})")
            shutil.copy(template_path, out_p)
            wb = app.books.open(out_p)

            current_template_config = PATTERN_CONFIGS.get(template_key, PATTERN_CONFIGS["FtoF"])

            try:
                # 0. Clean Names
                for n in wb.api.Names:
                    try:
                        if any(x in n.Value for x in ["#REF!", "[", ".xls"]) or "_FilterDatabase" in n.Name:
                            n.Delete()
                    except: pass

                # 1. Update Cover & History
                s_cover = next((s for s in wb.sheets if '表紙' in s.name), None)
                if s_cover:
                    s_cover.range('A14').number_format = '@'
                    s_cover.range('A25').number_format = '@'
                    s_cover.range('A14').value, s_cover.range('A25').value = f"({if_ag_id}_{if_id})", processing_date
                s_history = next((s for s in wb.sheets if '改定履歴' in s.name), None)
                if s_history: 
                    s_history.range('B3').number_format = '@'
                    s_history.range('B3').value = processing_date

                # 1.1 Update 基本情報
                s_kihon = next((s for s in wb.sheets if '基本情報' in s.name), None)
                if s_kihon:
                    s_kihon.range('C3:C10').number_format = '@'
                    s_kihon.range('C3').value = if_id
                    direction = ""
                    foderName = ""
                    if found_design:
                        upper_design = found_design.upper()
                        if "OUTBOUND" in upper_design:
                            direction = "O"
                            foderName = "_OUT_0"
                        elif "INBOUND" in upper_design:
                            direction = "I"
                            foderName = "_IN_0"
                    if_id_4 = if_id[-4:] if len(if_id) >= 4 else if_id
                    s_kihon.range('C4').value = f"P1J{if_id_4}@{direction}{if_id_4}0001"

                # Các giá trị khác của 基本情報
                aq_ext = item.get('aq_val', '').lower()
                out_filename = f"{if_id}_output.{aq_ext}" if aq_ext else f"{if_id}_output"
                s_kihon.range('C7').value = out_filename
                s_kihon.range('C8').value = f"{if_id_4}{foderName}"
                s_kihon.range('C9').value = fr"\\192.168.1.212\ファイルサーバー\TEST_STEP3\Work\P1J{if_id_4}\parameter.conf"
                s_kihon.range('C10').value = fr"\\192.168.1.212\ファイルサーバー\TEST_STEP3\P1J{if_id_4}\receive"

                # 1.2 Update 外部定義ファイル
                s_gaibu = next((s for s in wb.sheets if '外部定義ファイル' in s.name), None)
                if s_gaibu:
                    s_gaibu.range('B10:B20').number_format = '@'
                    av_val = str(item['master_row'][47]).strip() if pd.notnull(item['master_row'][47]) else ""
                    av_flag = "1" if av_val == "有" else "0"
                    s_gaibu.range('B10').value = f"{if_id},true,true,true,true,1,{av_flag},true,false,true"
                    s_gaibu.range('B13').value = f"{direction}{if_id_4},{if_id_4}{foderName},1,{if_id}"
                    s_gaibu.range('B17').value = f"{direction}{if_id_4},false,false,0,0,false"
                    s_gaibu.range('B20').value = f"{direction}{if_id_4},1,{out_filename},{direction}{if_id_4}{foderName}"

                # 1.5 Xử lý sheet 個別 dựa trên config và dữ liệu Master
                kobetsu_rules = current_template_config.get("テスト計画書兼結果報告書(個別)")
                if kobetsu_rules:
                    s_kobetsu = next((s for s in wb.sheets if '個別' in s.name), None)
                    if s_kobetsu:
                        for r, cfg in kobetsu_rules.items():
                            val_func = cfg.get("value")
                            if callable(val_func):
                                try:
                                    s_kobetsu.range(f'{cfg.get("col", "C")}{r}').number_format = '@'
                                    s_kobetsu.range(f'{cfg.get("col", "C")}{r}').value = val_func(item)
                                except Exception as e:
                                    print(f"Lỗi tính toán sheet 個別 dòng {r}: {e}")

                # 2. Copy Mapping Sheet
                if found_design:

                    
                    wb_src = app.books.open(os.path.join(design_docs_dir, found_design), update_links=False, read_only=True)
                    try:
                        # 1. Tìm sheet gốc "マッピング定義" để copy
                        ss_copy = next((s for s in wb_src.sheets if s.name == "マッピング定義"), None)
                        if not ss_copy:
                            ss_copy = next((s for s in wb_src.sheets if "マッピング定義" in s.name and "_" not in s.name), None)
                        if not ss_copy:
                            ss_copy = next((s for s in wb_src.sheets if "マッピング定義" in s.name), None)

                        if ss_copy:
                            # Xóa các sheet cũ trước để tránh sai lệch index khi copy
                            s_old = next((s for s in wb.sheets if "IFA_マッピング定義" == s.name), None)
                            if s_old: s_old.delete()
                            s_old_sap = next((s for s in wb.sheets if "SAP連携イメージ" == s.name), None)
                            if s_old_sap: s_old_sap.delete()

                            # Tìm sheet テスト計画書兼結果報告書(マッピング) làm vị trí neo
                            anchor_sheet = next((s for s in wb.sheets if "テスト計画書兼結果報告書(マッピング)" in s.name), wb.sheets[-1])

                            ss_copy.api.Copy(After=anchor_sheet.api)
                            new_s = wb.sheets[anchor_sheet.api.Index]
                            new_s.name = "IFA_マッピング定義"
                            try: new_s.used_range.value = new_s.used_range.value
                            except: pass
                            try: new_s.api.Cells.Validation.Delete()
                            except: pass

                            # Xử lý copy thêm sheet SAP連携イメージ nếu có
                            ss_sap = next((s for s in wb_src.sheets if "SAP連携イメージ" in s.name), None)
                            if ss_sap:
                                ss_sap.api.Copy(After=new_s.api)
                                new_sap_s = wb.sheets[new_s.api.Index]
                                new_sap_s.name = "SAP連携イメージ"
                                try: new_sap_s.used_range.value = new_sap_s.used_range.value
                                except: pass
                                try: new_sap_s.api.Cells.Validation.Delete()
                                except: pass

                            # Xử lý copy thêm các sheet từ tham số truyền vào
                            if copy_sheets:
                                extra_sheets = [s.strip() for s in copy_sheets.split(',') if s.strip()]
                                current_anchor = new_sap_s if ss_sap else new_s
                                for sheet_name_to_copy in extra_sheets:
                                    ss_extra = next((s for s in wb_src.sheets if sheet_name_to_copy in s.name), None)
                                    if ss_extra:
                                        s_old_extra = next((s for s in wb.sheets if sheet_name_to_copy in s.name), None)
                                        if s_old_extra: s_old_extra.delete()

                                        ss_extra.api.Copy(After=current_anchor.api)
                                        new_extra_s = wb.sheets[current_anchor.api.Index]
                                        new_extra_s.name = sheet_name_to_copy
                                        try: new_extra_s.used_range.value = new_extra_s.used_range.value
                                        except: pass
                                        try: new_extra_s.api.Cells.Validation.Delete()
                                        except: pass
                                        current_anchor = new_extra_s

                            # 2. Xác định sheet để đọc data tính toán
                            if_h = str(item['master_row'][7]).strip() if pd.notnull(item['master_row'][7]) else ""
                            ss_calc = next((s for s in wb_src.sheets if f"マッピング定義_{if_h}" in s.name), None) if if_h else None
                            
                            source_sheet_for_data = ss_calc if ss_calc else new_s

                            mapping_data = []
                            # Mở rộng vùng đọc đến BZ (tương đương index 77) để tránh lỗi Out Of Bounds khi bạn map cột quá xa
                            data_block = source_sheet_for_data.range('A8:BZ1000').value
                            last_row_idx = 0
                            for i, row_data in enumerate(data_block):
                                def get_val(row, k): 
                                    idx = c_idx.get(k)
                                    return row[idx] if idx is not None and len(row) > idx else None
                                
                                v_in_seq, v_in_name, v_in_type, v_in_k, v_in_l = get_val(row_data, 'L_SEQ'), get_val(row_data, 'L_NAME'), get_val(row_data, 'L_TYPE'), get_val(row_data, 'L_DEC'), get_val(row_data, 'L_MINUS')
                                v_out_seq, v_out_name, v_out_w, v_out_type = get_val(row_data, 'R_SEQ'), get_val(row_data, 'R_NAME'), get_val(row_data, 'R_SAMPLE'), get_val(row_data, 'R_TYPE')
                                v_out_ad, v_out_ae = get_val(row_data, 'R_DEC'), get_val(row_data, 'R_MINUS')
                                
                                map_ah, map_al, map_aj = get_val(row_data, 'C_MAP1'), get_val(row_data, 'C_MAP2'), get_val(row_data, 'C_MAP3')

                                if is_numeric(v_in_seq) or is_numeric(v_out_seq):
                                    # Chuẩn hóa chuỗi Input Type
                                    in_t_raw = str(v_in_type).strip() if v_in_type else ""
                                    in_t_norm = in_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                                    is_num = (in_t_norm == "数値型/Number")
                                    is_text_zenkaku = (in_t_norm == "文字型(全角)/Text")

                                    # Chuẩn hóa chuỗi Output Type
                                    out_t_raw = str(v_out_type).strip() if v_out_type else ""
                                    out_t_norm = out_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                                    is_num_out = (out_t_norm == "数値型/Number")

                                    mapping_data.append({
                                        'in_seq': v_in_seq, 'in_name': v_in_name, 'in_type': v_in_type, 
                                        'in_dec': v_in_k, 'in_minus': v_in_l, 
                                        'out_seq': v_out_seq, 'out_name': v_out_name, 'out_type': v_out_type, 
                                    'out_sample': v_out_w, 'out_dec': v_out_ad, 'out_minus': v_out_ae,
                                    'c_code_conv': map_al, 'c_dec_rules': map_aj,
                                    'is_num': is_num, 'is_text_zenkaku': is_text_zenkaku, 'is_num_out': is_num_out,
                                    'map_ah': map_ah, 'map_al': map_al, 'map_aj': map_aj
                                    })
                                    last_row_idx = i
                                elif any(v is not None for v in [v_in_seq, v_in_name, v_out_seq, v_out_name]): 
                                    last_row_idx = i
                                if i > last_row_idx + 20: break
                            
                            print(f"  -> Số lượng item trong Mapping: {len(mapping_data)}")
                            total_items = len(mapping_data)
                            if total_items > 0:
                                current_template_config = PATTERN_CONFIGS.get(template_key, PATTERN_CONFIGS["FtoF"])
                                
                                for target_sheet_name, sheet_rules in current_template_config.items():
                                    s_test_report = next((s for s in wb.sheets if target_sheet_name in s.name), None)
                                    if not s_test_report: continue

                                    # --- Xử lý riêng cho sheet (共通) ---
                                    if "共通" in target_sheet_name:
                                        has_code_conv = any(
                                        is_numeric(d.get('out_seq')) and is_code_conv(d.get('c_code_conv'))
                                            for d in mapping_data
                                        )
                                        for r, cfg in sheet_rules.items():
                                            errs = []
                                            for check_item in cfg.get("checks", []):
                                                rule_name = check_item if isinstance(check_item, str) else check_item.get("rule")
                                                custom_msg = None if isinstance(check_item, str) else check_item.get("msg")
                                                
                                                err = evaluate_rule(rule_name, item, {}, has_code_conv)
                                                if err: errs.append(custom_msg if custom_msg else err)
                                            if errs:
                                                s_test_report.range(f'C{r}:M{r}').color = (166, 166, 166)
                                                skip_func = cfg.get("skip_msg")
                                                msg = skip_func(item, mapping_data) if callable(skip_func) else "検証不可"
                                                s_test_report.range(f'M{r}').value = msg
                                        continue

                                    # --- Xử lý cho sheet (マッピング) ---
                                    elif "テスト計画書兼結果報告書(マッピング)" in target_sheet_name:
                                        # Bỏ qua xử lý sheet マッピング nếu đang dùng Template 2 trở đi (idx > 0)
                                        
                                        
                                        
                                        if idx > 0: 
                                            val_e14 = s_test_report.range('E14').value
                                            if val_e14:
                                                s_test_report.range('E14').value = str(val_e14).replace("■展開元IF：XXXXX", f"■展開元IF：{item['cluster_first_if_id']}")  # Cập nhật tên IF ở phần header    
                                            val_e16 = s_test_report.range('E16').value
                                            if val_e16:
                                                s_test_report.range('E16').value = str(val_e16).replace("■展開元IF：XXXXX", f"■展開元IF：{item['cluster_first_if_id']}")  # Cập nhật tên IF ở phần header    
                                            continue
                                    
                                        # --- START: Xử lý động dòng 35, 36, 37 cho Code Convert ---
                                        code_conv_cols = [i for i, d in enumerate(mapping_data) if is_numeric(d.get('out_seq')) and is_code_conv(d.get('c_code_conv'))]
                                        num_code_cols = len(code_conv_cols)
                                        dynamic_rules = copy.deepcopy(sheet_rules)
                                        
                                        if 35 in dynamic_rules and 36 in dynamic_rules and 37 in dynamic_rules:
                                            if num_code_cols == 0:
                                                # Không có Code Convert: Xóa dòng 36, 37, giữ lại 1 dòng 35
                                                s_test_report.range('36:37').api.EntireRow.Delete()
                                                new_rules = {}
                                                for r, cfg in dynamic_rules.items():
                                                    if r in (36, 37): continue
                                                    if r > 37: new_rules[r - 2] = cfg # Đẩy các rule phía dưới lên 2 dòng
                                                    else: new_rules[r] = cfg
                                                dynamic_rules = dict(sorted(new_rules.items()))
                                            else:
                                                if num_code_cols > 1:
                                                    # Copy và Insert duy nhất dòng 36 (nét đứt) để format border không bị lỗi nét liền
                                                    for _ in range((num_code_cols - 1) * 3):
                                                        s_test_report.range('36:36').api.EntireRow.Copy()
                                                        s_test_report.range('37:37').api.EntireRow.Insert(Shift=-4121)
                                                    app.api.CutCopyMode = False
                                                
                                                new_rules = {}
                                                base_cfgs = {35: dynamic_rules[35], 36: dynamic_rules[36], 37: dynamic_rules[37]}
                                                
                                                for r, cfg in dynamic_rules.items():
                                                    if r in (35, 36, 37): continue
                                                    if r > 37: new_rules[r + (num_code_cols - 1) * 3] = cfg
                                                    else: new_rules[r] = cfg
                                                    
                                                for k in range(num_code_cols):
                                                    offset = k * 3
                                                    new_rules[35 + offset] = copy.deepcopy(base_cfgs[35])
                                                    new_rules[35 + offset]['target_col_idx'] = code_conv_cols[k]
                                                    
                                                    new_rules[36 + offset] = copy.deepcopy(base_cfgs[36])
                                                    new_rules[36 + offset]['target_col_idx'] = code_conv_cols[k]
                                                    # Thêm rule kiểm tra fixed length cho dòng 36, 39, 42... (bôi xám nếu là 固定長)
                                                    if "IS_IN_NOT_LENGTH" not in new_rules[36 + offset]['checks']:
                                                        new_rules[36 + offset]['checks'].append("IS_IN_NOT_LENGTH")
                                                    new_rules[36 + offset]['skip_msg'] = lambda item, mapping: "固定長のため検証不可" if item.get('x_val') == "固定長" else "コード変換処理がないため検証不可"
                                                    
                                                    new_rules[37 + offset] = copy.deepcopy(base_cfgs[37])
                                                    new_rules[37 + offset]['target_col_idx'] = code_conv_cols[k]
                                                    
                                                    # --- Tự động điền dữ liệu cột E, F, G cho Code Convert ---
                                                    t_data = mapping_data[code_conv_cols[k]]
                                                    out_name = str(t_data.get('out_name', '') or '').strip()
                                                    aj_val = str(t_data.get('c_dec_rules', '') or '').strip() # Cột AJ
                                                    
                                                    # Set format as Text for the entire block
                                                    s_test_report.range(f'D{35 + offset}:G{37 + offset}').number_format = '@'
                                                    
                                                    # Dòng 1 (VD: 35, 38...)
                                                    s_test_report.range(f'D{35 + offset}').value = k * 3 + 1
                                                    s_test_report.range(f'E{35 + offset}').value = out_name
                                                    s_test_report.range(f'F{35 + offset}').value = "コード変換表に値がある場合"
                                                    s_test_report.range(f'G{35 + offset}').value = aj_val
                                                    # Dòng 2 (VD: 36, 39...)
                                                    s_test_report.range(f'D{36 + offset}').value = k * 3 + 2
                                                    s_test_report.range(f'E{36 + offset}').value = ""
                                                    s_test_report.range(f'F{36 + offset}').value = "変換元の値が空白の場合"
                                                    s_test_report.range(f'G{36 + offset}').value = "空白のままとなること。"
                                                    # Dòng 3 (VD: 37, 40...)
                                                    s_test_report.range(f'D{37 + offset}').value = k * 3 + 3
                                                    s_test_report.range(f'E{37 + offset}').value = ""
                                                    s_test_report.range(f'F{37 + offset}').value = "コード変換表に値がない場合"
                                                    s_test_report.range(f'G{37 + offset}').value = "ダミー値に変換"
                                                        
                                                dynamic_rules = dict(sorted(new_rules.items()))
                                        # --- END: Xử lý động dòng 35, 36, 37 cho Code Convert ---

                                        if total_items > 1:
                                            # Dùng EntireColumn để Excel tự đẩy toàn bộ cột, giúp co giãn an toàn các Merge Cell ngang (nếu có)
                                            for _ in range(total_items - 1):
                                                s_test_report.range('H:H').api.EntireColumn.Copy()
                                                s_test_report.range('I:I').api.EntireColumn.Insert(Shift=-4161)
                                            app.api.CutCopyMode = False

                                        header_data = [[],[],[],[]]
                                        row_vals = {r: [] for r in dynamic_rules.keys()}
                                        row_cols = {r: [] for r in dynamic_rules.keys()}
                                        row_reasons = {r: [] for r in dynamic_rules.keys()}
                                        s_ok = {'val': s_test_report.range('A3').value, 'color': s_test_report.range('A3').color}
                                        s_ng = {'val': s_test_report.range('A4').value, 'color': s_test_report.range('A4').color}

                                        for i, data in enumerate(mapping_data):
                                            header_data[0].append(data['in_seq']); header_data[1].append(data['in_name'])
                                            header_data[2].append(data['out_seq']); header_data[3].append(data['out_name'])
                                            
                                            for r, cfg in dynamic_rules.items():
                                                use_ok = True
                                                item_reasons = []
                                                
                                                target_col = cfg.get('target_col_idx')
                                                if target_col is not None and i != target_col:
                                                    use_ok = False
                                                    item_reasons.append("別項目のテスト") # Bỏ qua vì đây là testcase của field khác
                                                else:
                                                    checks = cfg.get("checks", [])
                                                    if not checks: use_ok = False
                                                    
                                                    for check_item in checks:
                                                        rule_name = check_item if isinstance(check_item, str) else check_item.get("rule")
                                                        custom_msg = None if isinstance(check_item, str) else check_item.get("msg")
                                                        
                                                        err_code = evaluate_rule(rule_name, item, data)

                                                        if err_code: # Có lỗi -> trả về mã lỗi
                                                            use_ok = False
                                                            item_reasons.append(custom_msg if custom_msg else err_code)
                                                            break
                                                
                                                s = s_ok if use_ok else s_ng
                                                row_vals[r].append(s['val']); row_cols[r].append(s['color'])
                                                if not use_ok and item_reasons:
                                                    row_reasons[r].extend(item_reasons)

                                        last_col_idx = 8 + total_items - 1
                                        header_range = s_test_report.range(s_test_report.cells(2,8), s_test_report.cells(5, last_col_idx))
                                        header_range.number_format = '@'
                                        header_range.value = header_data
                                    
                                        for r, cfg in dynamic_rules.items():
                                            rg = s_test_report.range(s_test_report.cells(r,8), s_test_report.cells(r, last_col_idx))
                                            rg.number_format = '@'
                                            rg.value = row_vals[r]
                                            

                                            if all(v == s_ng['val'] for v in row_vals[r]):
                                                gray_color = (166, 166, 166)
                                                for col_idx in range(4, last_col_idx + 2):
                                                    try:
                                                        s_test_report.cells(r, col_idx).color = gray_color
                                                    except Exception:
                                                        pass
                                                
                                                skip_func = cfg.get("skip_msg")
                                                if callable(skip_func):
                                                    skip_msg = skip_func(item, mapping_data)
                                                else:
                                                    skip_msg = "対象項目がないため検証不可"
                                                    
                                                s_test_report.cells(r, last_col_idx + 1).value = skip_msg
                                            else:
                                                colors = row_cols[r]
                                                if colors:
                                                    ok_color, ng_color = s_ok['color'], s_ng['color']
                                                    majority_color = ok_color if colors.count(ok_color) > len(colors) // 2 else ng_color
                                                    
                                                    rg.color = majority_color
                                                    for idx, color in enumerate(colors):
                                                        if color != majority_color:
                                                            rg[idx].color = color
                                        
                                        s_test_report.range('A3:A4').clear()

                    finally: wb_src.close()

                # Final Clean & Save
                try: wb.api.UpdateLinks, wb.api.UpdateRemoteReferences = 3, False
                except: pass
                links = wb.api.LinkSources(1)
                if links:
                    for l in links:
                        try: wb.api.BreakLink(l, 1)
                        except: pass
                for s in wb.sheets:
                    try: s.activate(); s.range('A1').select()
                    except: pass
                wb.sheets[0].activate(); wb.save()
            except Exception as e: print(f"[ERR] {file_name}: {str(e)}")
            finally: wb.close()
    finally: app.quit()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('ids', nargs='*'); parser.add_argument('--date', default=None)
    parser.add_argument('--colsInput', type=str, default='', help="Cột Input theo thứ tự: SEQ,NAME,TYPE,LEN,DEC,MINUS,SAMPLE (Mặc định: B,C,I,J,K,L,M)")
    parser.add_argument('--colsOutput', type=str, default='', help="Cột Output theo thứ tự: SEQ,NAME,TYPE,LEN,DEC,MINUS,SAMPLE (Mặc định: U,V,AB,AC,AD,AE,W)")
    parser.add_argument('--colsCheck', type=str, default='', help="Cột Check chung: MAP1,MAP2(CODE_CONV),MAP3(DEC_RULES) (Mặc định: AH,AL,AJ)")
    parser.add_argument('--copySheets', type=str, default='', help="Tên các sheet cần copy thêm, cách nhau bằng dấu phẩy")
    args = parser.parse_args()
    generate_testcase(args.ids if args.ids else None, target_date=args.date, cols_input=args.colsInput, cols_output=args.colsOutput, cols_check=args.colsCheck, copy_sheets=args.copySheets)
