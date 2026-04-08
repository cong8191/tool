import xlwings as xw
import pandas as pd
import os
import shutil
from datetime import datetime

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
    elif rule_name == "IS_NOT_X_FIXED_LENGTH":
        return "固定長である" if item.get('x_val') == "固定長" else None
    elif rule_name == "IS_OUT_FIXED_LENGTH":
        return "固定長ではない" if item.get('ap_val') == "固定長" else None
    elif rule_name == "IS_OUT_NOT_FIXED_LENGTH":
        return "固定長ではない" if item.get('ap_val') != "固定長" else None
    elif rule_name == "ALWAYS_NG":
        return "常にNG"
    elif rule_name == "IS_NUM_IN":
        return "Inputが数値型ではない" if not data.get('is_num') else None
    elif rule_name == "IS_TEXT_ZENKAKU_IN":
        return "Inputが文字型(全角)ではない" if not data.get('is_text_zenkaku') else None
    elif rule_name == "IS_NUM_OUT":
        return "Outputが数値型ではない" if not data.get('is_num_out') else None
    
    elif rule_name == "HAS_NEGATIVE_IN":
        in_l = str(data.get('in_l', '')).upper() if pd.notnull(data.get('in_l')) else ""

        return "Inputにマイナス項目がない" if "X" not in in_l else None

    elif rule_name == "HAS_NEGATIVE_OUT":
        out_ae = str(data.get('out_ae', '')).upper() if pd.notnull(data.get('out_ae')) else ""

        return "Outputにマイナス項目がない" if "X" not in out_ae else None
    
    elif rule_name == "HAS_DECIMAL_IN":
        in_k = str(data.get('in_k', '')).upper() if pd.notnull(data.get('in_k')) else ""
        try:
            ad_val = float(in_k)
        except ValueError:
            ad_val = 0.0
            
        if ad_val <= 0:
            return "inputに小数点がない"
        
        return None
    
    elif rule_name == "HAS_GT0_IN":
        in_l = str(data.get('in_l', '')).strip() if pd.notnull(data.get('in_l')) else ""
        return "Inputに小数点がない" if in_l in ["", "0", "-"] else None
    elif rule_name == "HAS_X_OUT":
        out_ad = str(data.get('out_ad', '')).upper() if pd.notnull(data.get('out_ad')) else ""
        return "Outputにマイナス項目がない" if "S" not in out_ad and "-" not in out_ad else None
    elif rule_name == "HAS_DECIMAL_OUT":
        out_ad = str(data.get('out_ad', '')).strip() if pd.notnull(data.get('out_ad')) else ""
        out_ah = str(data.get('out_ah', '')).strip() if pd.notnull(data.get('out_ah')) else ""
        out_w = str(data.get('out_w', '')).strip() if pd.notnull(data.get('out_w')) else ""
        
        # 1. Check cột AD xem có > 0 không (ép kiểu float an toàn)
        try:
            ad_val = float(out_ad)
        except ValueError:
            ad_val = 0.0
            
        if ad_val <= 0:
            return "Outputに小数点がない"
            
        # 2. Nếu AD > 0, check tiếp AH xem có chứa "小数点なし" không
        if "小数点なし" in out_ah:
            return "Outputに小数点がない"
            
        # 3. Check sample data (Cột W), nếu có data thì phải chứa dấu "."
        if out_w and "." not in out_w:
            return "Outputに小数点がない"
            
        return None
    
    elif rule_name == "IS_CODE_CONV_OUT":
        return "コード変換処理ではない" if not is_code_conv(data.get('out_ag')) else None
    elif rule_name == "COMMON_CODE_CONV_GRAYOUT":
        return "コード変換処理がない" if not has_code_conv else None
    
    return None

def generate_testcase(input_ids=None, target_date=None):
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
                12: {"name": "桁あふれ", "checks": ["IS_NOT_X_FIXED_LENGTH", "HAS_IN", "HAS_OUT"], "skip_msg": lambda item, mapping: "固定長のため検証不可" if item.get('x_val') == "固定長" else "入力項目がないため検証不可"},
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
                31: {"name": "可変長ファイル（入力ファイルが固定長の場合）", "checks": ["IS_NOT_X_FIXED_LENGTH", "IS_NUM_OUT", "IS_OUT_NOT_FIXED_LENGTH"], "skip_msg": lambda item, mapping: get_mapping_message(item)},
                
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
                17: {"col": "F", "value": lambda item: f"{str(item['master_row'][41]).strip() if pd.notnull(item['master_row'][41]) else ''}ファイル。" + (f"{str(item['master_row'][50]).strip()}区切りであること。" if pd.notnull(item['master_row'][50]) and str(item['master_row'][50]).strip() in ['タブ', 'カンマ'] else "区切りがないこと。")},
                18: {"col": "F", "value": lambda item: "文字コードがないこと。" if pd.isnull(item['master_row'][48]) or str(item['master_row'][48]).strip() in ["", "対象外"] else f"文字コードが{str(item['master_row'][48]).strip()}であること。"},
                19: {"col": "F", "value": lambda item: "改行コードがないこと。" if pd.isnull(item['master_row'][49]) or str(item['master_row'][49]).strip() in ["", "対象外"] else f"改行コードが{str(item['master_row'][49]).strip()}であること。"},
                20: {"col": "F", "value": lambda item: "ヘッダー行があること。" if pd.notnull(item['master_row'][47]) and str(item['master_row'][47]).strip() == "有" else "ヘッダー行がないこと。"},
                21: {"col": "F", "value": lambda item: "フッター行があること。" if pd.notnull(item['master_row'][47]) and str(item['master_row'][47]).strip() == "有" else "フッター行がないこと。"},
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

    testcase_items = []
    for index, row in df.iterrows():
        if_ag_id = str(row[65]).strip() if pd.notnull(row[65]) else None
        if not if_ag_id or if_ag_id == '-': continue
        if input_ids and if_ag_id not in input_ids: continue

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

        for t_key in template_keys:
            testcase_items.append({
                'if_ag_id': if_ag_id, 'if_id': str(row[2]).strip(), 'if_name': str(row[3]).strip(),
                'template_key': t_key, 'aq_val': aq_val, 'x_val': x_val, ap_val: ap_val,
                'master_row': row  # Lưu toàn bộ dữ liệu dòng để mapping sau
            })

    if not testcase_items: return

    print(f"--- Đang khởi động Excel ---")
    app = xw.App(visible=False)
    app.display_alerts = False
    
    try:
        existing_design_files = os.listdir(design_docs_dir) if os.path.exists(design_docs_dir) else []
        for idx, item in enumerate(testcase_items):
            if_ag_id, if_id, template_key = item['if_ag_id'], item['if_id'], item['template_key']
            
            template_mapped = TEMPLATE_MAP.get(template_key, TEMPLATE_MAP["FtoF"])
            if isinstance(template_mapped, list):
                # Item đầu tiên của testcase_items (index = 0) dùng template 1, từ item thứ 2 trở đi dùng template 2
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
            out_p = os.path.join(output_dir, file_name)
            found_design = next((f for f in existing_design_files if if_ag_id in f), None)

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
                    s_cover.range('A14').value, s_cover.range('A25').value = f"({if_ag_id}_{if_id})", processing_date
                s_history = next((s for s in wb.sheets if '改定履歴' in s.name), None)
                if s_history: s_history.range('B3').value = processing_date

                # 1.5 Xử lý sheet 個別 dựa trên config và dữ liệu Master
                kobetsu_rules = current_template_config.get("テスト計画書兼結果報告書(個別)")
                if kobetsu_rules:
                    s_kobetsu = next((s for s in wb.sheets if '個別' in s.name), None)
                    if s_kobetsu:
                        for r, cfg in kobetsu_rules.items():
                            val_func = cfg.get("value")
                            if callable(val_func):
                                try:
                                    s_kobetsu.range(f'{cfg.get("col", "C")}{r}').value = val_func(item)
                                except Exception as e:
                                    print(f"Lỗi tính toán sheet 個別 dòng {r}: {e}")

                # 2. Copy Mapping Sheet
                if found_design:

                    
                    wb_src = app.books.open(os.path.join(design_docs_dir, found_design), update_links=False, read_only=True)
                    try:
                        ss = next((s for s in wb_src.sheets if "マッピング定義" in s.name), None)
                        if ss:
                            s_old = next((s for s in wb.sheets if "IFA_マッピング定義" == s.name), None)
                            if s_old: s_old.delete()
                            ss.api.Copy(After=wb.sheets[-1].api)
                            new_s = wb.sheets[-1]
                            new_s.name = "IFA_マッピング定義"
                            try: new_s.used_range.value = new_s.used_range.value
                            except: pass
                            try: new_s.api.Cells.Validation.Delete()
                            except: pass

                            mapping_data = []
                            # Đọc đến cột AH (index 33) để lấy đầy đủ thuộc tính Output (bao gồm AH và W)
                            data_block = new_s.range('A8:AH1000').value
                            last_row_idx = 0
                            for i, row_data in enumerate(data_block):
                                # B=1, C=2, I=8, K=10, L=11, U=20, V=21, AB=27, AD=29, AE=30, AG=32
                                v_b, v_c, v_i, v_k, v_l, v_u, v_v, v_ab = row_data[1], row_data[2], row_data[8], row_data[10], row_data[11], row_data[20], row_data[21], row_data[27]
                                v_w = row_data[22] if len(row_data) > 22 else None
                                v_ad = row_data[29] if len(row_data) > 29 else None
                                v_ae = row_data[30] if len(row_data) > 30 else None
                                v_ag = row_data[32] if len(row_data) > 32 else None
                                v_ah = row_data[33] if len(row_data) > 33 else None
                                if is_numeric(v_b) or is_numeric(v_u):
                                    # Chuẩn hóa chuỗi Input Type
                                    in_t_raw = str(v_i).strip() if v_i else ""
                                    in_t_norm = in_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                                    is_num = (in_t_norm == "数値型/Number")
                                    is_text_zenkaku = (in_t_norm == "文字型(全角)/Text")

                                    # Chuẩn hóa chuỗi Output Type
                                    out_t_raw = str(v_ab).strip() if v_ab else ""
                                    out_t_norm = out_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                                    is_num_out = (out_t_norm == "数値型/Number")

                                    mapping_data.append({
                                        'in_seq':v_b, 'in_name':v_c, 'in_type':v_i, 
                                        'in_k':v_k, 'in_l':v_l, 'out_seq':v_u, 'out_name':v_v, 'out_type': v_ab, 
                                        'out_w': v_w, 'out_ad': v_ad, 'out_ae': v_ae, 'out_ag': v_ag, 'out_ah': v_ah,
                                        'is_num': is_num, 'is_text_zenkaku': is_text_zenkaku, 'is_num_out': is_num_out
                                    })
                                    last_row_idx = i
                                elif any(v is not None for v in [v_b, v_c, v_u, v_v]): last_row_idx = i
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
                                            is_numeric(d.get('out_seq')) and is_code_conv(d.get('out_ag'))
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
                                    elif "マッピング" in target_sheet_name:
                                        # Bỏ qua xử lý sheet マッピング nếu đang dùng Template 2 trở đi (idx > 0)
                                        if idx > 0: continue

                                        if total_items > 1:
                                            s_test_report.api.Columns(8).Copy()
                                            target_range = s_test_report.api.Range(s_test_report.api.Columns(9), s_test_report.api.Columns(8 + total_items - 1))
                                            target_range.Insert(Shift=-4161)
                                            app.api.CutCopyMode = False

                                        header_data = [[],[],[],[]]
                                        row_vals = {r: [] for r in sheet_rules.keys()}
                                        row_cols = {r: [] for r in sheet_rules.keys()}
                                        row_reasons = {r: [] for r in sheet_rules.keys()}
                                        s_ok = {'val': s_test_report.range('A3').value, 'color': s_test_report.range('A3').color}
                                        s_ng = {'val': s_test_report.range('A4').value, 'color': s_test_report.range('A4').color}

                                        for i, data in enumerate(mapping_data):
                                            header_data[0].append(data['in_seq']); header_data[1].append(data['in_name'])
                                            header_data[2].append(data['out_seq']); header_data[3].append(data['out_name'])
                                            
                                            for r, cfg in sheet_rules.items():
                                                use_ok = True
                                                item_reasons = []
                                                
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
                                        s_test_report.range(s_test_report.cells(2,8), s_test_report.cells(5, last_col_idx)).value = header_data
                                       
                                        for r, cfg in sheet_rules.items():
                                            rg = s_test_report.range(s_test_report.cells(r,8), s_test_report.cells(r, last_col_idx))
                                            rg.value = row_vals[r]
                                            

                                            if all(v == s_ng['val'] for v in row_vals[r]):
                                                s_test_report.range(s_test_report.cells(r,4), s_test_report.cells(r, last_col_idx + 1)).color = (166,166,166)
                                                
                                                # unique_reasons = list(dict.fromkeys(row_reasons[r]))
                                                skip_func = cfg.get("skip_msg")
                                                if callable(skip_func):
                                                    skip_msg = skip_func(item, mapping_data)
                                                else:
                                                    # if unique_reasons:
                                                    #     skip_msg = "、".join(unique_reasons) + "ため検証不可"
                                                    # else:
                                                    #     skip_msg = "対象項目がないため検証不可"
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
    args = parser.parse_args()
    generate_testcase(args.ids if args.ids else None, target_date=args.date)
