import xlwings as xw
import os
import pandas as pd
import random
import string
from datetime import datetime, timedelta

# ==========================================
# KHU VỰC 1: HELPER CHUNG
# ==========================================
def is_numeric(val):
    if pd.isnull(val): return False
    if isinstance(val, (int, float)): return True
    try:
        float(str(val).strip())
        return True
    except ValueError:
        return False

# ==========================================
# KHU VỰC 2: CONFIG DÒNG (RULE MAPPING)
# ==========================================
MOCK_PATTERN_CONFIG = {
    "FtoF": {
        10: {"name": "最大(指定)桁数", "pattern": "MAX_LEN"},
        11: {"name": "指定桁数(固定長)未満", "pattern": "SHORT_LEN"},
        12: {"name": "桁あふれ", "pattern": "OVER_LEN"},
        13: {"name": "記号混在", "pattern": "SYMBOL_MIX"},
        14: {"name": "データ型不整合混在", "pattern": "INVALID_TYPE"},
        15: {"name": "全角文字混在", "pattern": "ZENKAKU_MIX"},
        16: {"name": "全項目値無", "pattern": "EMPTY"},
        17: {"name": "0の場合", "pattern": "ZERO"},
        18: {"name": "+の場合", "pattern": "POSITIVE"},
        19: {"name": "-の場合", "pattern": "NEGATIVE"},
        20: {"name": "小数点", "pattern": "DECIMAL"},
    }
}

# ==========================================
# KHU VỰC 3: PATTERN DATA (LOGIC SINH DATA)
# ==========================================
def generate_mock_value(pattern, field_index, data_type, length_str, scale_str):
    # --- Helper để xử lý độ dài ---
    try:
        if '.' in str(length_str):
            total_len, dec_len = map(int, str(length_str).split('.'))
            int_len = total_len - dec_len
        else:
            total_len = int(length_str) if str(length_str).isdigit() else 8
            int_len, dec_len = total_len, 0
    except (ValueError, TypeError):
        total_len, int_len, dec_len = 8, 8, 0

    id_str = str(field_index)

    # --- Logic cho pattern MAX_LEN ---
    if pattern == "MAX_LEN":
        # 1. Kiểu Số (Number)
        if "数値型/Number" in data_type:
           
            return '2323'

        # 2. Kiểu Ngày (Date)
        elif "日付型" in data_type:
            # Dùng index để ngày tháng là duy nhất
            return (datetime(2024, 1, 1) + timedelta(days=field_index)).strftime('%Y%m%d')

        # 3. Kiểu Text (Bán góc / Toàn góc)
        elif "文字型/Text" in data_type:
            
            return 'ưewe'

    # --- Logic cho pattern SHORT_LEN (Thiếu 1 ký tự) ---
    elif pattern == "SHORT_LEN":
        short_int_len = max(0, int_len - 1)
        short_total_len = max(0, total_len - 1)
        
        if short_total_len == 0:
            return ""
            
        if "数値型" in data_type:
            num_str = (id_str * (short_int_len // len(id_str) + 1))[:short_int_len]
            if dec_len > 0:
                dec_part = (id_str * (dec_len // len(id_str) + 1))[:dec_len]
                return f"{num_str}.{dec_part}" if num_str else f"0.{dec_part}"
            return num_str
            
        elif "日付型" in data_type:
            return (datetime(2024, 1, 1) + timedelta(days=field_index)).strftime('%Y%m%d')[:7]
            
        elif "文字型" in data_type:
            base_char = "あ" if "全角" in data_type else "A"
            return (f"{base_char}{id_str}" * (short_total_len // (len(id_str) + 1) + 1))[:short_total_len]

    # --- Logic cho pattern OVER_LEN (Dư 1 ký tự) ---
    elif pattern == "OVER_LEN":
        over_int_len = int_len + 1
        over_total_len = total_len + 1
        
        if "数値型/Number" in data_type:
            num_str = (id_str * (over_int_len // len(id_str) + 1))[:over_int_len]
            if dec_len > 0:
                dec_part = (id_str * (dec_len // len(id_str) + 1))[:dec_len]
                return f"{num_str}.{dec_part}"
            return num_str
            
        elif "日付型" in data_type:
            # Thêm ký tự 'X' hoặc '9' vào cuối để tạo ra lỗi tràn 9 ký tự
            return (datetime(2024, 1, 1) + timedelta(days=field_index)).strftime('%Y%m%d') + "9"
            
        elif "文字型/Text" in data_type:
            base_char = "あ" if "全角" in data_type else "A"
            return (f"{base_char}{id_str}" * (over_total_len // (len(id_str) + 1) + 1))[:over_total_len]

    # Fallback cho các pattern khác chưa được implement
    return f"[{pattern}]_{id_str}"

def generate_data_from_testcase(input_file):
    input_file = os.path.abspath(input_file)
    if not os.path.exists(input_file):
        print(f"Lỗi: Không tìm thấy file đầu vào: {input_file}")
        return

    # Tạo tên file output bằng cách thêm hậu tố "_MockData"
    dir_name = os.path.dirname(input_file)
    base_name = os.path.basename(input_file)
    name, ext = os.path.splitext(base_name)
    output_file = os.path.join(dir_name, f"{name}_MockData{ext}")

    print("--- Đang khởi động Excel (xlwings) ---")
    app = xw.App(visible=False)
    app.display_alerts = False

    wb_src = None
    wb_new = None
    try:
        wb_src = app.books.open(input_file, update_links=False, read_only=True)
        
        target_sheet_names = ["テスト計画書兼結果報告書(マッピング)", "IFA_マッピング定義"]
        found_sheets = []
        
        # Tìm các sheet cần thiết
        for name in target_sheet_names:
            sheet = next((s for s in wb_src.sheets if name in s.name), None)
            if sheet:
                found_sheets.append(sheet)
            else:
                print(f"Cảnh báo: Không tìm thấy sheet chứa chữ '{name}' trong file.")
        
        if found_sheets:
            print(f"--- Đang tạo file mới và copy {len(found_sheets)} sheet(s): {os.path.basename(output_file)} ---")
            
            wb_new = app.books.add()
            copied_names = []
            
            # Copy lần lượt các sheet sang file mới
            for sheet in found_sheets:
                print(f"  -> Đang copy sheet: {sheet.name}")
                sheet.api.Copy(After=wb_new.sheets[-1].api)
                copied_names.append(sheet.name)
            
            # Xóa các sheet mặc định (ví dụ: Sheet1) không nằm trong danh sách đã copy
            for s in wb_new.sheets:
                if s.name not in copied_names:
                    try: s.delete()
                    except: pass
                    
            wb_new.save(output_file)
            
            # ==========================================
            # KHU VỰC 4: ĐỌC DỮ LIỆU MAPPING TỪ SHEET
            # ==========================================
            s_mapping = next((s for s in wb_new.sheets if "IFA_マッピング定義" in s.name), None)
            mapping_data = []
            
            if s_mapping:
                print("--- Đang đọc các Field Input/Output từ IFA_マッピング定義 ---")
                data_block = s_mapping.range('A8:AH1000').value
                last_row_idx = 0
                
                for i, row_data in enumerate(data_block):
                    # B=1, C=2, I=8, K=10, L=11, U=20, V=21, AB=27
                    v_b, v_c, v_i, v_k, v_l, v_u, v_v, v_ab = row_data[1], row_data[2], row_data[8], row_data[10], row_data[11], row_data[20], row_data[21], row_data[27]
                    
                    if is_numeric(v_b) or is_numeric(v_u):
                        # Chuẩn hóa Data Type
                        in_t_raw = str(v_i).strip() if v_i else ""
                        in_t_norm = in_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                        
                        mapping_data.append({
                            'in_seq': v_b, 'in_name': v_c, 'in_type': v_i, 'in_type_norm': in_t_norm,
                            'in_k': v_k, 'in_l': v_l,
                            'out_seq': v_u, 'out_name': v_v, 'out_type': v_ab
                        })
                        last_row_idx = i
                    elif any(v is not None for v in [v_b, v_c, v_u, v_v]): 
                        last_row_idx = i
                        
                    if i > last_row_idx + 20: break
                
                print(f"  -> Đã lấy được {len(mapping_data)} fields mapping hợp lệ.")

            # ==========================================
            # KHU VỰC 5: GHI MOCK DATA VÀO SHEET
            # ==========================================
            s_test_report = next((s for s in wb_new.sheets if "テスト計画書兼結果報告書(マッピング)" in s.name), None)
            if s_test_report and mapping_data:
                print("--- Đang sinh và ghi Mock Data vào file ---")
                
                # Lấy config cho loại IF này, ví dụ "FtoF"
                # TODO: Cần xác định if_type từ file master sau, tạm hardcode FtoF
                if_type_config = MOCK_PATTERN_CONFIG.get("FtoF", {})
                
                for row_num, config in if_type_config.items():
                    pattern = config.get("pattern")
                    if not pattern: continue
                    
                    print(f"  -> Đang xử lý dòng {row_num} với pattern: {pattern}")
                    
                    # Lấy danh sách các field Input hợp lệ
                    input_fields = [f for f in mapping_data if is_numeric(f.get('in_seq'))]
                    if not input_fields:
                        continue
                        
                    # Đọc dữ liệu hiện tại của dòng bắt đầu từ cột H (index 8)
                    start_cell = s_test_report.cells(row_num, 8)
                    end_cell = s_test_report.cells(row_num, 8 + len(input_fields) - 1)
                    existing_marks = s_test_report.range(start_cell, end_cell).value
                    if not isinstance(existing_marks, list):
                        existing_marks = [existing_marks]
                        
                    row_values = []
                    used_values_in_row = set()
                    # Duyệt qua các field trong mapping để sinh data tương ứng
                    for i, field in enumerate(input_fields):
                        mark = str(existing_marks[i]).strip() if existing_marks[i] is not None else ""
                        
                        # Chỉ sinh data nếu ô hiện tại có chứa dấu 'o', 'O', '〇' (tròn to tiếng Nhật), hoặc '○'
                        if mark in ['o', 'O', '〇', '○']:
                            mock_val = generate_mock_value(
                                pattern=pattern, field_index=int(field.get('in_seq')),
                                data_type=field.get('in_type_norm', ''), length_str=field.get('in_k'),
                                scale_str=field.get('in_l'),
                                used_values=used_values_in_row
                            )
                            row_values.append(mock_val)
                        else:
                            # Không có dấu 'o', giữ nguyên giá trị cũ (VD: '-', '×', khoảng trắng)
                            row_values.append(existing_marks[i])
                    
                    if row_values:
                        s_test_report.range(f'H{row_num}').value = row_values
                print("--- Ghi Mock Data hoàn tất. Đang lưu file... ---")
                wb_new.save()
        else:
            print(f"Lỗi: Không tìm thấy bất kỳ sheet nào cần copy.")

    except Exception as e:
        print(f"Có lỗi xảy ra: {e}")
    finally:
        try:
            wb_src.close()
        except:
            pass
        try:
            wb_new.close()
        except:
            pass
        app.quit()

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Vui lòng truyền đường dẫn file. VD: python generate_mock_data.py path/to/file.xlsx")
    else:
        generate_data_from_testcase(sys.argv[1])