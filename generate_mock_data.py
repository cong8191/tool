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
def generate_mock_value(pattern, field_index, data_type, length_str, scale_str, used_values=None):
    if used_values is None:
        used_values = set()

    # --- Helper để xử lý độ dài ---
    try:
        # Lấy tổng độ dài từ length_str (Cột J)
        if pd.notnull(length_str) and str(length_str).strip() != "":
            total_len = int(float(str(length_str).strip()))
        else:
            total_len = 8
    except (ValueError, TypeError):
        total_len = 8

    try:
        # Lấy số lượng ký tự thập phân từ scale_str (Cột K)
        if pd.notnull(scale_str) and str(scale_str).strip() != "":
            dec_len = int(float(str(scale_str).strip()))
        else:
            dec_len = 0
    except (ValueError, TypeError):
        dec_len = 0
        
    int_len = max(0, total_len - dec_len)

    id_str = str(field_index)

    def to_zenkaku(text):
        # Chuyển đổi các ký tự ASCII (chữ/số) sang Toàn góc (Zenkaku - 2 bytes)
        return ''.join(chr(ord(c) + 0xFEE0) if 0x21 <= ord(c) <= 0x7E else c for c in text)

    def get_unique_value(generator_func):
        # Thử sinh ngẫu nhiên tối đa 50 lần để đảm bảo không trùng và giữ đúng chuẩn format
        for _ in range(50):
            val = generator_func()
            if val == "": 
                return ""  # Chuỗi rỗng thì không thể và không cần làm unique, tránh bị bôi bẩn thành '1'
            
            # BỎ strip() để các chuỗi có dấu cách ở vị trí khác nhau (như " A " và "  A") được coi là khác biệt
            if val not in used_values:
                used_values.add(val)
                return val
        
        # Fallback khi hết random: Cắt bớt đuôi chuỗi gốc để chèn suffix ép unique
        original_val = str(generator_func())
        if original_val == "": return ""
        
        val = original_val
        counter = 1
        
        def get_fallback_suffix(c):
            if "文字型" in data_type or "Text" in data_type:
                # Dùng chữ cái a, b, c... cho Text để không bị lọt số vào
                chars = string.ascii_lowercase
                res = ""
                while c > 0:
                    c -= 1
                    res = chars[c % 26] + res
                    c //= 26
                if "全角" in data_type:
                    return to_zenkaku(res)
                return res
            else:
                # Dùng số 1, 2, 3... cho Number và Date
                return str(c)

        while True:
            if val not in used_values:
                used_values.add(val)
                return val
                
            suffix = get_fallback_suffix(counter)
            
            if len(original_val) >= len(suffix):
                val = original_val[:-len(suffix)] + suffix
            else:
                val = suffix
            counter += 1

    def generate_random_text(length, is_zenkaku=False):
        if length <= 0: return ""
        if is_zenkaku:
            chars = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん"
            space = "　"
        else:
            chars = string.ascii_letters + string.digits
            space = " "
            
        if length == 1:
            return random.choice(chars)
            
        # Theo ý tưởng: hỗn hợp 1 ký tự và các dấu cách ở vị trí ngẫu nhiên
        res = [space] * length
        pos = random.randint(0, length - 1)
        res[pos] = random.choice(chars)
        return ''.join(res)

    def generate_random_number(length, is_decimal_part=False):
        if length <= 0: return ""
        if not is_decimal_part:
            # Phần nguyên không nên bắt đầu bằng số 0 để bảo toàn độ dài khi parse
            first_digit = random.choice(string.digits[1:]) if length > 0 else ""
            rest_digits = ''.join(random.choice(string.digits) for _ in range(length - 1))
            return first_digit + rest_digits
        else:
            # Phần thập phân thì có thể có số 0 thoải mái
            return ''.join(random.choice(string.digits) for _ in range(length))

    # --- Logic cho pattern MAX_LEN ---
    if pattern == "MAX_LEN":
        # 1. Kiểu Số (Number)
        if "数値型" in data_type or "Number" in data_type:
            def gen():
                val = generate_random_number(int_len)
                if dec_len > 0:
                    val = (val if val else "0") + '.' + generate_random_number(dec_len, True)
                return val
            return get_unique_value(gen)

        # 2. Kiểu Ngày (Date)
        elif "日付型" in data_type:
            def gen():
                return (datetime(2024, 1, 1) + timedelta(days=field_index + random.randint(0, 1000))).strftime('%Y%m%d')
            return get_unique_value(gen)

        # 3. Kiểu Text (Bán góc / Toàn góc)
        elif "文字型" in data_type or "Text" in data_type:
            if "全角" in data_type:
                def gen():
                    char_count = total_len // 2  # Tính 2 bytes 1 ký tự toàn góc
                    return generate_random_text(char_count, is_zenkaku=True) if char_count > 0 else ""
                return get_unique_value(gen)
            else:
                def gen():
                    char_count = total_len       # 1 byte 1 ký tự bán góc
                    return generate_random_text(char_count, is_zenkaku=False)
                return get_unique_value(gen)

    # --- Logic cho pattern SHORT_LEN (Thiếu 1 byte) ---
    elif pattern == "SHORT_LEN":
        short_total_len = max(0, total_len - 1)
        if short_total_len == 0: 
            return get_unique_value(lambda: "")
            
        if "数値型" in data_type or "Number" in data_type:
            def gen():
                return " " * (short_total_len - 1) + generate_random_number(1)
            return get_unique_value(gen)
            
        elif "日付型" in data_type:
            def gen():
                return " " * (short_total_len - 1) + generate_random_number(1)
            return get_unique_value(gen)
            
        elif "文字型" in data_type or "Text" in data_type:
            if "全角" in data_type:
                def gen():
                    char_count = short_total_len // 2
                    return generate_random_text(char_count, is_zenkaku=True) if char_count > 0 else ""
                return get_unique_value(gen)
            else:
                def gen():
                    char_count = short_total_len
                    return generate_random_text(char_count, is_zenkaku=False) if char_count > 0 else ""
                return get_unique_value(gen)

    # --- Logic cho pattern OVER_LEN (Dư 1 byte) ---
    elif pattern == "OVER_LEN":
        if "数値型" in data_type or "Number" in data_type:
            def gen():
                over_int_len = int_len + 1
                val = generate_random_number(over_int_len)
                if dec_len > 0:
                    val += '.' + generate_random_number(dec_len, True)
                return val
            return get_unique_value(gen)
            
        elif "日付型" in data_type:
            def gen():
                return (datetime(2024, 1, 1) + timedelta(days=field_index + random.randint(0, 1000))).strftime('%Y%m%d') + "9"
            return get_unique_value(gen)
            
        elif "文字型" in data_type or "Text" in data_type:
            if "全角" in data_type:
                def gen():
                    char_count = (total_len // 2) + 1 # Dư 1 byte, tương đương bắt buộc phải +1 ký tự toàn góc để tràn data
                    return generate_random_text(char_count, is_zenkaku=True)
                return get_unique_value(gen)
            else:
                def gen():
                    char_count = total_len + 1
                    return generate_random_text(char_count, is_zenkaku=False)
                return get_unique_value(gen)

    # Fallback cho các pattern khác chưa được implement
    return get_unique_value(lambda: f"[{pattern}]_{field_index}_{generate_random_text(3)}")

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
                    # B=1, C=2, I=8, J=9, K=10, U=20, V=21, AB=27
                    v_b, v_c, v_i, v_j, v_k, v_u, v_v, v_ab = row_data[1], row_data[2], row_data[8], row_data[9], row_data[10], row_data[20], row_data[21], row_data[27]
                    
                    if is_numeric(v_b) or is_numeric(v_u):
                        # Chuẩn hóa Data Type
                        in_t_raw = str(v_i).strip() if v_i else ""
                        in_t_norm = in_t_raw.replace(' ', '').replace('　', '').replace('／', '/')
                        
                        mapping_data.append({
                            'in_seq': v_b, 'in_name': v_c, 'in_type': v_i, 'in_type_norm': in_t_norm,
                            'in_length': v_j, 'in_decimal': v_k,
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
                                data_type=field.get('in_type_norm', ''), length_str=field.get('in_length'),
                                scale_str=field.get('in_decimal'),
                                used_values=used_values_in_row
                            )
                            row_values.append(mock_val)
                        else:
                            # Không có dấu 'o', giữ nguyên giá trị cũ (VD: '-', '×', khoảng trắng)
                            row_values.append(existing_marks[i])
                            
                    
                    if row_values:
                        print(row_values)
                        s_test_report.range(f'H{row_num}').value = row_values
                        
                for s in wb_new.sheets:
                    try:
                        s.activate()
                        # Với sheet mapping có FreezePanes, chỉ select để tránh lỗi RPC
                        if "テスト計画書兼結果報告書(マッピング)" in s.name:
                            s.range('A1').select()
                        else:
                            # Với các sheet khác, Goto an toàn và hiệu quả
                            app.api.Goto(s.range('A1').api, True)
                    except: pass
                wb_new.sheets[0].activate()
                
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