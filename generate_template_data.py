import xlwings as xw
import os
import random
import string
import pandas as pd
import re

# ==========================================
# KHU VỰC 1: CẤU HÌNH (CONFIG) ĐỌC FILE
# ==========================================
# Cấu hình vị trí các dòng/cột theo mô tả của bạn. 
# Nếu thứ tự các dòng (TYPE, DECIMAL, NAME...) bị sai lệch, bạn chỉ cần sửa số ở đây là được.
TEMPLATE_CONFIG = {
    "PATTERN_COL": 3,     # Cột C (Index = 3) chứa tên các Pattern
    "START_DATA_COL": 4,  # Cột D (Index = 4) là cột đầu tiên bắt đầu chứa dữ liệu Field
    
    "INPUT": {
        "META_ROWS": {
            "SEQ": 4,
            "TYPE": 5,      # Giả định dòng 5 là Kiểu dữ liệu
            "DECIMAL": 7,   # Giả định dòng 6 là Decimal
            "NAME": 9,      # Giả định dòng 7 là Tên Field
            "LENGTH": 6,    # Giả định dòng 8 là Độ dài
            # Dòng 9
        },
        # Map cứng Dòng -> Mã Pattern
        "PATTERN_MAP": {
            10: "MAX_LEN",
            11: "MIN_LEN", 
            12: "SYMBOL_MIX",
            13: "INVALID_TYPE",
            14: "EMPTY",
            16: "ZENKAKU_MIX",
            17: "ZERO"
        },
    },
    
    "OUTPUT": {
        "META_ROWS": {
            "SEQ": 29,
            "TYPE": 30,
            "LENGTH": 31,
            "DECIMAL": 32,
            "NAME": 34,
            
            "RULE": 35,     # Dòng 35 chứa ghi chú Rule (VD: padding 0, chuyển đổi thành full width...)
        },
        "PATTERN_MAP": {
            36: "MAX_LEN",  # Data thực sự bắt đầu đổ từ dòng 36
            37: "MIN_LEN",
            38: "SYMBOL_MIX",
            39: "OVER_LEN",
            40: "INVALID_TYPE",
            41: "EMPTY",
            42: "ZENKAKU_MIX",
            43: "ZERO"
        },
    }
}

# ==========================================
# KHU VỰC 2: LOGIC SINH DATA THEO PATTERN
# ==========================================
CODE_MASTER_PATH = r"D:\Project\151_ISA_AsteriaWrap\trunk\04_Testcase\Dummy_data_code.xlsx"
CODE_MASTER_SHEET = "コード変換表"
CODE_MASTER_CACHE = {}
CODE_MASTER_DF = None

def load_code_master_df():
    global CODE_MASTER_DF
    if CODE_MASTER_DF is None:
        if os.path.exists(CODE_MASTER_PATH):
            try:
                # Đọc không có Header để dễ dàng tìm kiếm tên field ở vị trí bất kỳ
                CODE_MASTER_DF = pd.read_excel(CODE_MASTER_PATH, sheet_name=CODE_MASTER_SHEET, header=None)
                print(f"[*] Đã tải xong Code Master ({len(CODE_MASTER_DF)} dòng).")
            except Exception as e:
                print(f"Lỗi khi đọc file Code Master: {e}")
                CODE_MASTER_DF = pd.DataFrame()
        else:
            print(f"Cảnh báo: Không tìm thấy file Code Master tại {CODE_MASTER_PATH}")
            CODE_MASTER_DF = pd.DataFrame()
    return CODE_MASTER_DF

def get_code_master_mapping(rule_desc, in_name, out_name):
    in_name_clean = str(in_name).strip() if in_name and str(in_name).strip() != "None" else ""
    out_name_clean = str(out_name).strip() if out_name and str(out_name).strip() != "None" else ""
    
    if not in_name_clean and not out_name_clean:
        return {}
        
    cache_key = f"{in_name_clean}__{out_name_clean}"
    if cache_key in CODE_MASTER_CACHE:
        return CODE_MASTER_CACHE[cache_key]

    df = load_code_master_df()
    if df.empty:
        return {"1111": "22", "3333": "44"} # Fallback để test nếu không có file

    mapping = {}
    found_row = -1
    
    # Tìm tên field (Input hoặc Output) CHỈ trong cột A (index 0)
    for r in range(len(df)):
        cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
        if (in_name_clean and in_name_clean in cell_val) or (out_name_clean and out_name_clean in cell_val):
            found_row = r
            break
            
    if found_row != -1:
        # Tìm chữ "##■フォーマット" nằm bên dưới field name vừa tìm được trong Cột A
        format_row = -1
        for r_search in range(found_row + 1, min(found_row + 50, len(df))):
            cell_val = str(df.iloc[r_search, 0]).strip() if pd.notnull(df.iloc[r_search, 0]) else ""
            if "##■フォーマット" in cell_val:
                format_row = r_search
                break 
                
        if format_row != -1:
            # Giá trị map bắt đầu cách ##■フォーマット 2 dòng
            start_data_row = format_row + 2
            for r_data in range(start_data_row, len(df)):
                # Đọc input từ cột A (0) và output từ cột B (1)
                in_val = df.iloc[r_data, 0]
                out_val = df.iloc[r_data, 1] if len(df.columns) > 1 else None

                if isinstance(in_val, float) and in_val.is_integer():
                    in_val_str = str(int(in_val))
                else:
                    in_val_str = str(in_val).strip() if pd.notnull(in_val) else ""
                    
                if isinstance(out_val, float) and out_val.is_integer():
                    out_val_str = str(int(out_val))
                else:
                    out_val_str = str(out_val).strip() if pd.notnull(out_val) else ""

                if not in_val_str:
                    break # Dừng đọc khi gặp ô trống ở cột input của mapping

                if in_val_str == "未決定" or out_val_str == "未決定":
                    continue # Bỏ qua các dòng có giá trị là "未決定"

                mapping[in_val_str] = out_val_str

    if not mapping:
        print(f"[*] Cảnh báo: Không tìm thấy mapping cho Input '{in_name_clean}' hoặc Output '{out_name_clean}' trong Code Master.")
        mapping = {"1111": "22"} # Fallback để script không dừng giữa chừng

    CODE_MASTER_CACHE[cache_key] = mapping

    return mapping

def to_zenkaku(text):
    """Chuyển đổi ký tự Nửa góc (Half-width) sang Toàn góc (Full-width)"""
    if not text: return ""
    res = ""
    for c in str(text):
        if c == ' ': res += '　'
        elif 0x21 <= ord(c) <= 0x7E: res += chr(ord(c) + 0xFEE0)
        else: res += c
    return res

def generate_full_random_text(length, is_zenkaku=False):
    if length <= 0: return ""
    if is_zenkaku:
        chars = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ"
    else:
        chars = string.ascii_letters + string.digits
    return ''.join(random.choice(chars) for _ in range(length))

def generate_random_number(length):
    if length <= 0: return ""
    first_digit = random.choice(string.digits[1:]) if length > 0 else ""
    rest_digits = ''.join(random.choice(string.digits) for _ in range(length - 1))
    return first_digit + rest_digits

def generate_mock_value(pattern_key, field_meta):
    """
    Hàm này nhận tên Pattern (từ cột C) và Metadata của Field hiện tại.
    (Sẽ được bổ sung logic thật theo yêu cầu của bạn ở bước sau).
    """
    if not pattern_key:
        return None 
        
    field_name = field_meta.get('NAME')
    
    # 1. Yêu cầu: Chỉ tạo data cho input nào có điền NAME
    if not field_name or str(field_name).strip() in ["", "None"]:
        return None
        
    # 1.1 Kiểm tra xem đây có phải là trường Code Convert không (đọc RULE từ dòng 35)
    rule_desc = str(field_meta.get('RULE_VAL') or '')
    is_code_conv = "code" in rule_desc.lower() or "変換" in rule_desc

    if is_code_conv:
        code_mapping = get_code_master_mapping(rule_desc, field_meta.get('IN_NAME'), field_meta.get('OUT_NAME'))
        if code_mapping:
            if pattern_key == "EMPTY":
                pass # Bỏ qua để chạy xuống logic EMPTY bên dưới (Điền space)
            elif pattern_key == "INVALID_TYPE":
                pass # Trả về một mã không có trong Master
            else:
                return random.choice(list(code_mapping.keys())) # Bốc 1 mã hợp lệ bất kỳ

    data_type = str(field_meta.get('TYPE') or '')
    
    # 2. Xử lý độ dài (Length / Decimal)
    length_val = str(field_meta.get('LENGTH')).strip()
    try:
        total_len = int(float(length_val)) if length_val not in ["", "None", "-"] else 8
    except:
        total_len = 8
        
    dec_val = str(field_meta.get('DECIMAL')).strip()
    try:
        dec_len = int(float(dec_val)) if dec_val not in ["", "None", "-"] else 0
    except:
        dec_len = 0
        
    # print(f"[DEBUG] Field: {field_name} | Type: {data_type} | Length: {total_len} | Decimal: {dec_len}")

    int_len = max(0, total_len - dec_len)
    
    # 3. Logic cho từng Pattern
    if pattern_key == "MAX_LEN":
        if "数値" in data_type or "Number" in data_type:
            int_part = generate_random_number(int_len) if int_len > 0 else "0"
            if dec_len > 0:
                dec_part = ''.join(random.choice(string.digits) for _ in range(dec_len))
                return f"{int_part}.{dec_part}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            # Mặc định là Text (文字型)
            if "全角" in data_type:
                char_count = total_len // 2 # Zenkaku (Toàn góc) tính 2 bytes / 1 ký tự
                return generate_full_random_text(char_count, is_zenkaku=True)
            else:
                return generate_full_random_text(total_len, is_zenkaku=False)
                
    elif pattern_key == "MIN_LEN":
        if "数値" in data_type or "Number" in data_type:
            # MIN_LEN cho số Fixed-length: 1 số và toàn bộ là dấu cách
            int_part = (" " * (int_len - 1) + generate_random_number(1)) if int_len > 0 else "0"
            if dec_len > 0:
                dec_part = generate_random_number(1) + (" " * (dec_len - 1))
                return f"{int_part}.{dec_part}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            if total_len <= 0: return ""
            return "2" + (" " * (total_len - 1))
        else:
            if "全角" in data_type:
                char_count = total_len // 2
                if char_count <= 0: return ""
                return generate_full_random_text(1, is_zenkaku=True) + ("　" * (char_count - 1))
            else:
                if total_len <= 0: return ""
                return generate_full_random_text(1, is_zenkaku=False) + (" " * (total_len - 1))

    elif pattern_key == "OVER_LEN":
        over_int = int_len + 1
        if "数値" in data_type or "Number" in data_type:
            int_part = generate_random_number(over_int)
            if dec_len > 0:
                dec_part = ''.join(random.choice(string.digits) for _ in range(dec_len))
                return f"{int_part}.{dec_part}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            return "202412319" # Dư 1 byte
        else:
            if "全角" in data_type:
                return generate_full_random_text((total_len // 2) + 1, is_zenkaku=True)
            else:
                return generate_full_random_text(total_len + 1, is_zenkaku=False)

    elif pattern_key == "SYMBOL_MIX":
        if "文字" in data_type or "Text" in data_type:
            sym_chars = r"~!@`#$%^&*()[]{}\|"
            if "全角" in data_type:
                char_count = total_len // 2
                if char_count <= 0: return ""
                zenkaku_syms = "～！＠｀＃＄％＾＆＊（）［］｛｝＼｜"
                return "".join(random.choice(zenkaku_syms) for _ in range(char_count))
            else:
                if total_len <= 0: return ""
                return "".join(random.choice(sym_chars) for _ in range(total_len))
        else:
            return "" # Không áp dụng chèn ký tự đặc biệt cho Số và Ngày

    elif pattern_key == "INVALID_TYPE":
        if "数値" in data_type or "Number" in data_type:
            # Kiểu Số -> Phá lỗi bằng cách sinh toàn Chữ cái (A-Z, a-z)
            if total_len <= 0: return ""
            return ''.join(random.choice(string.ascii_letters) for _ in range(total_len))
        elif "日付" in data_type or "Date" in data_type:
            # Kiểu Ngày -> Phá lỗi bằng cách sinh toàn Chữ cái
            if total_len <= 0: return ""
            return ''.join(random.choice(string.ascii_letters) for _ in range(total_len))
        else:
            # Kiểu Chữ -> Phá lỗi bằng cách sinh toàn Số (0-9)
            if "全角" in data_type:
                char_count = total_len // 2
                if char_count <= 0: return ""
                zenkaku_nums = "０１２３４５６７８９"
                return "".join(random.choice(zenkaku_nums) for _ in range(char_count))
            else:
                if total_len <= 0: return ""
                return "".join(random.choice(string.digits) for _ in range(total_len))
                
    elif pattern_key == "EMPTY":
        # EMPTY cho fixed-length: Lấp đầy bằng toàn bộ dấu cách
        if "全角" in data_type:
            char_count = total_len // 2
            if char_count <= 0: return ""
            return "　" * char_count
        else:
            if total_len <= 0: return ""
            return " " * total_len

    return f"[{pattern_key}] {field_name}"

def process_output_logic(in_val, cell_desc, field_meta):
    if isinstance(in_val, float) and in_val.is_integer():
        in_val_str = str(int(in_val))
    else:
        in_val_str = str(in_val) if in_val is not None and str(in_val) != "None" else ""
        
    rule_desc_str = str(cell_desc).strip() if cell_desc is not None else ""
    rule_desc_lower = rule_desc_str.lower()
    
    data_type = str(field_meta.get('TYPE') or '')
    is_number = "数値" in data_type or "Number" in data_type
    is_zenkaku = "全角" in data_type

    length_val = str(field_meta.get('LENGTH')).strip()
    try:
        total_len = int(float(length_val)) if length_val not in ["", "None", "-"] else 8
    except:
        total_len = 8
        
    if is_zenkaku:
        char_limit = total_len // 2
        pad_char = "０" # Dùng số 0 toàn góc nếu là kiểu Zenkaku
        x_char = "Ｘ"   # Ký tự X toàn góc dùng khi không map được code
        pad_char_text = "　" # Khoảng trắng toàn góc
    else:
        char_limit = total_len
        pad_char = "0"
        x_char = "X"
        pad_char_text = " " # Khoảng trắng nửa góc

    # 0.1 Xử lý Code Convert từ Master
    is_code_conv = "code" in rule_desc_lower or "変換" in rule_desc_str
    if is_code_conv:
        code_mapping = get_code_master_mapping(rule_desc_str, field_meta.get('IN_NAME'), field_meta.get('OUT_NAME'))
        if in_val_str in code_mapping:
            in_val_str = str(code_mapping[in_val_str])
        else:
            # Không tìm thấy trong mapping (VD: Data cố tình làm sai hoặc case EMPTY toàn space), điền toàn chữ X theo độ dài Output
            return x_char * char_limit
            
    # 0.2 Xử lý Format Date (YYYY/MM/DD)
    if "yyyy/mm/dd" in rule_desc_lower and len(in_val_str) == 8 and in_val_str.isdigit():
        in_val_str = f"{in_val_str[:4]}/{in_val_str[4:6]}/{in_val_str[6:]}"
    
    # 0. Xử lý Rule Convert Full-width trước khi tính độ dài
    if "0e/0f" in rule_desc_lower or ("fullwidth" in rule_desc_lower and ("0e" in rule_desc_lower or "0f" in rule_desc_lower)):
        in_val_str = " " + to_zenkaku(in_val_str) + " "
    elif "chuyển đổi thành full width" in rule_desc_lower or "全角" in rule_desc_str or "fullwidth" in rule_desc_lower:
        in_val_str = to_zenkaku(in_val_str)
        
    # 0.3 Bước 1: Xử lý đặc thù - Cắt bỏ N ký tự đầu hoặc cuối bằng Regex
    match_tail = re.search(r'末尾(\d+)桁カット|truncate last\s*(\d+)', rule_desc_lower)
    if match_tail and len(in_val_str) > 0:
        cut_len = int(match_tail.group(1) or match_tail.group(2))
        in_val_str = in_val_str[:-cut_len] if cut_len < len(in_val_str) else ""
        
    match_head = re.search(r'先頭(\d+)桁カット|truncate head\s*(\d+)', rule_desc_lower)
    if match_head and len(in_val_str) > 0:
        cut_len = int(match_head.group(1) or match_head.group(2))
        in_val_str = in_val_str[cut_len:]

    # --- LOGIC TỰ ĐỘNG & THEO RULE CHUẨN HOÁ FIXED-LENGTH ---
    # 1. Bước 2: Cắt bớt nếu dữ liệu Input dài hơn Output (Truncate / Crop)
    if "上位桁カット" in rule_desc_str or "前カット" in rule_desc_str or "crop leading" in rule_desc_lower:
        # Ép buộc: Cắt từ bên TRÁI (giữ lại phần ĐUÔI)
        if len(in_val_str) > char_limit:
            out_val = in_val_str[-char_limit:]
        else:
            out_val = in_val_str
    elif "下位桁カット" in rule_desc_str or "後カット" in rule_desc_str or "crop trailing" in rule_desc_lower:
        # Ép buộc: Cắt từ bên PHẢI (giữ lại phần ĐẦU)
        if len(in_val_str) > char_limit:
            out_val = in_val_str[:char_limit]
        else:
            out_val = in_val_str
    else:
        # Tự động theo Kiểu dữ liệu
        if is_number:
            if len(in_val_str) > char_limit: out_val = in_val_str[-char_limit:]
            else: out_val = in_val_str
        else:
            if len(in_val_str) > char_limit: out_val = in_val_str[:char_limit]
            else: out_val = in_val_str
        
    # 2. Bước 3: Padding lấp đầy chiều dài
    clean_val = out_val.strip()
    is_empty = (clean_val == "")
    
    # Nhận diện Rule Padding kiểu mới (Padding X / Padding Space)
    pad_match_new = re.search(r'padding\s*(space|.)', rule_desc_lower)
    # Nhận diện Rule Padding kiểu Nhật (VD: 前ゼロ埋め)
    pad_match_old = re.search(r'(前|後)(ゼロ|0|スペース|空白)埋め', rule_desc_str)
    
    if pad_match_new:
        custom_pad_char = ' ' if pad_match_new.group(1) == 'space' else pad_match_new.group(1)
        # Yêu cầu: Padding thêm vào trước (Padding Left)
        out_val = clean_val.rjust(char_limit, custom_pad_char)
    elif pad_match_old:
        direction = pad_match_old.group(1)
        char_type = pad_match_old.group(2)
        
        custom_pad_char = pad_char if char_type in ['ゼロ', '0'] else pad_char_text
        
        if direction == '前': # Pad Left
            out_val = clean_val.rjust(char_limit, custom_pad_char)
        else: # Pad Right
            out_val = clean_val.ljust(char_limit, custom_pad_char)
    else:
        # Tự động theo Kiểu dữ liệu (Chỉ áp dụng nếu không phải chuỗi rỗng hoàn toàn)
        if not is_empty:
            if is_number or "0埋め" in rule_desc_str:
                out_val = clean_val.rjust(char_limit, pad_char)
            else:
                out_val = clean_val.ljust(char_limit, pad_char_text)
        else:
            # Rỗng hoàn toàn (như Pattern EMPTY) và không có Rule ép buộc -> Lấp đầy bằng khoảng trắng
            out_val = out_val.ljust(char_limit, pad_char_text)
            
    return out_val

# ==========================================
# KHU VỰC 3: XỬ LÝ ĐỌC/GHI EXCEL BẰNG XLWINGS
# ==========================================
def process_template(file_path):
    input_file = os.path.abspath(file_path)
    if not os.path.exists(input_file):
        print(f"Lỗi: Không tìm thấy file: {input_file}")
        return

    # Đặt tên file đầu ra là ..._MockData.xlsx để không ghi đè file gốc
    dir_name = os.path.dirname(input_file)
    name, ext = os.path.splitext(os.path.basename(input_file))
    output_file = os.path.join(dir_name, f"{name}_MockData{ext}")

    print("--- Đang khởi động Excel ---")
    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        wb = app.books.open(input_file)
        sheet = wb.sheets[0] # Xử lý trên Sheet đầu tiên

        def process_block(block_config, block_name):
            print(f"--- Đang phân tích Block: {block_name} ---")
            meta_rows = block_config["META_ROWS"]
            start_col = TEMPLATE_CONFIG["START_DATA_COL"]
            
            in_name_row = TEMPLATE_CONFIG["INPUT"]["META_ROWS"]["NAME"]
            out_name_row = TEMPLATE_CONFIG["OUTPUT"]["META_ROWS"]["NAME"]
            rule_row = TEMPLATE_CONFIG["OUTPUT"]["META_ROWS"]["RULE"]
            
            # 1. Quét dọc theo hàng NAME (Cột D trở đi, quét biên độ 100 cột) để lấy danh sách các Fields
            fields = []
            for col_idx in range(start_col, start_col + 100):
                name_val = sheet.cells(meta_rows["NAME"], col_idx).value
                if name_val is not None and str(name_val).strip() != "":
                    field_meta = {k: sheet.cells(r, col_idx).value for k, r in meta_rows.items()}
                    field_meta["col_idx"] = col_idx
                    field_meta["IN_NAME"] = sheet.cells(in_name_row, col_idx).value
                    field_meta["OUT_NAME"] = sheet.cells(out_name_row, col_idx).value
                    field_meta["RULE_VAL"] = sheet.cells(rule_row, col_idx).value
                    fields.append(field_meta)
            
            print(f"[*] Tìm thấy {len(fields)} fields trong block {block_name}.")

            # 2. Quét qua Config PATTERN_MAP để gọi logic tương ứng cho từng dòng
            pattern_map = block_config.get("PATTERN_MAP", {})
            input_pattern_map = {v: k for k, v in TEMPLATE_CONFIG["INPUT"]["PATTERN_MAP"].items()}

            for row_idx, pattern_key in pattern_map.items():
                for field in fields:
                    if block_name == "INPUT":
                        mock_val = generate_mock_value(pattern_key, field)
                        if mock_val is not None:
                            sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                            sheet.cells(row_idx, field["col_idx"]).value = mock_val
                    elif block_name == "OUTPUT":
                        # Lấy Config/Rule từ dòng RULE được định nghĩa trong cấu hình
                        rule_row = block_config["META_ROWS"].get("RULE")
                        rule_desc = sheet.cells(rule_row, field["col_idx"]).value if rule_row else ""
                        
                        # Kiểm tra xem cột này có Input không (Check dòng NAME của phần INPUT)
                        in_name_row = TEMPLATE_CONFIG["INPUT"]["META_ROWS"]["NAME"]
                        in_name_val = sheet.cells(in_name_row, field["col_idx"]).value
                        has_input = bool(in_name_val and str(in_name_val).strip() != "")
                        
                        if not has_input:
                            # Không có Input -> Check xem dòng RULE (dòng 35) chứa Rule hay chứa Giá trị cố định (Fixed Value)
                            rule_str = str(rule_desc).strip() if rule_desc is not None else ""
                            
                            is_rule = False
                            if rule_str:
                                rule_lower = rule_str.lower()
                                if ("code" in rule_lower or "変換" in rule_str or 
                                    "yyyy/mm/dd" in rule_lower or "全角" in rule_str or 
                                    "chuyển đổi" in rule_lower or "カット" in rule_str or 
                                    "truncate" in rule_lower or "crop" in rule_lower or 
                                    "padding" in rule_lower or "埋め" in rule_str or
                                    "fullwidth" in rule_lower or "0e/0f" in rule_lower):
                                    is_rule = True
                                    
                            if is_rule:
                                # Nếu là Rule (VD: "Padding 1"), lấy input ảo là chuỗi rỗng và áp dụng Rule
                                out_val = process_output_logic("", rule_str, field)
                            else:
                                # Nếu không phải Rule (Fixed Value), dùng chính nó làm input ảo, để rule trống để Padding/Truncate chuẩn form
                                fixed_val_str = str(int(rule_desc)) if isinstance(rule_desc, float) and rule_desc.is_integer() else rule_str
                                out_val = process_output_logic(fixed_val_str, "", field)
                                
                            if out_val is not None:
                                sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                                sheet.cells(row_idx, field["col_idx"]).value = out_val
                        else:
                            # Nếu có Input -> Map data và áp dụng logic Transfer bình thường
                            in_row_idx = input_pattern_map.get(pattern_key)
                            in_val = sheet.cells(in_row_idx, field["col_idx"]).value if in_row_idx else ""
                            out_val = process_output_logic(in_val, rule_desc, field)
                            if out_val is not None:
                                sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                                sheet.cells(row_idx, field["col_idx"]).value = out_val

        process_block(TEMPLATE_CONFIG["INPUT"], "INPUT")
        process_block(TEMPLATE_CONFIG["OUTPUT"], "OUTPUT")

        wb.save(output_file)
        print(f"--- Hoàn tất! Đã sinh data và lưu tại: {os.path.basename(output_file)} ---")
        
    finally:
        try: wb.close()
        except: pass
        app.quit()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Sinh Mock Data dựa trên Template Excel")
    parser.add_argument('file', nargs='?', default='template_create_data.xlsx', help="Đường dẫn file excel input")
    args = parser.parse_args()
    process_template(args.file)