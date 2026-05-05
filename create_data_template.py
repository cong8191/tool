# -*- coding: utf-8 -*-
import xlwings as xw
import pandas as pd
import os
import sys
import argparse
import shutil
import random
import string
import re
from decimal import Decimal, ROUND_HALF_UP
import unicodedata

def safe_str(val):
    if val is None: return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

def is_numeric(val):
    if pd.isnull(val): return False
    if isinstance(val, (int, float)): return True
    try:
        float(str(val).strip())
        return True
    except ValueError:
        return False

def col2idx(col_str):
    if not col_str: return -1
    idx = 0
    for i, char in enumerate(reversed(col_str.strip().upper())):
        idx += (ord(char) - 64) * (26 ** i)
    return idx - 1

# ==========================================
# KHU VỰC: CẤU HÌNH & LOGIC SINH MOCK DATA
# ==========================================
TEMPLATE_CONFIG = {
    "PATTERN_COL": 3,
    "START_DATA_COL": 4,
    "INPUT": {
        "META_ROWS": {"SEQ": 4, "TYPE": 5, "DECIMAL": 7, "NAME": 9, "LENGTH": 6},
        "PATTERN_MAP": {10: "MAX_LEN", 11: "MIN_LEN", 12: "OVER_LEN", 13: "SYMBOL_MIX", 14: "INVALID_TYPE", 15: "ZENKAKU_MIX", 16: "EMPTY", 17: "ZERO", 18: "POSITIVE", 19: "NEGATIVE", 20: "DECIMAL", 21: "SAMPLE"},
    },
    "OUTPUT": {
        "META_ROWS": {"SEQ": 32, "TYPE": 33, "LENGTH": 34, "DECIMAL": 35, "NAME": 37, "RULE": 38},
        "PATTERN_MAP": {39: "MAX_LEN", 40: "MIN_LEN", 41: "OVER_LEN", 42: "SYMBOL_MIX", 43: "INVALID_TYPE", 44: "ZENKAKU_MIX", 45: "EMPTY", 46: "ZERO", 47: "POSITIVE", 48: "NEGATIVE", 49: "DECIMAL", 50: "SAMPLE"},
    }
}

CODE_MASTER_PATH = r"D:\Project\151_ISA_AsteriaWrap\trunk\04_Testcase\Dummy_data_code.xlsx"
CODE_MASTER_SHEET = "コード変換表"
CODE_DEFAULT_SHEET = "コード_Outbound"
CODE_MASTER_CACHE = {}
CODE_MASTER_DF = None
CODE_DEFAULT_DF = None

def load_code_master_df():
    global CODE_MASTER_DF
    if CODE_MASTER_DF is None:
        if os.path.exists(CODE_MASTER_PATH):
            try:
                CODE_MASTER_DF = pd.read_excel(CODE_MASTER_PATH, sheet_name=CODE_MASTER_SHEET, header=None, dtype=str)
            except Exception:
                CODE_MASTER_DF = pd.DataFrame()
        else:
            CODE_MASTER_DF = pd.DataFrame()
    return CODE_MASTER_DF

def load_code_default_df():
    global CODE_DEFAULT_DF
    if CODE_DEFAULT_DF is None:
        if os.path.exists(CODE_MASTER_PATH):
            try:
                CODE_DEFAULT_DF = pd.read_excel(CODE_MASTER_PATH, sheet_name=CODE_DEFAULT_SHEET, header=None, dtype=str)
            except Exception:
                CODE_DEFAULT_DF = pd.DataFrame()
        else:
            CODE_DEFAULT_DF = pd.DataFrame()
    return CODE_DEFAULT_DF

def get_code_default_value(rule_desc, in_name, out_name):
    rule_desc_str = str(rule_desc) if rule_desc is not None else ""
    lines = rule_desc_str.split('\n')
    code_name_from_rule = lines[1].strip() if len(lines) > 1 else ""
    if code_name_from_rule in ["None", "-"]: code_name_from_rule = ""

    in_name_clean = str(in_name).strip() if in_name and str(in_name).strip() != "None" else ""
    out_name_clean = str(out_name).strip() if out_name and str(out_name).strip() != "None" else ""

    search_names = []
    if code_name_from_rule: search_names.append(re.sub(r'_(outbound|inbound)', '', code_name_from_rule, flags=re.IGNORECASE))
    if in_name_clean: search_names.append(re.sub(r'_(outbound|inbound)', '', in_name_clean, flags=re.IGNORECASE))
    if out_name_clean: search_names.append(re.sub(r'_(outbound|inbound)', '', out_name_clean, flags=re.IGNORECASE))
    if not search_names: return None
    df = load_code_default_df()
    if df.empty: return None

    # Vòng 1: Tìm chính xác (Exact match)
    for r in range(len(df)):
        cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
        if not cell_val: continue
        for s_name in search_names:
            if s_name and s_name.lower() == cell_val.lower():
                val = df.iloc[r, 1] if len(df.columns) > 1 else None
                if pd.notnull(val) and str(val).strip().lower() != 'nan': return str(val).strip()

    # Vòng 2: Tìm gần đúng (Tên field nằm trong Master)
    for r in range(len(df)):
        cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
        if not cell_val: continue
        for s_name in search_names:
            if s_name and s_name.lower() in cell_val.lower():
                val = df.iloc[r, 1] if len(df.columns) > 1 else None
                if pd.notnull(val) and str(val).strip().lower() != 'nan': return str(val).strip()
                
    # Vòng 3: Tìm gần đúng (Tên Master nằm trong field)
    for r in range(len(df)):
        cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
        if not cell_val: continue
        for s_name in search_names:
            if s_name and len(cell_val) > 2 and cell_val.lower() in s_name.lower():
                val = df.iloc[r, 1] if len(df.columns) > 1 else None
                if pd.notnull(val) and str(val).strip().lower() != 'nan': return str(val).strip()

    return None

def get_code_master_mapping(rule_desc, in_name, out_name):
    rule_desc_str = str(rule_desc) if rule_desc is not None else ""
    lines = rule_desc_str.split('\n')
    code_name_from_rule = lines[1].strip() if len(lines) > 1 else ""
    if code_name_from_rule in ["None", "-"]: code_name_from_rule = ""

    in_name_clean = str(in_name).strip() if in_name and str(in_name).strip() != "None" else ""
    out_name_clean = str(out_name).strip() if out_name and str(out_name).strip() != "None" else ""
    
    search_names = []
    if code_name_from_rule: search_names.append(code_name_from_rule)
    if in_name_clean: search_names.append(in_name_clean)
    if out_name_clean: search_names.append(out_name_clean)
    
    if not search_names: return {}
    cache_key = "__".join(search_names)
    if cache_key in CODE_MASTER_CACHE:
        return CODE_MASTER_CACHE[cache_key]
    df = load_code_master_df()
    if df.empty:
        return {}
    mapping = {}
    found_row = -1

    # Vòng 1: Tìm chính xác (Exact match)
    for r in range(len(df)):
        cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
        if not cell_val: continue
        for s_name in search_names:
            if s_name and s_name.lower() == cell_val.lower():
                found_row = r
                break
        if found_row != -1:
            break
            
    # Vòng 2: Tìm gần đúng (Tên field nằm trong Master)
    if found_row == -1:
        for r in range(len(df)):
            cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
            if not cell_val: continue
            for s_name in search_names:
                if s_name and s_name.lower() in cell_val.lower():
                    found_row = r
                    break
            if found_row != -1:
                break
                
    # Vòng 3: Tìm gần đúng (Tên Master nằm trong field)
    if found_row == -1:
        for r in range(len(df)):
            cell_val = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
            if not cell_val: continue
            for s_name in search_names:
                if s_name and len(cell_val) > 2 and cell_val.lower() in s_name.lower():
                    found_row = r
                    break
            if found_row != -1:
                break

    if found_row != -1:
        format_row = -1
        for r_search in range(found_row + 1, min(found_row + 50, len(df))):
            cell_val = str(df.iloc[r_search, 0]).strip() if pd.notnull(df.iloc[r_search, 0]) else ""
            if "##■フォーマット" in cell_val:
                format_row = r_search
                break 
        if format_row != -1:
            start_data_row = format_row + 2
            for r_data in range(start_data_row, len(df)):
                in_val = df.iloc[r_data, 0]
                out_val = df.iloc[r_data, 1] if len(df.columns) > 1 else None
                in_val_str = str(in_val).strip() if pd.notnull(in_val) and str(in_val).strip().lower() != 'nan' else ""
                out_val_str = str(out_val).strip() if pd.notnull(out_val) and str(out_val).strip().lower() != 'nan' else ""
                if not in_val_str: break
                if in_val_str == "未決定" or out_val_str == "未決定": continue
                mapping[in_val_str] = out_val_str
    if not mapping: mapping = {"1111": "22"}
    CODE_MASTER_CACHE[cache_key] = mapping
    return mapping

def to_zenkaku(text):
    if not text: return ""
    res = ""
    for c in str(text):
        if c == ' ': res += '　'
        elif 0x21 <= ord(c) <= 0x7E: res += chr(ord(c) + 0xFEE0)
        else: res += c
    return res

def generate_full_random_text(length, is_zenkaku=False):
    if length <= 0: return ""
    chars = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ" if is_zenkaku else string.ascii_letters + string.digits
    return ''.join(random.choice(chars) for _ in range(length))

def generate_random_number(length):
    if length <= 0: return ""
    first_digit = random.choice(string.digits[1:]) if length > 0 else ""
    return first_digit + ''.join(random.choice(string.digits) for _ in range(length - 1))

def generate_mock_value(pattern_key, field_meta, is_fixed_length=True):
    if not pattern_key: return None 
    field_name = field_meta.get('NAME')
    if not field_name or str(field_name).strip() in ["", "None"]: return None
    rule_desc = str(field_meta.get('RULE_VAL') or '')
    is_code_conv = "code" in rule_desc.lower() or "変換" in rule_desc
    if is_code_conv:
        code_mapping = get_code_master_mapping(rule_desc, field_meta.get('IN_NAME'), field_meta.get('OUT_NAME'))
        if code_mapping and pattern_key not in ["EMPTY", "INVALID_TYPE"]:
            return random.choice(list(code_mapping.keys()))
    data_type = str(field_meta.get('TYPE') or '')
    length_val = str(field_meta.get('LENGTH')).strip()
    try:
        total_len = int(float(length_val)) if length_val not in ["", "None", "-"] else 8
    except ValueError:
        total_len = 8
    dec_val = str(field_meta.get('DECIMAL')).strip()
    try:
        dec_len = int(float(dec_val)) if dec_val not in ["", "None", "-"] else 0
    except ValueError:
        dec_len = 0
    # Tính toán chiều dài phần nguyên (Trừ đi phần thập phân và 1 byte cho dấu chấm)
    int_len = max(0, total_len - dec_len - (1 if dec_len > 0 else 0))
    
    if pattern_key == "MAX_LEN":
        if "数値" in data_type or "Number" in data_type:
            int_part = generate_random_number(int_len) if int_len > 0 else "0"
            if dec_len > 0: return f"{int_part}.{''.join(random.choice(string.digits) for _ in range(dec_len))}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            return generate_full_random_text(total_len // 2 if "全角" in data_type else total_len, is_zenkaku="全角" in data_type)
    elif pattern_key == "MIN_LEN":
        if "数値" in data_type or "Number" in data_type:
            base_int = generate_random_number(1) if int_len > 0 else "0"
            int_part = (" " * (int_len - 1) + base_int) if (is_fixed_length and int_len > 0) else base_int
            if dec_len > 0:
                base_dec = generate_random_number(1)
                dec_part = (base_dec + ' ' * (dec_len - 1)) if is_fixed_length else base_dec
                return f"{int_part}.{dec_part}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            return "2" + ((" " * (total_len - 1)) if is_fixed_length else "") if total_len > 0 else ""
        else:
            if "全角" in data_type:
                char_count = total_len // 2
                return generate_full_random_text(1, is_zenkaku=True) + (("　" * (char_count - 1)) if is_fixed_length else "") if char_count > 0 else ""
            else:
                return generate_full_random_text(1, is_zenkaku=False) + ((" " * (total_len - 1)) if is_fixed_length else "") if total_len > 0 else ""
    elif pattern_key == "OVER_LEN":
        over_int = int_len + 1
        if "数値" in data_type or "Number" in data_type:
            int_part = generate_random_number(over_int)
            if dec_len > 0: return f"{int_part}.{''.join(random.choice(string.digits) for _ in range(dec_len + 1))}"
            return int_part
        elif "日付" in data_type or "Date" in data_type:
            return "202412319"
        else:
            return generate_full_random_text((total_len // 2) + 1 if "全角" in data_type else total_len + 1, is_zenkaku="全角" in data_type)
    elif pattern_key == "SYMBOL_MIX":
        if "文字" in data_type or "Text" in data_type or "数値" in data_type or "Number" in data_type:
            if "全角" in data_type:
                char_count = total_len // 2
                return "".join(random.choice("！＠｀＃＄％＾＆＊（）［］｛｝＼｜") for _ in range(char_count)) if char_count > 0 else ""
            else:
                return "".join(random.choice(r"!@`#$%^&*()[]{}\|") for _ in range(total_len)) if total_len > 0 else ""
        return ""
    elif pattern_key == "INVALID_TYPE":
        if "数値" in data_type or "Number" in data_type or "日付" in data_type or "Date" in data_type:
            return ''.join(random.choice(string.ascii_letters) for _ in range(total_len)) if total_len > 0 else ""
        else:
            if "全角" in data_type:
                char_count = total_len // 2
                return "".join(random.choice("０１２３４５６７８９") for _ in range(char_count)) if char_count > 0 else ""
            else:
                return "".join(random.choice(string.digits) for _ in range(total_len)) if total_len > 0 else ""
    elif pattern_key == "ZENKAKU_MIX":
        if "数値" in data_type or "Number" in data_type:
            int_part = generate_random_number(int_len) if int_len > 0 else "0"
            res = f"{int_part}.{''.join(random.choice(string.digits) for _ in range(dec_len))}" if dec_len > 0 else int_part
            return res.rjust(total_len, ' ') if is_fixed_length else res
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            if "全角" in data_type:
                char_count = total_len // 2
                if char_count < 1: return ""
                res = [random.choice("あいうえおかきくけこさしすせそ") for _ in range(char_count)]
                res[random.randint(0, char_count - 1)] = ''.join(random.choice(string.ascii_letters) for _ in range(2))
                return "".join(res)
            else:
                if total_len < 2: return " " if is_fixed_length else "A"
                res = [random.choice(string.ascii_letters) for _ in range(total_len)]
                pos = random.randint(0, total_len - 2)
                res[pos] = random.choice("あいうえおかきくけこ")
                del res[pos+1]
                return "".join(res)
    elif pattern_key == "EMPTY":
        if not is_fixed_length: return ""
        if "全角" in data_type:
            return "　" * (total_len // 2) if total_len // 2 > 0 else ""
        else:
            return " " * total_len if total_len > 0 else ""
    elif pattern_key == "ZERO":
        if "数値" in data_type or "Number" in data_type:
            val = "0." + "0" * dec_len if dec_len > 0 else "0"
            return val.rjust(total_len, ' ') if is_fixed_length else val
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            return generate_full_random_text(total_len // 2 if "全角" in data_type else total_len, is_zenkaku="全角" in data_type)
    elif pattern_key in ["POSITIVE", "NEGATIVE"]:
        if "数値" in data_type or "Number" in data_type:
            sign_len = 1 if pattern_key == "NEGATIVE" else 0
            random_len = random.randint(1, max(1, int_len - sign_len))
            val = generate_random_number(random_len)
            if dec_len > 0: val += '.' + ''.join(random.choice(string.digits) for _ in range(dec_len))
            res = ("" if pattern_key == "POSITIVE" else "-") + val
            return res.rjust(total_len, ' ') if is_fixed_length else res
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            return generate_full_random_text(total_len // 2 if "全角" in data_type else total_len, is_zenkaku="全角" in data_type)
    elif pattern_key == "DECIMAL":
        if "数値" in data_type or "Number" in data_type:
            res = f"{generate_random_number(random.randint(1, max(1, int_len)))}.{''.join(random.choice(string.digits) for _ in range(dec_len if dec_len > 0 else 1))}"
            return res.rjust(total_len, ' ') if is_fixed_length else res
        elif "日付" in data_type or "Date" in data_type:
            return "20241231"[:total_len] if total_len < 8 else "20241231"
        else:
            return generate_full_random_text(total_len // 2 if "全角" in data_type else total_len, is_zenkaku="全角" in data_type)
    return f"[{pattern_key}] {field_name}"

def process_output_logic(in_val, cell_desc, field_meta, is_output_fixed_length=True):
    in_val_str = str(int(in_val)) if isinstance(in_val, float) and in_val.is_integer() else str(in_val) if in_val is not None and str(in_val) != "None" else ""
    rule_desc_str = str(cell_desc).strip() if cell_desc is not None else ""
    rule_desc_lower = rule_desc_str.lower()
    data_type = str(field_meta.get('TYPE') or '')
    is_number = "数値" in data_type or "Number" in data_type
    is_zenkaku = "全角" in data_type
    length_val = str(field_meta.get('LENGTH')).strip()
    try:
        total_len = int(float(length_val)) if length_val not in ["", "None", "-"] else 8
    except ValueError:
        total_len = 8

    dec_val = str(field_meta.get('DECIMAL')).strip()
    try:
        dec_len = int(float(dec_val)) if dec_val not in ["", "None", "-"] else 0
    except ValueError:
        dec_len = 0

    is_0e0f = "0e/0f" in rule_desc_lower or ("fullwidth" in rule_desc_lower and ("0e" in rule_desc_lower or "0f" in rule_desc_lower))
    char_limit = max(0, (total_len - 2) // 2) if is_0e0f else total_len // 2 if is_zenkaku or is_0e0f else total_len
    pad_char = "０" if is_zenkaku or is_0e0f else "0"
    x_char = "Ｘ" if is_zenkaku or is_0e0f else "X"
    pad_char_text = "　" if is_zenkaku or is_0e0f else " "
    if "code" in rule_desc_lower or "変換" in rule_desc_str:
        if in_val_str == "" or in_val_str.strip() == "":
            return ""
            
        code_mapping = get_code_master_mapping(rule_desc_str, field_meta.get('IN_NAME'), field_meta.get('OUT_NAME'))
        if in_val_str in code_mapping: in_val_str = str(code_mapping[in_val_str])
        else:
            default_val = get_code_default_value(rule_desc_str, field_meta.get('IN_NAME'), field_meta.get('OUT_NAME'))
            if default_val is not None:
                in_val_str = default_val
            else:
                out_val = x_char * char_limit
                return " " + out_val + " " if is_0e0f else out_val
    if "yyyy/mm/dd" in rule_desc_lower and len(in_val_str) == 8 and in_val_str.isdigit():
        in_val_str = f"{in_val_str[:4]}/{in_val_str[4:6]}/{in_val_str[6:]}"
        
    # Xử lý Rule: 四捨五入 (Làm tròn)
    # Match: "小数第3位を四捨五入" (Làm tròn ở vị trí thứ 3 -> lấy 2 số thập phân), "小数第三位を四捨五入", "round 2"
    round_match_jp = re.search(r'小数第([0-9０-９一二三四五六七八九]+)位を四捨五入', rule_desc_str)
    round_match_en = re.search(r'round\s*([0-9０-９]+)', rule_desc_lower)
    
    if (round_match_jp or round_match_en) and in_val_str.replace('.', '', 1).replace('-', '', 1).isdigit():
        keep_decimals = 0
        if round_match_jp:
            kanji_to_num = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
            val_str = round_match_jp.group(1)
            target_digit = kanji_to_num.get(val_str) if val_str in kanji_to_num else int(val_str)
            keep_decimals = max(0, target_digit - 1)
        elif round_match_en:
            keep_decimals = int(round_match_en.group(1))
            
        try:
            d = Decimal(in_val_str)
            quantizer = Decimal('1.' + '0' * keep_decimals) if keep_decimals > 0 else Decimal('1')
            in_val_str = str(d.quantize(quantizer, rounding=ROUND_HALF_UP))
        except:
            pass

    # Xử lý Rule: 小数部をカット (Cắt bỏ phần thập phân)
    if "小数部をカット" in rule_desc_str or "cut decimal" in rule_desc_lower or "truncate decimal" in rule_desc_lower:
        in_val_str = in_val_str.split('.')[0]

    # Xử lý Rule: 小数点不要 (Bỏ dấu thập phân) hoặc no decimal
    if "小数点不要" in rule_desc_str or "no decimal" in rule_desc_lower:
        if dec_len > 0 and in_val_str.replace("-", "").replace(".", "").isdigit():
            if "." in in_val_str:
                parts = in_val_str.split(".")
                int_part = parts[0]
                # Padding thêm 0 vào đuôi nếu thiếu, hoặc cắt bớt nếu dư
                dec_part = parts[1].ljust(dec_len, "0")[:dec_len]
                in_val_str = int_part + dec_part
            else:
                # Nếu là số nguyên không có dấu chấm, cộng thẳng số lượng 0 vào sau
                in_val_str = in_val_str + "0" * dec_len
        else:
            in_val_str = in_val_str.replace(".", "")

    match_tail = re.search(r'末尾(\d+)桁カット|truncate last\s*(\d+)', rule_desc_lower)
    if match_tail and len(in_val_str) > 0:
        cut_len = int(match_tail.group(1) or match_tail.group(2))
        in_val_str = in_val_str[:-cut_len] if cut_len < len(in_val_str) else ""
    match_head = re.search(r'先頭(\d+)桁カット|truncate head\s*(\d+)', rule_desc_lower)
    if match_head and len(in_val_str) > 0:
        cut_len = int(match_head.group(1) or match_head.group(2))
        in_val_str = in_val_str[cut_len:]

    head_str = ""
    match_add_head = re.search(r'先頭「([^」]+)」|add\s+(?:head|prefix)\s+["\'「]?([^"\'」\s]+)', rule_desc_str, re.IGNORECASE)
    if match_add_head:
        head_str = match_add_head.group(1) or match_add_head.group(2)
        
    last_str = ""
    match_add_last = re.search(r'(?:後尾|末尾)「([^」]+)」|add\s+(?:last|suffix)\s+["\'「]?([^"\'」\s]+)', rule_desc_str, re.IGNORECASE)
    if match_add_last:
        last_str = match_add_last.group(1) or match_add_last.group(2)

    if is_0e0f or "chuyển đổi thành full width" in rule_desc_lower or "全角" in rule_desc_str or "fullwidth" in rule_desc_lower:
        in_val_str = to_zenkaku(in_val_str)
        head_str = to_zenkaku(head_str)
        last_str = to_zenkaku(last_str)

    in_val_str = head_str + in_val_str + last_str

    if "上位桁カット" in rule_desc_str or "前カット" in rule_desc_str or "crop leading" in rule_desc_lower:
        in_val_str = in_val_str[-char_limit:] if len(in_val_str) > char_limit else in_val_str
    elif "下位桁カット" in rule_desc_str or "後カット" in rule_desc_str or "crop trailing" in rule_desc_lower:
        in_val_str = in_val_str[:char_limit] if len(in_val_str) > char_limit else in_val_str
    else:
        if is_number: in_val_str = in_val_str[-char_limit:] if len(in_val_str) > char_limit else in_val_str
        else: in_val_str = in_val_str[:char_limit] if len(in_val_str) > char_limit else in_val_str
        
    clean_val = in_val_str.strip()
    is_empty = (clean_val == "")

    # Nếu output là fixed length, trường required (có giá trị là X), input rỗng -> padding X (Text) hoặc 0 (Number)
    req_val = str(field_meta.get('OUT_REQ', '')).strip().upper()
    
    if is_output_fixed_length and req_val in ['X', 'Ｘ'] and is_empty:
        if is_number:
            in_val_str = "".ljust(char_limit, pad_char)
        else:
            in_val_str = "".ljust(char_limit, x_char)
        out_val = in_val_str
        return " " + out_val + " " if is_0e0f else out_val

    pad_match_en = re.search(r'pad(?:ding)?\s*(left|right)?\s*(space|0|[^\s])?', rule_desc_lower)
    pad_match_jp = re.search(r'(前|後)(ゼロ|0|スペース|空白)埋め', rule_desc_str)
    if pad_match_en:
        direction = pad_match_en.group(1)
        pad_char_match = pad_match_en.group(2)
        custom_pad_char = ' ' if not pad_char_match or pad_char_match == 'space' else pad_char_match
        
        if direction == 'left' or (not direction and (is_number or "0埋め" in rule_desc_str)):
            in_val_str = clean_val.rjust(char_limit, custom_pad_char)
        else:
            in_val_str = clean_val.ljust(char_limit, custom_pad_char)
    elif pad_match_jp:
        direction, char_type = pad_match_jp.group(1), pad_match_jp.group(2)
        custom_pad_char = pad_char if char_type in ['ゼロ', '0'] else pad_char_text
        in_val_str = clean_val.rjust(char_limit, custom_pad_char) if direction == '前' else clean_val.ljust(char_limit, custom_pad_char)
    else:
        if not is_empty: in_val_str = clean_val.rjust(char_limit, pad_char) if is_number or "0埋め" in rule_desc_str else clean_val.ljust(char_limit, pad_char_text)
        else: in_val_str = in_val_str.ljust(char_limit, pad_char_text)
        
    out_val = in_val_str
    return " " + out_val + " " if is_0e0f else out_val

def fill_template_data(sheet, mapping_data, is_input_fixed_length=True, is_output_fixed_length=True):
    def is_invalid_len(val):
        v = str(val).strip()
        if v in ["", "None", "-"]: return False
        try:
            float(v)
            return False
        except ValueError:
            return True

    def is_valid_sample(val, data_type):
        if val is None: return False
        val_str = str(val).strip()
        if not val_str or val_str in ["None", "未定", "-", "NaN"]: return False
        dt_str = str(data_type)
        dt_lower = dt_str.lower()
        
        # 1. Nếu kiểu KHÔNG phải 全角 -> Báo lỗi nếu có ký tự Toàn góc (F, W)
        if "全角" not in dt_str and any(unicodedata.east_asian_width(c) in ['F', 'W'] for c in val_str):
            return False
            
        # 2. Nếu kiểu LÀ 全角 -> Báo lỗi nếu có ký tự Nửa góc (H: Half-width Katakana, Na: Narrow ASCII)
        if "全角" in dt_str and any(unicodedata.east_asian_width(c) in ['H', 'Na'] for c in val_str):
            return False

        if "数値" in dt_lower or "number" in dt_lower:
            val_str = val_str.replace(',', '')
            try:
                float(val_str)
                return True
            except ValueError:
                return False
        elif "日付" in dt_lower or "date" in dt_lower:
            return val_str.replace('/', '').replace('-', '').isdigit()
        return True

    skipped_input_patterns = set()

    def process_block(block_config, block_name):
        meta_rows = block_config["META_ROWS"]
        start_col = TEMPLATE_CONFIG["START_DATA_COL"]
        in_name_row = TEMPLATE_CONFIG["INPUT"]["META_ROWS"]["NAME"]
        out_name_row = TEMPLATE_CONFIG["OUTPUT"]["META_ROWS"]["NAME"]
        rule_row = TEMPLATE_CONFIG["OUTPUT"]["META_ROWS"]["RULE"]
        fields = []
        for i, col_idx in enumerate(range(start_col, start_col + len(mapping_data))):
            name_val = sheet.cells(meta_rows["NAME"], col_idx).value
            if name_val is not None and str(name_val).strip() != "":
                field_meta = {k: sheet.cells(r, col_idx).value for k, r in meta_rows.items()}
                field_meta["col_idx"] = col_idx
                field_meta["IN_NAME"] = sheet.cells(in_name_row, col_idx).value
                field_meta["OUT_NAME"] = sheet.cells(out_name_row, col_idx).value
                field_meta["RULE_VAL"] = sheet.cells(rule_row, col_idx).value
                field_meta["IN_REQ"] = mapping_data[i].get('in_req')
                field_meta["OUT_REQ"] = mapping_data[i].get('out_req')
                fields.append(field_meta)
                
                # Bôi đỏ ô Length trên header nếu có giá trị không hợp lệ (VD: 未定)
                if is_invalid_len(field_meta.get('LENGTH')):
                    sheet.cells(meta_rows["LENGTH"], col_idx).color = (255, 0, 0)

        pattern_map = block_config.get("PATTERN_MAP", {})
        input_pattern_map = {v: k for k, v in TEMPLATE_CONFIG["INPUT"]["PATTERN_MAP"].items()}
        
        max_len_cache = {}  # Bộ nhớ đệm lưu giá trị MAX_LEN theo IN_SEQ
        for row_idx, pattern_key in pattern_map.items():
            row_vals = {}
            has_any_dot = False
            for field in fields:
                val = sheet.cells(row_idx, field["col_idx"]).value
                val_str = str(val).strip() if val is not None else ""
                row_vals[field["col_idx"]] = val_str
                if val_str in ['o', 'O', '〇', '○']:
                    has_any_dot = True
            if block_name == "INPUT" and not has_any_dot and pattern_key != "SAMPLE":
                skipped_input_patterns.add(pattern_key)
                for field in fields:
                    sheet.cells(row_idx, field["col_idx"]).value = ""
                continue
            if block_name == "OUTPUT":
                in_row_idx = input_pattern_map.get(pattern_key)
                if in_row_idx:
                    if pattern_key in skipped_input_patterns:
                        for field in fields:
                            sheet.cells(row_idx, field["col_idx"]).value = ""
                            sheet.cells(row_idx, field["col_idx"]).color = (191, 191, 191)
                        continue
            row_mock_cache = {} # Cache mock values cho dòng hiện tại theo IN_SEQ
            for field in fields:
                # Nếu không xác định được độ dài -> bôi đỏ và bỏ qua, không fill data (giữ nguyên chấm tròn)
                if is_invalid_len(field.get('LENGTH')):
                    sheet.cells(row_idx, field["col_idx"]).color = (255, 0, 0)
                    continue

                if block_name == "INPUT":
                    if pattern_key == "SAMPLE":
                        sample_val = sheet.cells(row_idx, field["col_idx"]).value
                        if is_valid_sample(sample_val, field.get('TYPE')):
                            continue # Giữ nguyên giá trị Sample Data đã lấy từ file thiết kế (Hợp lệ)
                        else:
                            # Sai kiểu -> Bôi vàng cảnh báo và giữ nguyên giá trị
                            sheet.cells(row_idx, field["col_idx"]).color = (255, 255, 0)
                            continue
                    else:
                        has_dot = row_vals[field["col_idx"]] in ['o', 'O', '〇', '○']
                        actual_pattern = pattern_key if has_dot else "MAX_LEN"
                    
                    in_seq_raw = field.get("SEQ")
                    in_seq = str(in_seq_raw).strip() if in_seq_raw is not None and str(in_seq_raw).strip() != "" else f"col_{field['col_idx']}"
                    
                    if actual_pattern == "MAX_LEN":
                        if in_seq not in max_len_cache:
                            max_len_cache[in_seq] = generate_mock_value("MAX_LEN", field, is_input_fixed_length)
                        mock_val = max_len_cache[in_seq]
                    else:
                        if in_seq not in row_mock_cache:
                            row_mock_cache[in_seq] = generate_mock_value(actual_pattern, field, is_input_fixed_length)
                        mock_val = row_mock_cache[in_seq]
                    if mock_val is not None:
                        sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                        sheet.cells(row_idx, field["col_idx"]).value = mock_val
                elif block_name == "OUTPUT":
                    rule_row_idx = block_config["META_ROWS"].get("RULE")
                    rule_desc = sheet.cells(rule_row_idx, field["col_idx"]).value if rule_row_idx else ""
                    in_name_val = sheet.cells(in_name_row, field["col_idx"]).value
                    has_input = bool(in_name_val and str(in_name_val).strip() != "")
                    if not has_input:
                        rule_str = str(rule_desc).strip() if rule_desc is not None else ""
                        is_rule = False
                        if rule_str:
                            rule_lower = rule_str.lower()
                            if any(x in rule_lower for x in ["code", "変換", "yyyy/mm/dd", "全角", "chuyển đổi", "カット", "truncate", "crop", "pad", "埋め", "fullwidth", "0e/0f", "四捨五入", "round", "不要", "no decimal", "付与", "先頭", "後尾", "末尾", "add head", "add last", "add prefix", "add suffix", "小数部", "cut decimal"]):
                                is_rule = True
                        if is_rule:
                            out_val = process_output_logic("", rule_str, field, is_output_fixed_length)
                        else:
                            fixed_val_str = str(int(rule_desc)) if isinstance(rule_desc, float) and rule_desc.is_integer() else rule_str
                            out_val = process_output_logic(fixed_val_str, "", field, is_output_fixed_length)
                        if out_val is not None:
                            sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                            sheet.cells(row_idx, field["col_idx"]).value = out_val
                    else:
                        in_row_idx = input_pattern_map.get(pattern_key)
                        in_val = sheet.cells(in_row_idx, field["col_idx"]).value if in_row_idx else ""
                        out_val = process_output_logic(in_val, rule_desc, field, is_output_fixed_length)
                        if out_val is not None:
                            sheet.cells(row_idx, field["col_idx"]).number_format = '@'
                            sheet.cells(row_idx, field["col_idx"]).value = out_val
    process_block(TEMPLATE_CONFIG["INPUT"], "INPUT")
    process_block(TEMPLATE_CONFIG["OUTPUT"], "OUTPUT")

def create_data_template(if_id_target, cols_input='', cols_output='', cols_check='', is_fixed_length=True, custom_testcase_path=None):
    # --- CẤU HÌNH ĐƯỜNG DẪN ---
    source_path = os.path.abspath('01.要件定義_インターフェース一覧（STEP3）.xlsx')
    design_docs_dir = r'D:\Project\151_ISA_AsteriaWrap\trunk\99_FromJP\10_プロジェクト資材\02.IFAgreement\03.確定'
    output_dir = os.getcwd()

    if not os.path.exists(source_path):
        print(f"Lỗi: Không tìm thấy file Master: {source_path}")
        return

    print(f"--- 1. Đang đọc dữ liệu Master ---")
    df = pd.read_excel(source_path, sheet_name='【STEP3】インターフェース一覧', skiprows=20, header=None)

    # Tìm thông tin IF_ID
    target_row = None
    for index, row in df.iterrows():
        if_ag_id = str(row[65]).strip() if pd.notnull(row[65]) else None
        if_id = str(row[2]).strip() if pd.notnull(row[2]) else None
        
        if if_ag_id == if_id_target or if_id == if_id_target:
            target_row = row
            break

    if target_row is None:
        print(f"Lỗi: Không tìm thấy Interface ID '{if_id_target}' trong file Master.")
        return
        
    actual_if_id = str(target_row[2]).strip()
    if_ag_id = str(target_row[65]).strip()
    if_h = str(target_row[7]).strip() if pd.notnull(target_row[7]) else ""

    # Đọc cột X (index 23) từ master để quyết định Input là Fixed-length hay không
    # Giá trị này sẽ ghi đè tham số --no-fixed từ command line
    x_val = str(target_row[23]).strip() if pd.notnull(target_row[23]) else ""
    is_input_fixed_length = (x_val == "固定長") if x_val else is_fixed_length
    print(f" -> Chế độ Input Fixed-Length: {is_input_fixed_length} (Dựa trên cột X='{x_val}')")

    ap_val = str(target_row[41]).strip() if pd.notnull(target_row[41]) else ""
    is_output_fixed_length = (ap_val == "固定長") if ap_val else is_fixed_length
    print(f" -> Chế độ Output Fixed-Length: {is_output_fixed_length} (Dựa trên cột AP='{ap_val}')")

    print(f"--- 2. Đang tìm file thiết kế cho IF: {actual_if_id} ---")
    existing_design_files = os.listdir(design_docs_dir) if os.path.exists(design_docs_dir) else []
    matches = [f for f in existing_design_files if (if_ag_id in f or actual_if_id in f) and not f.startswith('~$')]
    found_design = sorted(matches, reverse=True)[0] if matches else None

    if not found_design:
        print(f"Lỗi: Không tìm thấy file thiết kế nào chứa '{if_ag_id}' hoặc '{actual_if_id}' trong thư mục {design_docs_dir}")
        return

    design_path = os.path.join(design_docs_dir, found_design)
    print(f" -> Đã tìm thấy: {found_design}")

    # --- SETUP CỘT MAPPING TỪ THIẾT KẾ ---
    cols_map = {
        'L_SEQ': 'B', 'L_NAME': 'C', 'L_TYPE': 'I', 'L_LEN': 'J', 'L_DEC': 'K', 'L_MINUS': 'L', 'L_SAMPLE': 'D', 'L_REQ': 'F',
        'R_SEQ': 'U', 'R_NAME': 'V', 'R_TYPE': 'AB', 'R_LEN': 'AC', 'R_DEC': 'AD', 'R_MINUS': 'AE', 'R_REQ': 'Y',
        'C_MAP1': 'AH', 'C_MAP2': 'AL', 'C_MAP3': 'AJ'
    }
    if cols_input:
        for k, v in zip(['L_SEQ', 'L_NAME', 'L_TYPE', 'L_LEN', 'L_DEC', 'L_MINUS', 'L_SAMPLE', 'L_REQ'], [x.strip() for x in cols_input.split(',') if x.strip()]):
            cols_map[k] = v.upper()
    if cols_output:
        for k, v in zip(['R_SEQ', 'R_NAME', 'R_TYPE', 'R_LEN', 'R_DEC', 'R_MINUS', 'R_REQ'], [x.strip() for x in cols_output.split(',') if x.strip()]):
            cols_map[k] = v.upper()
    if cols_check:
        for k, v in zip(['C_MAP1', 'C_MAP2', 'C_MAP3'], [x.strip() for x in cols_check.split(',') if x.strip()]):
            cols_map[k] = v.upper()

    c_idx = {k: col2idx(v) for k, v in cols_map.items()}

    print(f"--- 3. Đang đọc dữ liệu Mapping (xlwings) ---")
    app = xw.App(visible=False)
    app.display_alerts = False

    mapping_data = []
    wb_src = None
    try:
        wb_src = app.books.open(design_path, update_links=False, read_only=True)
        
        # Tìm sheet マッピング定義
        ss_copy = None
        if if_h:
            ss_copy = next((s for s in wb_src.sheets if f"マッピング定義_{if_h}" in s.name), None)
        if not ss_copy:
            ss_copy = next((s for s in wb_src.sheets if "マッピング定義" in s.name and "_" not in s.name), None)
        if not ss_copy:
            ss_copy = next((s for s in wb_src.sheets if "マッピング定義" in s.name), None)

        if not ss_copy:
            print("Lỗi: Không tìm thấy sheet 'マッピング定義' trong file thiết kế.")
            return

        data_block = ss_copy.range('A8:BZ1000').value
        last_row_idx = 0
        
        for i, row_data in enumerate(data_block):
            def get_val(row, k): 
                idx = c_idx.get(k)
                return row[idx] if idx is not None and idx >= 0 and len(row) > idx else None
            
            v_in_seq, v_in_name, v_in_type = get_val(row_data, 'L_SEQ'), get_val(row_data, 'L_NAME'), get_val(row_data, 'L_TYPE')
            v_in_len, v_in_dec, v_in_minus, v_in_sample, v_in_req = get_val(row_data, 'L_LEN'), get_val(row_data, 'L_DEC'), get_val(row_data, 'L_MINUS'), get_val(row_data, 'L_SAMPLE'), get_val(row_data, 'L_REQ')
            
            
            v_out_seq, v_out_name, v_out_type = get_val(row_data, 'R_SEQ'), get_val(row_data, 'R_NAME'), get_val(row_data, 'R_TYPE')
            v_out_len, v_out_dec, v_out_minus, v_out_req = get_val(row_data, 'R_LEN'), get_val(row_data, 'R_DEC'), get_val(row_data, 'R_MINUS'), get_val(row_data, 'R_REQ')
            
            
            map_ah, map_al, map_aj = get_val(row_data, 'C_MAP1'), get_val(row_data, 'C_MAP2'), get_val(row_data, 'C_MAP3')

            if is_numeric(v_in_seq) or is_numeric(v_out_seq):
                # Lấy trực tiếp giá trị từ cột thứ 3 trong colsCheck (mặc định là C_MAP3 -> AJ)
                combined_rule = str(map_aj).strip() if map_aj is not None and str(map_aj).strip() not in ["", "None", "-"] else ""

                mapping_data.append({
                    'in_seq': v_in_seq, 'in_name': v_in_name, 'in_type': v_in_type, 'in_len': v_in_len, 'in_dec': v_in_dec, 'in_minus': v_in_minus, 'in_sample': v_in_sample, 'in_req': v_in_req,
                    'out_seq': v_out_seq, 'out_name': v_out_name, 'out_type': v_out_type, 'out_len': v_out_len, 'out_dec': v_out_dec, 'out_minus': v_out_minus, 'out_req': v_out_req,
                    'rule': combined_rule
                })
                last_row_idx = i
            elif any(v is not None for v in [v_in_seq, v_in_name, v_out_seq, v_out_name]): 
                last_row_idx = i
                
            if i > last_row_idx + 20: break # Dừng nếu quá 20 dòng trống liên tiếp
            
    finally:
        if wb_src: wb_src.close()

    if not mapping_data:
        print("Không tìm thấy field mapping nào hợp lệ.")
        app.quit()
        return

    # Sắp xếp mapping_data theo thứ tự Output SEQ
    def sort_by_out_seq(item):
        seq = item.get('out_seq')
        if is_numeric(seq):
            return (0, float(seq))
        return (1, str(seq) if seq else "")
        
    mapping_data.sort(key=sort_by_out_seq)

    print(f" -> Đã lấy được {len(mapping_data)} fields.")
    
    # --- 3.5. ĐỌC TESTCASE ĐỂ LẤY PATTERN INPUT NẾU CÓ ---
    if custom_testcase_path and os.path.exists(custom_testcase_path):
        testcase_path = custom_testcase_path
        testcase_file = os.path.basename(custom_testcase_path)
        testcase_dir = os.path.dirname(custom_testcase_path)
    else:
        testcase_dir = os.path.join(output_dir, 'testcases', actual_if_id)
        testcase_file = f"単体テスト仕様書兼成績書_{if_ag_id}_{actual_if_id}.xlsx"
        testcase_path = os.path.join(testcase_dir, testcase_file)
    
    testcase_data = {}
    if os.path.exists(testcase_path):
        print(f"--- 3.5. Đang đọc Testcase để lấy pattern Input ({testcase_file}) ---")
        wb_tc = None
        try:
            wb_tc = app.books.open(testcase_path, update_links=False, read_only=True)
            s_tc = next((s for s in wb_tc.sheets if "テスト計画書兼結果報告書(マッピング)" in s.name), None)
            if s_tc:
                tc_data = s_tc.range('H2:BZ20').value
                if tc_data and isinstance(tc_data, list):
                    for c_idx in range(len(tc_data[0])): 
                        in_seq = tc_data[0][c_idx]
                        in_name = tc_data[1][c_idx]
                        if in_seq is not None and str(in_seq).strip() != "":
                            key = f"{safe_str(in_seq)}_{safe_str(in_name)}"
                            testcase_data[key] = {}
                            for r_idx in range(8, min(19, len(tc_data))): # Row 10 (index 8) to 20
                                testcase_data[key][r_idx + 2] = tc_data[r_idx][c_idx]
        except Exception as e:
            print(f"Cảnh báo: Không thể đọc file Testcase: {e}")
        finally:
            if wb_tc: wb_tc.close()
    else:
        print(f" -> Không tìm thấy file testcase ({testcase_file}), sẽ dùng giá trị mặc định toàn bộ.")

    print(f"--- 4. Đang tạo file Template Output ---")

    output_file_name = f"template_create_data_{actual_if_id}.xlsx"
    output_file_path = os.path.join(output_dir, output_file_name)
    template_path = os.path.join(output_dir, "template_create_data.xlsx")

    try:
        has_template = os.path.exists(template_path)
        if has_template:
            print(f" -> Sử dụng template: {template_path}")
            shutil.copy(template_path, output_file_path)
            wb_new = app.books.open(output_file_path)
            sheet = wb_new.sheets[0]
        else:
            print(" -> Không tìm thấy 'template_create_data.xlsx', sẽ tạo file mới.")
            wb_new = app.books.add()
            sheet = wb_new.sheets[0]
            sheet.name = "Template_Data"

        max_rows = 60
        INPUT_PATTERNS = { 9: "MAX_LEN", 10: "MIN_LEN", 11: "OVER_LEN", 12: "SYMBOL_MIX", 13: "INVALID_TYPE", 14: "ZENKAKU_MIX", 15: "EMPTY", 16: "ZERO", 17: "POSITIVE", 18: "NEGATIVE", 19: "DECIMAL", 20: "SAMPLE" }
        OUTPUT_PATTERNS = { 38: "MAX_LEN", 39: "MIN_LEN", 40: "OVER_LEN", 41: "SYMBOL_MIX", 42: "INVALID_TYPE", 43: "ZENKAKU_MIX", 44: "EMPTY", 45: "ZERO", 46: "POSITIVE", 47: "NEGATIVE", 48: "DECIMAL", 49: "SAMPLE" }

        if not has_template:
            header_matrix = [[None for _ in range(3)] for _ in range(max_rows)]
            header_matrix[3][1] = "INPUT"; header_matrix[3][2] = "SEQ"
            header_matrix[4][2] = "TYPE"; header_matrix[5][2] = "LENGTH"
            header_matrix[6][2] = "DECIMAL"; header_matrix[8][2] = "NAME"
            for r_idx, pat in INPUT_PATTERNS.items(): header_matrix[r_idx][2] = pat
            
            header_matrix[31][1] = "OUTPUT"; header_matrix[31][2] = "SEQ"
            header_matrix[32][2] = "TYPE"; header_matrix[33][2] = "LENGTH"
            header_matrix[34][2] = "DECIMAL"; header_matrix[36][2] = "NAME"
            header_matrix[37][2] = "RULE"
            for r_idx, pat in OUTPUT_PATTERNS.items(): header_matrix[r_idx][2] = pat
            
            sheet.range((1, 1), (max_rows, 3)).number_format = '@'
            sheet.range((1, 1), (max_rows, 3)).value = header_matrix
            
            sheet.range('C4:C9').color = (255, 242, 204)
            sheet.range('C32:C38').color = (226, 239, 218)
            sheet.range('B4').font.bold = True
            sheet.range('B32').font.bold = True
            sheet.range('C:C').autofit()

        # --- ĐIỀN DỮ LIỆU MAPPING VÀO CÁC CỘT (Bắt đầu từ cột D) ---
        data_matrix = [[None for _ in range(len(mapping_data))] for _ in range(max_rows)]
        for i, field in enumerate(mapping_data):
            # INPUT
            data_matrix[3][i] = field['in_seq']
            data_matrix[4][i] = field['in_type']
            data_matrix[5][i] = field['in_len']
            data_matrix[6][i] = field['in_dec']
            data_matrix[8][i] = field['in_name']
            data_matrix[20][i] = safe_str(field['in_sample'])
            
            # Đánh dấu "o" cho Input Patterns dựa trên file Testcase
            key = f"{safe_str(field['in_seq'])}_{safe_str(field['in_name'])}"
            tc_col_data = testcase_data.get(key)
            
            if field['in_name'] and str(field['in_name']).strip() not in ["", "None"]:
                for r_idx in INPUT_PATTERNS.keys():
                    if INPUT_PATTERNS[r_idx] == "SAMPLE":
                        continue
                    if tc_col_data:
                        tc_row = r_idx + 1 # Ánh xạ dòng Template (9..19) -> Testcase (10..20)
                        val = tc_col_data.get(tc_row)
                        if val in ['○', '〇', 'o', 'O']:
                            data_matrix[r_idx][i] = "o"
                    else:
                        data_matrix[r_idx][i] = "o" # Mặc định nếu không tìm thấy file testcase

            # OUTPUT
            data_matrix[31][i] = field['out_seq']
            data_matrix[32][i] = field['out_type']
            data_matrix[33][i] = field['out_len']
            data_matrix[34][i] = field['out_dec']
            data_matrix[36][i] = field['out_name']
            data_matrix[37][i] = field['rule']
            
            # Đánh dấu "o" cho Output Patterns
            if field['out_name'] and str(field['out_name']).strip() not in ["", "None"]:
                for r_idx in OUTPUT_PATTERNS.keys():
                    data_matrix[r_idx][i] = "o"

        # Nhân bản toàn bộ cột D (bao gồm định dạng, độ rộng cột, viền) ra cho đủ số lượng field
        if has_template and len(mapping_data) > 1:
            try:
                sheet.range('D:D').api.Copy()
                target_range = sheet.range((1, 5), (1, 3 + len(mapping_data)))
                target_range.api.PasteSpecial(Paste=-4104) # xlPasteAll
                app.api.CutCopyMode = False
            except Exception as e:
                print(f"Cảnh báo: Không thể nhân bản cột D: {e}")

        sheet.range((1, 4), (max_rows, 3 + len(mapping_data))).number_format = '@'
        sheet.range((1, 4), (max_rows, 3 + len(mapping_data))).value = data_matrix
        
        # Tô màu từng ô Pattern: Xám nếu không có 'o', Trắng nếu có 'o'
        gray_color = (191, 191, 191)
        white_color = (255, 255, 255)
        for i, field in enumerate(mapping_data):
            col_idx = 4 + i
            
            # Input patterns
            for r_idx in INPUT_PATTERNS.keys():
                val = data_matrix[r_idx][i]
                if INPUT_PATTERNS[r_idx] == "SAMPLE":
                    sheet.cells(r_idx + 1, col_idx).color = white_color
                else:
                    sheet.cells(r_idx + 1, col_idx).color = white_color if val in ['o', 'O', '〇', '○'] else gray_color
                
            # Output patterns
            for r_idx in OUTPUT_PATTERNS.keys():
                val = data_matrix[r_idx][i]
                sheet.cells(r_idx + 1, col_idx).color = white_color if val in ['o', 'O', '〇', '○'] else gray_color
                
            # Bôi xám khu vực Header nếu Name bị trống
            if not field['in_name'] or str(field['in_name']).strip() in ["", "None"]:
                sheet.range((4, col_idx), (9, col_idx)).color = gray_color
            if not field['out_name'] or str(field['out_name']).strip() in ["", "None"]:
                sheet.range((32, col_idx), (38, col_idx)).color = gray_color

        print(f"--- 5. Đang fill Mock Data vào file Template ---")
        fill_template_data(sheet, mapping_data, is_input_fixed_length, is_output_fixed_length)

        wb_new.save(output_file_path)
        print(f"--- HOÀN TẤT ---")
        print(f"File Template kèm Mock Data đã được tạo thành công tại: {output_file_path}")

    except Exception as e:
        print(f"Lỗi khi tạo file excel: {e}")
    finally:
        try: wb_new.close()
        except: pass
        app.quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tạo file Template chuẩn từ IF ID")
    parser.add_argument("if_id", help="Interface ID (Ví dụ: I1110 hoặc IF_AG_ID)")
    parser.add_argument('--colsInput', type=str, default='', help="Cột Input theo thứ tự: SEQ,NAME,TYPE,LEN,DEC,MINUS,SAMPLE,REQ")
    parser.add_argument('--colsOutput', type=str, default='', help="Cột Output theo thứ tự: SEQ,NAME,TYPE,LEN,DEC,MINUS,REQ")
    parser.add_argument('--colsCheck', type=str, default='', help="Cột Check chung: MAP1,MAP2,MAP3")
    parser.add_argument('--no-fixed', action='store_true', help="Tắt chế độ tự động Padding độ dài cho Input (dùng cho Variable-length)")
    parser.add_argument('--testcase', type=str, default=None, help="Đường dẫn đến file testcase (tùy chọn)")
    args = parser.parse_args()
    
    create_data_template(args.if_id, args.colsInput, args.colsOutput, args.colsCheck, not args.no_fixed, args.testcase)