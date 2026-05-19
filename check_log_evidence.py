# -*- coding: utf-8 -*-
import openpyxl
import argparse
import os
import datetime
import re

try:
    import pytesseract
    from PIL import Image
    import io
    
    # Cấu hình đường dẫn cài đặt Tesseract-OCR trên Windows
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

def get_result_sheet_name(evidence_sheet_name):
    if "共通" in evidence_sheet_name:
        return "テスト計画書兼結果報告書(共通)"
    elif "個別" in evidence_sheet_name:
        return "テスト計画書兼結果報告書(個別)"
    return None

def clean_text(text):
    """Làm sạch khoảng trắng để so sánh header chính xác hơn"""
    if text is None: return ""
    return str(text).replace(' ', '').replace('　', '').strip()

def check_log_evidence(file_path):
    if not os.path.exists(file_path):
        print(f"Lỗi: Không tìm thấy file:\n{file_path}")
        return

    print(f"Đang đọc dữ liệu từ file: {os.path.basename(file_path)} ...\n")
    try:
        # Dùng data_only=True để lấy giá trị thực thay vì công thức
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"Lỗi khi mở file Excel: {e}")
        return

    print(f"================ TRẠNG THÁI OCR ================\n{'✅ Đã bật (Sẵn sàng đọc ảnh)' if HAS_OCR else '⚠️ Đã tắt (Vui lòng cài pytesseract, Pillow và Tesseract-OCR)'}\n")

    print("================ KIỂM TRA TRẠNG THÁI MỞ FILE ================")
    active_sheet = wb.active
    first_sheet = wb.worksheets[0]
    
    if active_sheet != first_sheet:
        print(f"❌ LỖI: File đang mở ở sheet '{active_sheet.title}' (yêu cầu sheet đầu tiên '{first_sheet.title}').")
    else:
        print(f"✅ File đang mở ở sheet đầu tiên '{first_sheet.title}'.")
        
    try:
        active_cell = "Unknown"
        if active_sheet.views and active_sheet.views.sheetView:
            selections = active_sheet.views.sheetView[0].selection
            if selections and len(selections) > 0:
                active_cell = selections[0].activeCell
            else:
                active_cell = "A1" # Mặc định nếu file không lưu thông tin selection
                
        if active_cell != "A1":
            print(f"❌ LỖI: Con trỏ chuột đang nằm ở ô '{active_cell}' (yêu cầu ô 'A1').")
        else:
            print(f"✅ Con trỏ chuột đang nằm ở ô 'A1'.")
    except Exception:
        pass
    print("")

    # Pre-load result sheets into memory for quick lookup
    result_sheets = {}
    for rs_name in ["テスト計画書兼結果報告書(共通)", "テスト計画書兼結果報告書(個別)"]:
        if rs_name in wb.sheetnames:
            rs = wb[rs_name]
            result_sheets[rs_name] = rs
            
            # Kiểm tra lỗi Cột C có dữ liệu nhưng Cột K trống
            empty_k_rows = []
            for row_idx, row in enumerate(rs.iter_rows(values_only=True), start=1):
                val_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                col_k_val = row[10] if len(row) > 10 else None
                
                # Bỏ qua dòng tiêu đề (thường chứa chữ '項目')
                if val_c and "項目" not in val_c and (col_k_val is None or str(col_k_val).strip() == ""):
                    empty_k_rows.append(row_idx)
                    
            if empty_k_rows:
                print(f"================ KIỂM TRA: {rs_name} ================")
                print(f"❌ LỖI: Cột C có dữ liệu nhưng Cột K (số lần test) bị trống tại các dòng: {', '.join(map(str, empty_k_rows))}\n")

    target_sheets = ['エビデンス(共通)', 'エビデンス(個別)']
    found_any_sheet = False

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        found_any_sheet = True
        ws = wb[sheet_name]
        print(f"================ SHEET: {sheet_name} ================")
        
        result_sheet_name = get_result_sheet_name(sheet_name)
        rs = result_sheets.get(result_sheet_name)

        images_by_row = {}
        for img in getattr(ws, '_images', []):
            try:
                if hasattr(img.anchor, '_from'):
                    r = img.anchor._from.row + 1
                    images_by_row.setdefault(r, []).append(img)
            except Exception:
                pass
                
        print(f"  -> Tìm thấy {sum(len(v) for v in images_by_row.values())} ảnh trong sheet này.")

        current_group = "Unknown_Group"
        group_start_row = 1
        in_table = False
        col_idx_gyo = -1
        col_idx_nichiji = -1
        col_idx_msg = -1
        
        group_data = []

        def flush_group():
            if group_data:
                group_images = []
                log_start_row = group_data[0]['row_idx']
                for r in range(group_start_row, log_start_row):
                    if r in images_by_row:
                        group_images.extend(images_by_row[r])
                analyze_group(current_group, group_data, rs, result_sheet_name, group_images)
                group_data.clear()

        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            col_a_val = row[0]
            
            # Nhận diện nhóm chạy mới (Có giá trị ở cột A)
            if col_a_val is not None and str(col_a_val).strip() != "":
                if in_table and group_data:
                    flush_group()
                    in_table = False
                current_group = str(col_a_val).strip()
                group_start_row = row_idx

            # Đang tìm kiếm Header của bảng Log
            if not in_table:
                row_strs = [clean_text(cell) for cell in row]
                if "行" in row_strs and "日時" in row_strs and "レベル" in row_strs:
                    col_idx_gyo = row_strs.index("行")
                    col_idx_nichiji = row_strs.index("日時")
                    col_idx_msg = row_strs.index("メッセージ") if "メッセージ" in row_strs else -1
                    in_table = True
                    continue
                else:
                    # Thử nhận diện dòng log trực tiếp (trường hợp copy log không có header)
                    # Cấu trúc: [Số 行] -> [Thời gian 日時] -> [INFO/WARN/ERROR レベル] -> [Message]
                    found_gyo, found_nichiji, found_level = -1, -1, -1
                    for c_idx, cell_val in enumerate(row):
                        if cell_val is None: continue
                        val_str = str(cell_val).strip()
                        if not val_str: continue

                        if found_gyo == -1:
                            if re.match(r'^\d+(\.0)?$', val_str):
                                found_gyo = c_idx
                        elif found_nichiji == -1:
                            # Hỗ trợ cả định dạng chuỗi "2026-05-18T11:07:13" lẫn kiểu datetime của Excel
                            if isinstance(cell_val, (datetime.datetime, datetime.time)) or re.search(r'\d{2}:\d{2}:\d{2}', val_str):
                                found_nichiji = c_idx
                        elif found_level == -1:
                            if val_str in ["INFO", "WARN", "ERROR", "DEBUG", "TRACE", "FATAL"]:
                                found_level = c_idx
                                break
                                
                    if found_gyo != -1 and found_nichiji != -1 and found_level != -1:
                        col_idx_gyo = found_gyo
                        col_idx_nichiji = found_nichiji
                        
                        # Tìm cột chứa Message (thường là cột xa nhất có chứa text phía sau cột Level)
                        col_idx_msg = -1
                        for c_idx in range(len(row) - 1, found_level, -1):
                            if row[c_idx] is not None and str(row[c_idx]).strip() != "":
                                col_idx_msg = c_idx
                                break
                        if col_idx_msg == -1:
                            col_idx_msg = found_level + 1
                            
                        in_table = True
                        # Không dùng `continue` để code đi tiếp xuống khối `if in_table` bên dưới và xử lý luôn dòng này

            # Đang đọc dữ liệu trong bảng Log
            if in_table:
                # Nếu độ dài của dòng không đủ để lấy cột đã lưu, nghĩa là kết thúc bảng
                if len(row) <= max(col_idx_gyo, col_idx_nichiji):
                    flush_group()
                    in_table = False
                    continue

                gyo_val = row[col_idx_gyo]
                nichiji_val = row[col_idx_nichiji]

                # Điều kiện kết thúc bảng: Cột "行" bị trống hoặc không phải là số
                if gyo_val is None or str(gyo_val).strip() == "":
                    flush_group()
                    in_table = False
                    continue

                try:
                    gyo_num = int(float(str(gyo_val).strip()))
                    msg_val = row[col_idx_msg] if col_idx_msg != -1 and col_idx_msg < len(row) else ""
                    group_data.append({
                        'row_idx': row_idx,
                        'gyo': gyo_num,
                        'nichiji': nichiji_val,
                        'message': msg_val
                    })
                except ValueError:
                    # Không convert được sang số -> Hết bảng
                    flush_group()
                    in_table = False

        # Quét xong sheet, nếu còn data chưa phân tích thì phân tích nốt
        flush_group()
            
    if not found_any_sheet:
        print(f"Lỗi: Không tìm thấy sheet nào có tên 'エビデンス(共通)' hoặc 'エビデンス(個別)' trong file này.")

def analyze_group(group_name, data, rs, rs_name, group_images=None):
    if group_images is None:
        group_images = []

    if not data:
        return

    print(f"\n[Nhóm Test]: {group_name}")
    print(f"  - Tổng số dòng log: {len(data)}")

    # 1. Check xem cột "行" có tăng dần không
    is_sorted = True
    unsorted_details = []
    for i in range(len(data) - 1):
        if data[i]['gyo'] >= data[i+1]['gyo']:
            is_sorted = False
            unsorted_details.append(f"Dòng Excel {data[i]['row_idx']} (行: {data[i]['gyo']}) -> Dòng Excel {data[i+1]['row_idx']} (行: {data[i+1]['gyo']})")

    if is_sorted:
        print("  - Trạng thái Sort (行): ✅ Đã sắp xếp tăng dần hợp lệ.")
    else:
        print("  - Trạng thái Sort (行): ❌ LỖI - Chưa tăng dần!")
        for detail in unsorted_details:
            print(f"      + Lệch tại: {detail}")

    # 2. Xử lý thời gian test
    raw_dates = set()
    for item in data:
        val = item['nichiji']
        if val is None: continue
        
        if isinstance(val, datetime.datetime):
            date_str = val.strftime("%Y/%m/%d")
        else:
            # Nếu openpyxl đọc ra dưới dạng chuỗi (String)
            val_str = str(val).strip()
            # Thay thế chữ 'T' bằng khoảng trắng (nếu có) rồi mới cắt chuỗi để loại bỏ hoàn toàn phần thời gian
            date_part = val_str.replace("T", " ").split(" ")[0]
            date_str = date_part.replace("-", "/")
        raw_dates.add(date_str)

    sorted_dates = sorted(list(raw_dates))
    if len(sorted_dates) == 1:
        print(f"  - Ngày chạy log: {sorted_dates[0]}")
    else:
        print(f"  - Ngày chạy log: Từ {sorted_dates[0]} đến {sorted_dates[-1]}")

    # 3. Check xem có sự khác nhau ở Date không (Nhiều lần test khác ngày)
    if len(raw_dates) > 1:
        print(f"  - CẢNH BÁO DATE: ⚠️ Cụm này chứa log của nhiều ngày khác nhau ({', '.join(sorted_dates)}). Có vẻ cụm này đã được chạy check nhiều lần!")

    # Kiểm tra thời gian từ ảnh (OCR)
    if not group_images:
        print("  - [Ảnh Status Panel]: ❌ LỖI: Có log chạy nhưng không tìm thấy ảnh (Status Panel) nào ở trên bảng log!")
    elif HAS_OCR:
        panel_start_time = None
        panel_duration_ms = None
        for idx, img in enumerate(group_images):
            try:
                image_data = img._data() if callable(getattr(img, '_data', None)) else getattr(img, '_data', None)
                if image_data:
                    pil_img = Image.open(io.BytesIO(image_data))
                    
                    # Tiền xử lý ảnh: Chuyển sang ảnh xám
                    pil_img = pil_img.convert('L')
                    width, height = pil_img.size
                    
                    # Phóng to toàn bộ ảnh lên 2 lần để nét chữ rõ hơn
                    full_img = pil_img.resize((width * 2, height * 2), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                    
                    # Cách 1: Đọc TOÀN BỘ ảnh với PSM 11 (Sparse text - chuyên tìm chữ rải rác trên màn hình UI)
                    text = pytesseract.image_to_string(full_img, lang='jpn+eng', config='--psm 11')
                    
                    # Regex thông minh: Tìm Thời gian và tuỳ chọn lấy thêm số ms phía sau (Dùng re.DOTALL để quét qua dấu xuống dòng)
                    time_regex = r'(\d{1,2})\s*[:：]\s*(\d{2})\s*[:：]\s*(\d{2})(?:.*?(\d+)\s*ms)?'
                    m = re.search(time_regex, text, re.IGNORECASE | re.DOTALL)
                    
                    if not m:
                        # Cách 2: Đọc TOÀN BỘ ảnh với PSM 3 (Chế độ tự phân tích layout mặc định của Tesseract)
                        text_fallback = pytesseract.image_to_string(full_img, lang='jpn+eng', config='--psm 3')
                        m2 = re.search(time_regex, text_fallback, re.IGNORECASE | re.DOTALL)
                        if m2:
                            text = text_fallback
                            m = m2
                            
                    if not m:
                        # Cách 3: Ép TOÀN BỘ ảnh thành Trắng/Đen (Binarization) để khử nhiễu nền
                        threshold = 150
                        bw_img = full_img.point(lambda p: p > threshold and 255)
                        text_bw = pytesseract.image_to_string(bw_img, lang='jpn+eng', config='--psm 11')
                        m3 = re.search(time_regex, text_bw, re.IGNORECASE | re.DOTALL)
                        if m3:
                            text = text_bw
                            m = m3

                    if m:
                        print(m.group(4))
                        # Định dạng lại thành chuẩn HH:MM:SS
                        panel_start_time = f"{int(m.group(1)):02d}:{m.group(2)}:{m.group(3)}"
                        panel_duration_ms = m.group(4)
                        break
            except Exception as e:
                print(f"      ⚠️ Lỗi khi xử lý ảnh (OCR): {e}")
        
        if not panel_start_time:
            print("  - [Ảnh Status Panel]: ⚠️ Không tìm thấy hoặc không đọc được thời gian '実行開始:' từ ảnh.")
        
        if panel_start_time:
            print(f"  - [Ảnh Status Panel]: Tìm thấy giờ bắt đầu là {panel_start_time}")
            first_log_val = data[0]['nichiji']
            if first_log_val:
                first_log_time_str = ""
                if isinstance(first_log_val, datetime.datetime):
                    first_log_time_str = first_log_val.strftime("%H:%M:%S")
                else:
                    time_m = re.search(r'(\d{1,2}:\d{2}:\d{2})', str(first_log_val))
                    if time_m:
                        first_log_time_str = time_m.group(1)
                
                if first_log_time_str:
                    try:
                        fmt = "%H:%M:%S"
                        t_panel = datetime.datetime.strptime(panel_start_time, fmt).time()
                        t_log = datetime.datetime.strptime(first_log_time_str, fmt).time()
                        
                        if t_log >= t_panel:
                            print(f"      ✅ OK: Thời gian log ({first_log_time_str}) >= Thời gian ảnh ({panel_start_time})")
                        else:
                            print(f"      ❌ LỖI: Thời gian log ({first_log_time_str}) < Thời gian ảnh ({panel_start_time})")
                    except ValueError:
                        print(f"      ⚠️ Không thể so sánh thời gian: Log ({first_log_time_str}), Ảnh ({panel_start_time})")

        if panel_duration_ms and data:
            print(data)
            log_duration_ms = None
            # Quét ngược từ dưới lên để tìm chính xác dòng log kết thúc subflow chứa duration
            for item in reversed(data):
                msg = str(item.get('message', ''))
                if "サブフローの実行が終了しました" in msg:
                    m_log_dur = re.search(r'\[\s*(\d+)(?:\s*ms)?\s*\]', msg, re.IGNORECASE)
                    if m_log_dur:
                        log_duration_ms = m_log_dur.group(1)
                        break

            if log_duration_ms:
                try:
                    if int(log_duration_ms) < int(panel_duration_ms):
                        print(f"      ✅ OK: Duration Log Main ({log_duration_ms}ms) < Duration Ảnh ({panel_duration_ms}ms)")
                    else:
                        print(f"      ❌ LỖI: Duration Log Main ({log_duration_ms}ms) >= Duration Ảnh ({panel_duration_ms}ms)")
                except ValueError:
                    pass
            else:
                print(f"      ⚠️ Không tìm thấy dòng log 'サブフローの実行が終了しました' chứa thời gian '[Nms]' hoặc '[N]' để so sánh Duration.")

    elif not HAS_OCR:
        print("  - [Ảnh Status Panel]: ⚠️ Cần cài đặt thư viện 'pytesseract' và 'Pillow' để đọc text từ ảnh.")

    # Bóc tách số lần chạy từ tên nhóm (Ví dụ: ■1回目 -> 1)
    m_run = re.search(r'(\d+)回目', group_name)
    extracted_run = None
    base_group_name = group_name
    if m_run:
        extracted_run = int(m_run.group(1))
        base_group_name = re.sub(r'^■?\s*\d+回目\s*', '', group_name).strip()

    # 4. Check result sheet
    if rs:
        found_rows = []
        last_b, last_c, last_d = "", "", ""
        
        for row_idx, row in enumerate(rs.iter_rows(values_only=True), start=1):
            val_b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            val_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
            val_d = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
            
            if val_b: last_b = val_b
            if val_c: last_c = val_c
            if val_d: last_d = val_d
            
            # Yêu cầu: Chỉ check nếu cột C (項目) của dòng đó có giá trị
            if not val_c:
                continue
                
            if group_name in [val_b, val_c, val_d]:
                found_rows.append(row_idx)
                continue
                
            if base_group_name and base_group_name in [val_b, val_c, val_d]:
                found_rows.append(row_idx)
                continue
                
            tc_id = f"{last_b}-{last_c}-{last_d}"
            if last_b and last_c and last_d and (group_name == tc_id or base_group_name == tc_id):
                found_rows.append(row_idx)
                continue
                
        # Fallback (Dự phòng): Nếu tìm theo tên không ra, tìm theo số lần chạy ở cột K (áp dụng cho cả 共通 và 個別)
        if not found_rows:
            last_c_fallback = "" # Xử lý cho trường hợp cột C bị merge
            target_run_num = extracted_run
            
            for row_idx, row in enumerate(rs.iter_rows(values_only=True), start=1):
                val_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                if val_c:
                    last_c_fallback = val_c

                col_k_val = row[10] if len(row) > 10 else None
                
                if last_c_fallback and col_k_val is not None and str(col_k_val).strip() != "":
                    # Bóc tách số từ cột K (Ví dụ: "1" -> 1, "1回目" -> 1)
                    m_run = re.search(r'\d+', str(col_k_val))
                    run_num = int(m_run.group()) if m_run else None

                    if run_num is not None:
                        if target_run_num is not None:
                            if run_num == target_run_num:
                                found_rows.append(row_idx)
                        else:
                            target_run_num = run_num
                            found_rows.append(row_idx)

        if found_rows:
            print(f"  - [Đối chiếu {rs_name}]: Tìm thấy {len(found_rows)} dòng ({', '.join(map(str, found_rows))})")
            
            invalid_k_rows = []
            invalid_g_rows = []
            empty_g_rows = []
            
            for f_row in found_rows:
                row_data = list(rs.iter_rows(min_row=f_row, max_row=f_row, values_only=True))[0]
                col_g_date = row_data[6] if len(row_data) > 6 else None
                col_i_date = row_data[8] if len(row_data) > 8 else None
                col_k_runs = row_data[10] if len(row_data) > 10 else None
                
                # Check số lần chạy (Cột K)
                expected_runs = None
                if col_k_runs is not None and str(col_k_runs).strip() != "":
                    m = re.search(r'\d+', str(col_k_runs))
                    if m:
                        expected_runs = int(m.group())
                        
                if expected_runs is not None and extracted_run is not None:
                    if extracted_run != expected_runs:
                        invalid_k_rows.append((f_row, expected_runs))
                        
                # Check Ngày (Ưu tiên cột I, nếu trống lấy cột G)
                date_source_val = col_i_date
                source_date_col_name = "I"
                if date_source_val is None or str(date_source_val).strip() == "":
                    date_source_val = col_g_date
                    source_date_col_name = "G"

                expected_date = None
                if date_source_val is not None and str(date_source_val).strip() != "":
                    if isinstance(date_source_val, datetime.datetime):
                        expected_date = date_source_val.strftime("%Y/%m/%d")
                    else:
                        date_str = str(date_source_val).strip()
                        date_part = date_str.replace("T", " ").split(" ")[0]
                        expected_date = date_part.replace("-", "/")
                        
                if expected_date:
                    if expected_date not in sorted_dates:
                        invalid_g_rows.append((f_row, expected_date, source_date_col_name))
                else:
                    empty_g_rows.append(f_row)

            # In kết quả số lần chạy (Cột K)
            if invalid_k_rows:
                for r, exp in invalid_k_rows:
                    print(f"      + Số lần test (Dòng {r}): ❌ LỆCH! Log là lần {extracted_run}, nhưng Cột K ghi {exp}")
            elif extracted_run is not None:
                print(f"      + Số lần test (Cột K): ✅ Tất cả {len(found_rows)} dòng đều khớp (Lần {extracted_run})")
                
            # In kết quả Ngày test (Cột I/G)
            if invalid_g_rows:
                for r, exp_d, col_name in invalid_g_rows:
                    print(f"      + Ngày test (Dòng {r}): ❌ LỆCH! Cột {col_name} ghi '{exp_d}' nhưng Log chạy ngày {', '.join(sorted_dates)}")
            
            if empty_g_rows:
                print(f"      + Ngày test (Cột I/G): ⚠️ Trống tại các dòng: {', '.join(map(str, empty_g_rows))}")
                
            if not invalid_g_rows and len(found_rows) > len(empty_g_rows):
                print(f"      + Ngày test (Cột I/G): ✅ Khớp ngày ({', '.join(sorted_dates)}) cho tất cả các dòng có dữ liệu")
        else:
            print(f"  - [Đối chiếu]: ⚠️ Không tìm thấy testcase '{group_name}' trong sheet {rs_name}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tool kiểm tra Evidence Log (Sort tăng dần, check thời gian) cho file hoặc folder.")
    parser.add_argument("path", help="Đường dẫn file excel hoặc thư mục chứa các file excel.")
    args = parser.parse_args()
    
    input_path = args.path

    if not os.path.exists(input_path):
        print(f"Lỗi: Đường dẫn không tồn tại: {input_path}")
    elif os.path.isfile(input_path):
        # Nếu là một file đơn lẻ, xử lý trực tiếp
        check_log_evidence(input_path)
    elif os.path.isdir(input_path):
        # Nếu là một thư mục, quét đệ quy để tìm file .xlsx
        print(f"--- Đang quét thư mục '{input_path}' để tìm file .xlsx ---")
        excel_files = []
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.endswith('.xlsx') and not file.startswith('~$'):
                    excel_files.append(os.path.join(root, file))
        
        if not excel_files:
            print("Không tìm thấy file .xlsx nào trong thư mục được chỉ định.")
        else:
            print(f"Tìm thấy {len(excel_files)} file. Bắt đầu xử lý...\n")
            for i, file_path in enumerate(excel_files):
                print(f"\n{'='*25} [{i+1}/{len(excel_files)}] ĐANG XỬ LÝ FILE: {os.path.basename(file_path)} {'='*25}")
                check_log_evidence(file_path)