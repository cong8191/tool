# -*- coding: utf-8 -*-
import openpyxl
import argparse
import os
import datetime
import re
import json

# Imports for image processing, needed by both manual Gemini and Tesseract
try:
    from PIL import Image
    import io
    HAS_IMAGE_LIBS = True
except ImportError:
    HAS_IMAGE_LIBS = False

# Flag for manual mode
MANUAL_GEMINI_WEB = True # Đổi thành True để dùng Gemini web miễn phí, False để dùng Tesseract OCR (local)

# Tesseract specific import
try:
    import pytesseract
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
    expected_runs_by_sheet = {}
    for rs_name in ["テスト計画書兼結果報告書(共通)", "テスト計画書兼結果報告書(個別)"]:
        if rs_name in wb.sheetnames:
            rs = wb[rs_name]
            result_sheets[rs_name] = rs
            
            # Kiểm tra lỗi Cột C có dữ liệu nhưng Cột K trống
            empty_k_rows = []
            expected_runs = {}
            last_b, last_c, last_d = "", "", ""
            
            for row_idx, row in enumerate(rs.iter_rows(values_only=True), start=1):
                val_b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
                val_c = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                val_d = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
                
                if val_b:
                    last_b = val_b
                    if not val_c: last_c = ""
                    if not val_d: last_d = ""
                if val_c:
                    last_c = val_c
                    if not val_d: last_d = ""
                if val_d:
                    last_d = val_d
                    
                val_e = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
                val_f = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
                col_k_val = row[10] if len(row) > 10 else None
                
                if last_c and "項目" not in last_c and "No" not in last_c:
                    if col_k_val is not None and str(col_k_val).strip() != "":
                        m_run = re.search(r'\d+', str(col_k_val))
                        if m_run:
                            run_num = int(m_run.group())
                            tc_parts = [p for p in [last_b, last_c, last_d] if p]
                            tc_id = "-".join(tc_parts)
                            expected_runs.setdefault(run_num, []).append(tc_id)
                    else:
                        # Tránh báo nhầm dòng bị merge hoặc dòng trống hoàn toàn
                        # Chỉ báo nếu dòng thực sự có nội dung Test (Cột C, D, E hoặc F có chữ)
                        if val_c:
                            empty_k_rows.append(row_idx)
                            
            if empty_k_rows:
                print(f"================ KIỂM TRA: {rs_name} ================")
                print(f"⚠️ CẢNH BÁO: Dòng có nội dung Test nhưng Cột K (số lần test) bị trống tại: {', '.join(map(str, empty_k_rows))}\n")
            expected_runs_by_sheet[rs_name] = expected_runs

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
        col_idx_level = -1
        col_idx_msg = -1
        
        group_data = []
        found_runs_set = set()

        def flush_group():
            if group_data:
                group_images = []
                log_start_row = group_data[0]['row_idx']
                for r in range(group_start_row, log_start_row):
                    if r in images_by_row:
                        group_images.extend(images_by_row[r])
                # Truyền thêm file_name để check Interface ID
                analyze_group(current_group, group_data, rs, result_sheet_name, group_images, os.path.basename(file_path), found_runs_set)
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
                    col_idx_level = row_strs.index("レベル")
                    col_idx_msg = row_strs.index("メッセージ") if "メッセージ" in row_strs else (col_idx_level + 1)
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
                        col_idx_level = found_level
                        col_idx_msg = found_level + 1
                            
                        in_table = True

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
                    
                    # Gom toàn bộ text từ cột message trở về sau 
                    # (Giải quyết triệt để vấn đề khoảng trắng tab / cột rỗng chèn giữa Level và Message)
                    msg_parts = []
                    start_col = col_idx_msg if col_idx_msg != -1 else (col_idx_level + 1 if col_idx_level != -1 else 3)
                    for c_idx in range(start_col, len(row)):
                        if row[c_idx] is not None and str(row[c_idx]).strip() != "":
                            msg_parts.append(str(row[c_idx]).strip())
                    
                    msg_val = " ".join(msg_parts)
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
            
        # Tổng kết Coverage
        print(f"\n================ TỔNG KẾT COVERAGE CHO SHEET {sheet_name} ================")
        expected_runs = expected_runs_by_sheet.get(result_sheet_name, {})
        if not expected_runs:
            print("  - ⚠️ Không có thông tin nhóm test (số lần chạy) nào được khai báo ở cột K trong Kế hoạch.")
        else:
            missing_runs = []
            for run_num in sorted(expected_runs.keys()):
                # Loại bỏ duplicate TCs để hiển thị đẹp hơn
                tcs = list(dict.fromkeys(expected_runs[run_num]))
                if run_num in found_runs_set:
                    print(f"  - ✅ Nhóm {run_num} (gồm TC: {', '.join(tcs)}): Đã có Log Evidence.")
                else:
                    print(f"  - ❌ LỖI: Nhóm {run_num} (gồm TC: {', '.join(tcs)}): KHÔNG TÌM THẤY Log Evidence!")
                    missing_runs.append(run_num)
            
            if not missing_runs:
                print("  => TOÀN BỘ CÁC NHÓM TEST TRONG KẾ HOẠCH ĐỀU ĐÃ CÓ EVIDENCE!")
        print("")

    if not found_any_sheet:
        print(f"Lỗi: Không tìm thấy sheet nào có tên 'エビデンス(共通)' hoặc 'エビデンス(個別)' trong file này.")

def analyze_group(group_name, data, rs, rs_name, group_images=None, file_name="", found_runs_set=None):
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

    panel_start_time = None
    panel_duration_ms = None
    jobnet_id = None
    interface_id = None

    # Kiểm tra thời gian từ ảnh (OCR)
    if not group_images:
        print("  - [Ảnh Status Panel]: ❌ LỖI: Có log chạy nhưng không tìm thấy ảnh (Status Panel) nào ở trên bảng log!")
    elif MANUAL_GEMINI_WEB and HAS_IMAGE_LIBS:
        prompt = """
Please analyze the provided software dialog and log viewer images (Japanese UI) and extract the following 5 pieces of information.
Return ONLY a valid JSON object matching this exact schema:

1. "start_time": The time value (format HH:MM:SS) located next to "実行開始" or similar in the status bar at the bottom.
2. "duration_ms": The integer number of milliseconds located before "ms" and next to "正常終了" or similar in the status bar.
3. "jobnet_id": The string value for the parameter named "P_JobnetID", "P_JobnetId", or similar.
4. "interface_id": The string value for the parameter named "P_InterfaceID", "P_InterfaceId", or similar.
5. "end_time_log_img": The time value (format HH:MM:SS) in the "日時" column associated with the message "サブフローの実行が終了しました" (or similar ending message) from the log table image.

Search the entire images carefully. If a parameter is present but its value is empty, return "". If it is absolutely missing or unreadable, use null.
Example output:
{"start_time": "16:11:08", "duration_ms": 350, "jobnet_id": "P1J0111@I0111", "interface_id": "RFSH0111", "end_time_log_img": "11:52:13"}
"""
        # Lọc các ảnh đủ lớn (tránh gửi nhầm icon bé xíu lên Gemini)
        valid_images = []
        for img in group_images:
            try:
                img_data = img._data() if callable(getattr(img, '_data', None)) else getattr(img, '_data', None)
                if img_data:
                    p_img = Image.open(io.BytesIO(img_data))
                    if p_img.width >= 100 and p_img.height >= 50:
                        if p_img.mode != 'RGB':
                            p_img = p_img.convert('RGB')
                        valid_images.append(p_img)
            except Exception: pass

        if not valid_images:
            print("  - [Ảnh Status Panel]: ❌ LỖI: Không tìm thấy ảnh nào đủ lớn trong nhóm này.")
        else:
            # Ghép tất cả các ảnh thành 1 ảnh duy nhất (xếp dọc) để tải lên web tiện lợi hơn
            widths, heights = zip(*(i.size for i in valid_images))
            max_width = max(widths)
            total_height = sum(heights)

            merged_img = Image.new('RGB', (max_width, total_height), (255, 255, 255))
            y_offset = 0
            for img in valid_images:
                merged_img.paste(img, (0, y_offset))
                y_offset += img.size[1]
                
            # --- CHẾ ĐỘ THỦ CÔNG DÙNG GEMINI WEB ---
            manual_dir = os.path.join(os.path.dirname(os.path.abspath(file_path)), "gemini_manual_upload")
            os.makedirs(manual_dir, exist_ok=True)
            
            prompt_path = os.path.join(manual_dir, "prompt.txt")
            with open(prompt_path, "w", encoding="utf-8") as f:
                f.write(prompt)
                
            # Xóa các file ảnh cũ để tránh nhầm lẫn với các lượt quét trước
            for fname in os.listdir(manual_dir):
                if fname.endswith(".png"):
                    try: os.remove(os.path.join(manual_dir, fname))
                    except: pass
                
            safe_group = re.sub(r'[^a-zA-Z0-9_\-]', '_', group_name)
            img_path = os.path.join(manual_dir, f"{safe_group}_merged.png")
            merged_img.save(img_path)
                
            print(f"  - [Ảnh Status Panel]: ⚠️ CHẾ ĐỘ THỦ CÔNG ĐANG BẬT!")
            print(f"      -> Đã ghép {len(valid_images)} ảnh thành 1 ảnh duy nhất và xuất ra thư mục: {manual_dir}")
            print(f"      -> VUI LÒNG kéo thả 1 ảnh duy nhất và copy nội dung prompt.txt lên: https://gemini.google.com")
            print(f"      -> Sau đó, copy khối JSON kết quả trả về, dán vào đây và nhấn Enter 2 lần:")
            
            user_lines = []
            while True:
                line = input()
                if line.strip() == "":
                    if user_lines: break
                else:
                    user_lines.append(line)
            
            json_text = "\n".join(user_lines).strip()

            try:
                if json_text.startswith("```json"): json_text = json_text[7:]
                elif json_text.startswith("```"): json_text = json_text[3:]
                if json_text.endswith("```"): json_text = json_text[:-3]
                
                ocr_result = json.loads(json_text.strip())
                
                panel_start_time = ocr_result.get("start_time")
                panel_duration_ms = ocr_result.get("duration_ms")
                jobnet_id = ocr_result.get("jobnet_id")
                interface_id = ocr_result.get("interface_id")
                end_time_log_img = ocr_result.get("end_time_log_img")
            except Exception as e:
                print(f"      ⚠️ Lỗi khi phân tích kết quả JSON: {e}")

    elif HAS_OCR and HAS_IMAGE_LIBS:
        print("  - [Ảnh Status Panel]: ⚠️ Không tìm thấy Gemini API Key. Đang thử dùng Tesseract OCR (độ chính xác thấp hơn)...")
        panel_start_time = None
        panel_duration_ms = None
        end_time_log_img = None
        for idx, img in enumerate(group_images):
            try:
                image_data = img._data() if callable(getattr(img, '_data', None)) else getattr(img, '_data', None)
                if image_data:
                    pil_img = Image.open(io.BytesIO(image_data))
                    
                    # Tiền xử lý ảnh: Chuyển sang ảnh xám
                    pil_img = pil_img.convert('L')
                    width, height = pil_img.size
                    
                    # CHIẾN THUẬT MỚI: Cắt riêng phần thanh Status (thường nằm sát đáy, cao khoảng 80px)
                    bottom_crop = pil_img.crop((0, max(0, height - 80), width, height))
                    
                    # Phóng to ảnh bằng 2 phương pháp (LANCZOS làm mượt, NEAREST giữ nét vuông vức chống nhìn nhầm 5 thành 6)
                    resample_filter = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
                    resample_nearest = Image.Resampling.NEAREST if hasattr(Image, 'Resampling') else Image.NEAREST
                    
                    full_img = pil_img.resize((width * 2, height * 2), resample_filter)
                    bottom_img_lanczos = bottom_crop.resize((width * 2, bottom_crop.size[1] * 2), resample_filter)
                    bottom_img_nearest = bottom_crop.resize((width * 2, bottom_crop.size[1] * 2), resample_nearest)
                    
                    time_regex = r'([01]?\d|2[0-3])\s*[:：]\s*([0-5]\d)\s*[:：]\s*([0-5]\d)(?:.*?(\d+)\s*ms)?'
                    
                    m = None
                    
                    # 1. ĐỌC THỜI GIAN TỪ THANH STATUS
                    # Bỏ hoàn toàn tiếng Nhật, chỉ dùng tiếng Anh (eng) để đọc số chính xác nhất
                    ocr_strategies = [
                        (bottom_img_nearest, 'eng', '--psm 7'),     # Ưu tiên 1: Cắt đáy, ảnh nét vuông (tránh nhầm 5 thành 6)
                        (bottom_img_lanczos, 'eng', '--psm 7'),     # Ưu tiên 2: Cắt đáy, ảnh mượt
                        (bottom_img_nearest, 'eng', '--psm 6'),
                        (bottom_img_lanczos, 'eng', '--psm 6'),
                        (full_img, 'eng', '--psm 11'),      # Dự phòng: Đọc rải rác toàn ảnh
                        (full_img, 'eng', '--psm 3'),       # Dự phòng: Đọc layout toàn ảnh
                    ]
                    
                    for img_variant, lang, psm in ocr_strategies:
                        text_variant = pytesseract.image_to_string(img_variant, lang=lang, config=psm)
                        m_variant = re.search(time_regex, text_variant, re.IGNORECASE | re.DOTALL)
                        if m_variant:
                            m = m_variant
                            break
                            
                    # 2. ĐỌC THÔNG SỐ JOBNET / INTERFACE ID TỪ BẢNG
                    # Dùng tiếng Anh và PSM 3 để đọc Layout, bỏ qua tiếng Nhật
                    full_text = pytesseract.image_to_string(full_img, lang='eng', config='--psm 3')

                    if m:
                        # Định dạng lại thành chuẩn HH:MM:SS
                        panel_start_time = f"{int(m.group(1)):02d}:{m.group(2)}:{m.group(3)}"
                        panel_duration_ms = m.group(4)
                        
                        # Trích xuất Jobnet ID và Interface ID thông minh bằng mảng từ vựng (Words)
                        # Giúp vượt qua các lỗi khoảng trắng / ký tự rác của OCR
                        words = re.findall(r'[A-Za-z0-9@_\-]+', full_text)
                        skip_words = {'string', 'type', 'name', 'id', 'interface', 'jobnet', 'value'}
                        
                        jobnet_id = None
                        interface_id = None
                        
                        for i, w in enumerate(words):
                            w_lower = w.lower()
                            if 'jobnet' in w_lower and not jobnet_id:
                                for next_w in words[i+1:]:
                                    nl = next_w.lower()
                                    if nl not in skip_words and len(next_w) > 3 and 'str' not in nl:
                                        jobnet_id = next_w
                                        break
                                        
                            if 'interface' in w_lower and not interface_id:
                                for next_w in words[i+1:]:
                                    nl = next_w.lower()
                                    if nl not in skip_words and len(next_w) > 3 and 'str' not in nl:
                                        interface_id = next_w
                                        break

                        if jobnet_id is not None:
                            print(f"  - [Ảnh Status Panel]: Tìm thấy Jobnet ID: '{jobnet_id}'")
                        else:
                            print(f"  - [Ảnh Status Panel]: ⚠️ Không tìm thấy Jobnet ID trong ảnh.")
                            
                        if interface_id is not None:
                            print(f"  - [Ảnh Status Panel]: Tìm thấy Interface ID: '{interface_id}'")
                            if interface_id != "":
                                if file_name and interface_id.lower() in file_name.lower():
                                    print(f"      ✅ OK: Interface ID '{interface_id}' có tồn tại trong tên file Excel.")
                                else:
                                    print(f"      ❌ LỖI: Interface ID '{interface_id}' KHÔNG tồn tại trong tên file Excel ({file_name}).")
                            else:
                                print(f"      ⚠️ Interface ID là chuỗi rỗng nên bỏ qua kiểm tra với tên file.")
                        else:
                            print(f"  - [Ảnh Status Panel]: ⚠️ Không tìm thấy Interface ID trong ảnh.")

                        break
            except Exception as e:
                print(f"      ⚠️ Lỗi khi xử lý ảnh (OCR): {e}")

    else:
        print("  - [Ảnh Status Panel]: ⚠️ Cần cài đặt thư viện OCR (pytesseract, Pillow) để đọc text từ ảnh.")
        
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
                    
                    # Tính độ trễ giây (Dung sai)
                    t_panel_dt = datetime.datetime.combine(datetime.date.today(), t_panel)
                    t_log_dt = datetime.datetime.combine(datetime.date.today(), t_log)
                    diff_seconds = (t_panel_dt - t_log_dt).total_seconds()
                    
                    if t_log >= t_panel:
                        print(f"      ✅ OK: Thời gian log ({first_log_time_str}) >= Thời gian ảnh ({panel_start_time})")
                    elif diff_seconds <= 2:
                        print(f"      ✅ OK (Chấp nhận dung sai {int(diff_seconds)}s): Thời gian log ({first_log_time_str}) ~ Thời gian ảnh ({panel_start_time})")
                    else:
                        print(f"      ❌ LỖI: Thời gian log ({first_log_time_str}) < Thời gian ảnh ({panel_start_time}) (Chênh lệch {int(diff_seconds)}s)")
                except ValueError:
                    print(f"      ⚠️ Không thể so sánh thời gian: Log ({first_log_time_str}), Ảnh ({panel_start_time})")
    else:
        print("  - [Ảnh Status Panel]: ⚠️ Không tìm thấy hoặc không đọc được thời gian '実行開始:' từ ảnh.")

    log_duration_ms = None
    log_end_time_str = None
    if data:
        # Quét ngược từ dưới lên để tìm chính xác dòng log kết thúc subflow
        for item in reversed(data):
            msg = str(item.get('message', ''))
            if "サブフローの実行が終了しました" in msg:
                m_log_dur = re.search(r'\[\s*(\d+)(?:\s*ms)?\s*\]', msg, re.IGNORECASE)
                if m_log_dur:
                    log_duration_ms = m_log_dur.group(1)
                
                val = item['nichiji']
                if val:
                    if isinstance(val, datetime.datetime):
                        log_end_time_str = val.strftime("%H:%M:%S")
                    else:
                        time_m = re.search(r'(\d{1,2}:\d{2}:\d{2})', str(val))
                        if time_m:
                            log_end_time_str = time_m.group(1)
                    break

    if panel_duration_ms and data:
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

    if end_time_log_img:
        print(f"  - [Ảnh Log Viewer]: Tìm thấy giờ kết thúc (サブフローの実行が終了しました) là {end_time_log_img}")
        if log_end_time_str:
            try:
                fmt = "%H:%M:%S"
                t_img = datetime.datetime.strptime(end_time_log_img, fmt).time()
                t_txt = datetime.datetime.strptime(log_end_time_str, fmt).time()
                
                if t_img == t_txt:
                    print(f"      ✅ OK: Thời gian kết thúc trong ảnh ({end_time_log_img}) KHỚP với log text ({log_end_time_str}).")
                else:
                    print(f"      ❌ LỖI: Thời gian kết thúc trong ảnh ({end_time_log_img}) LỆCH với log text ({log_end_time_str}).")
            except ValueError:
                print(f"      ⚠️ Không thể so sánh thời gian kết thúc: Ảnh ({end_time_log_img}), Log ({log_end_time_str})")
        else:
            print(f"      ⚠️ Không tìm thấy thời gian kết thúc trong log text để đối chiếu.")

    if jobnet_id is not None and data:
        jobnet_in_log = False
        log_jobnet_found_val = None
        
        # Chỉ lấy phần trước '@' của Jobnet ID trong ảnh để so sánh
        jobnet_id_compare = jobnet_id.split('@')[0] if jobnet_id else ""
        
        for item in data:
            msg = str(item.get('message', ''))
            if "ジョブネットID：" in msg:
                m_log_jobnet = re.search(r'ジョブネットID：([^\s　]*)', msg)
                if m_log_jobnet:
                    log_jobnet_found_val = m_log_jobnet.group(1)
                    log_jobnet_compare = log_jobnet_found_val.split('@')[0] if log_jobnet_found_val else ""
                    if log_jobnet_compare == jobnet_id_compare:
                        jobnet_in_log = True
                        break
            elif jobnet_id_compare and jobnet_id_compare in msg:
                jobnet_in_log = True
                break
                
        if jobnet_in_log:
            print(f"      ✅ OK: Jobnet ID '{jobnet_id}' khớp với dữ liệu trong Log.")
        else:
            if log_jobnet_found_val is not None:
                print(f"      ❌ LỖI: Jobnet ID trong ảnh là '{jobnet_id}', nhưng trong Log là '{log_jobnet_found_val}'.")
            else:
                print(f"      ❌ LỖI: Không tìm thấy dòng log chứa 'ジョブネットID：' hoặc Jobnet ID '{jobnet_id}' trong Log.")

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

    if target_run_num is not None and found_runs_set is not None:
        found_runs_set.add(target_run_num)

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