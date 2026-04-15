# -*- coding: utf-8 -*-
import argparse
import openpyxl
from openpyxl.styles import PatternFill
import os
import sys
import datetime
import shutil
import subprocess
import tempfile

def try_fallback_tools(input_path, output_path, from_enc, to_enc, keep_sosi=False):
    """Tự động tìm và gọi công cụ iconv hoặc java có sẵn trên Windows để convert"""
    from_enc_lower = from_enc.lower()
    to_enc_lower = to_enc.lower()

    if from_enc_lower in ['cp930', 'ibm930']: iconv_from, java_from, mixed_enc = 'IBM930', 'Cp930', 'Cp930'
    elif from_enc_lower in ['cp939', 'ibm939']: iconv_from, java_from, mixed_enc = 'IBM939', 'Cp939', 'Cp939'
    elif from_enc_lower in ['cp20290', 'ibm20290', 'cp290', 'ibm290']: iconv_from, java_from, mixed_enc = 'IBM290', 'Cp290', 'Cp930'
    else: iconv_from, java_from, mixed_enc = from_enc, from_enc, from_enc

    iconv_to = 'SHIFT-JIS' if to_enc_lower in ['shift_jis', 'sjis', 'cp932'] else to_enc
    java_to = 'MS932' if to_enc_lower in ['shift_jis', 'sjis', 'cp932'] else ('UTF-8' if to_enc_lower == 'utf-8' else to_enc)

    # 1. Ưu tiên dùng Java trước vì Java xử lý các bảng mã EBCDIC IBM chuẩn xác hơn iconv trên Windows
    if shutil.which('java') and shutil.which('javac'):
        if keep_sosi:
            print(" -> [DEBUG] Kích hoạt chế độ giữ lại Shift-Out(0x0E) và Shift-In(0x0F)")
            java_code = f"""import java.io.ByteArrayOutputStream; import java.nio.file.Files; import java.nio.file.Paths; import java.util.Arrays; import java.nio.charset.Charset;
            public class TmpConverter {{ public static void main(String[] args) throws Exception {{
                String mixedEnc = "{mixed_enc}";
                if (!Charset.isSupported(mixedEnc)) {{
                    String[] fallbacks = {{"Cp930", "IBM930", "x-IBM930", "Cp939", "IBM939", "x-IBM939"}};
                    for (String enc : fallbacks) {{
                        if (Charset.isSupported(enc)) {{ mixedEnc = enc; break; }}
                    }}
                }}
                byte[] input = Files.readAllBytes(Paths.get(args[0]));
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                int start = 0; boolean isDbcs = false;
                for (int i = 0; i < input.length; i++) {{
                    if (input[i] == 0x0E) {{
                        if (i > start) {{
                            byte[] chunk = Arrays.copyOfRange(input, start, i);
                            out.write(new String(chunk, mixedEnc).getBytes("{java_to}"));
                        }}
                        out.write(0x0E); start = i + 1; isDbcs = true;
                    }} else if (input[i] == 0x0F) {{
                        if (i > start) {{
                            byte[] chunk = Arrays.copyOfRange(input, start, i);
                            byte[] wrapped = new byte[chunk.length + 2];
                            wrapped[0] = 0x0E;
                            System.arraycopy(chunk, 0, wrapped, 1, chunk.length);
                            wrapped[wrapped.length - 1] = 0x0F;
                            out.write(new String(wrapped, mixedEnc).getBytes("{java_to}"));
                        }}
                        out.write(0x0F); start = i + 1; isDbcs = false;
                    }}
                }}
                if (start < input.length) {{
                    byte[] chunk = Arrays.copyOfRange(input, start, input.length);
                    if (isDbcs) {{
                        byte[] wrapped = new byte[chunk.length + 2];
                        wrapped[0] = 0x0E;
                        System.arraycopy(chunk, 0, wrapped, 1, chunk.length);
                        wrapped[wrapped.length - 1] = 0x0F;
                        out.write(new String(wrapped, mixedEnc).getBytes("{java_to}"));
                    }} else {{
                        out.write(new String(chunk, mixedEnc).getBytes("{java_to}"));
                    }}
                }}
                Files.write(Paths.get(args[1]), out.toByteArray());
            }} }}"""
        else:
            java_code = f"""import java.nio.file.Files; import java.nio.file.Paths; import java.nio.charset.Charset;
            public class TmpConverter {{ public static void main(String[] args) throws Exception {{
                String fromEnc = "{java_from}";
                if (!Charset.isSupported(fromEnc)) {{
                    String[] fallbacks = {{"x-IBM290", "Cp290", "IBM290", "Cp930", "IBM930", "x-IBM930"}};
                    for (String enc : fallbacks) {{
                        if (Charset.isSupported(enc)) {{
                            fromEnc = enc;
                            break;
                        }}
                    }}
                }}
                Files.write(Paths.get(args[1]), new String(Files.readAllBytes(Paths.get(args[0])), fromEnc).getBytes("{java_to}"));
            }} }}"""
        with open("TmpConverter.java", "w", encoding="utf-8") as f: f.write(java_code)
        try:
            subprocess.run(['javac', 'TmpConverter.java'], check=True)
            subprocess.run(['java', 'TmpConverter', input_path, output_path], check=True)
            print(" -> [DEBUG] Đã dùng Java để convert file thành công.")
            return True
        except subprocess.CalledProcessError as e: print(f"Lỗi khi chạy Java: {e}", file=sys.stderr)
        finally:
            if os.path.exists("TmpConverter.java"): os.remove("TmpConverter.java")
            if os.path.exists("TmpConverter.class"): os.remove("TmpConverter.class")

    # 2. Thử dùng iconv nếu Java lỗi hoặc không có Java
    iconv_path = shutil.which('iconv')
    if not iconv_path and os.path.exists(r"C:\Program Files\Git\usr\bin\iconv.exe"):
        iconv_path = r"C:\Program Files\Git\usr\bin\iconv.exe"

    if iconv_path and not keep_sosi:
        cmd = [iconv_path, '-c', '-f', iconv_from, '-t', iconv_to, '-o', output_path, input_path]
        try:
            subprocess.run(cmd, check=True)
            print(" -> [DEBUG] Đã dùng iconv để convert file thành công.")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Lỗi khi chạy iconv: {e}", file=sys.stderr)

    return False

def convert_file_to_temp(input_path, output_path, from_encoding, to_encoding, keep_sosi=False):
    try:
        with open(input_path, 'rb') as f_in:
            binary_data = f_in.read()
        decoded_string = binary_data.decode(from_encoding, errors='replace')
        output_data = decoded_string.encode(to_encoding, errors='replace')
        with open(output_path, 'wb') as f_out:
            f_out.write(output_data)
        print(" -> [DEBUG] Đã dùng Python thuần để convert file thành công.")
        return True
    except LookupError as e:
        return try_fallback_tools(input_path, output_path, from_encoding, to_encoding, keep_sosi)
    except Exception as e:
        print(f"Lỗi chuyển đổi: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Tool tìm kiếm data trong file output dựa trên tổng số byte và đánh dấu trực tiếp vào file Excel.")
    parser.add_argument("--excel", required=True, help="Đường dẫn file Excel (Dòng 1: Name, Dòng 2: Length, Dòng 3 trở đi: Data testcase).")
    parser.add_argument("--output", required=True, help="Đường dẫn file text Output.")
    parser.add_argument("--out_excel", default="Result.xlsx", help="Đường dẫn file Excel xuất ra báo cáo (Mặc định: Result.xlsx).")
    parser.add_argument("--encoding", default="shift_jis", help="Bảng mã file text (Mặc định: shift_jis).")
    parser.add_argument("--from_enc", default=None, help="Bảng mã gốc của file Output (VD: cp20290). Nếu có, tool sẽ tự động convert trước khi xử lý.")
    parser.add_argument("--keep_sosi", action="store_true", help="Giữ nguyên ký tự điều khiển Shift-Out (0x0E) và Shift-In (0x0F) để bảo toàn chiều dài byte.")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.excel):
        print(f"Lỗi: Không tìm thấy file excel '{args.excel}'")
        return
        
    if not os.path.exists(args.output):
        print(f"Lỗi: Không tìm thấy file output '{args.output}'")
        return
        
    print("1. Đang đọc file Excel...")
    wb = openpyxl.load_workbook(args.excel)
    sheet = wb.active

    # Đọc layout từ dòng 1 (name) và dòng 2 (length)
    layout_total_bytes = 0
    layout = []
    
    if sheet.max_row < 2:
        print("Lỗi: File Excel phải có ít nhất 2 dòng (dòng 1 cho Name, dòng 2 cho Length).")
        return

    field_names = [cell.value for cell in sheet[1]]
    field_lengths = [cell.value for cell in sheet[2]]

    for i in range(sheet.max_column):
        name = field_names[i]
        length = field_lengths[i]
        
        if length is not None and str(length).strip().isdigit():
            length = int(length)
            layout_total_bytes += length
            # Lưu lại vị trí cột (1-based) để dùng sau
            layout.append({'name': str(name).strip(), 'length': length, 'col': i + 1})

    print(f" -> Tìm thấy {len(layout)} fields. Tổng số byte layout (dự kiến): {layout_total_bytes}")
    if layout_total_bytes == 0:
        print("Lỗi: Tổng số byte layout = 0. Vui lòng kiểm tra lại dòng 2 trong file Excel có chứa số hay không.")
        return

    print("2. Đang đọc file Output...")
    output_file_to_read = args.output

    if args.from_enc:
        print(f" -> Tự động convert output từ '{args.from_enc}' sang '{args.encoding}'...")
        
        # Đổi thành lưu file cố định để user có thể mở lên xem nội dung đã convert
        base_name, _ = os.path.splitext(args.output)
        converted_file_path = f"{base_name}_converted.txt"
        
        if convert_file_to_temp(args.output, converted_file_path, args.from_enc, args.encoding, args.keep_sosi):
            output_file_to_read = converted_file_path
            print(f" -> [DEBUG] File convert thành công. Đã lưu bản sao tại: {converted_file_path}")
        else:
            print("Lỗi: Convert thất bại. Vui lòng kiểm tra lại môi trường hoặc bảng mã.")
            return

    with open(output_file_to_read, 'rb') as f:
        output_bytes = f.read()

    if args.from_enc:
        # Loại bỏ SOSI (0x0E, 0x0F) để preview không bị lỗi font khi decode Shift-JIS
        preview_str = output_bytes[:500].replace(b'\x0E', b'').replace(b'\x0F', b'').decode(args.encoding, errors='replace')
        print(f" -> [PREVIEW NỘI DUNG SAU CONVERT (500 byte đầu)]:\n{preview_str}\n" + "-"*50)

    # Nhận diện xem file output có ký tự xuống dòng hay không
    use_lines = False
    if b'\n' in output_bytes[:layout_total_bytes + 100]:
        print(" -> Nhận diện file có chia dòng (Newline). Sẽ trích xuất đoạn chunk theo từng dòng.")
        lines = output_bytes.splitlines()
        use_lines = True
    else:
        print(" -> File không có ký tự xuống dòng. Sẽ tự động tính length từng dòng data và cộng dồn offset.")

    print("3. Đang tìm kiếm data và đánh dấu vào Excel...")
    fill_match = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Xanh lá (Khớp)
    fill_diff = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Đỏ (Lệch)

    # Dòng data đầu tiên là dòng 3
    start_data_row = 3
    if sheet.max_row < start_data_row:
        print("Lỗi: Không có dòng data nào để kiểm tra (cần data từ dòng 3 trở đi).")
        return

    current_offset = 0

    # Duyệt từ dòng 3 đến dòng cuối cùng
    for row_idx in range(start_data_row, sheet.max_row + 1):
        record_index = row_idx - start_data_row
        
        row_total_bytes = 0
        field_data_list = []
        
        # Tính chiều dài thực tế cho dòng hiện tại dựa trên giá trị Testcase
        for field in layout:
            col_idx = field['col']
            cell = sheet.cell(row=row_idx, column=col_idx)
            expected_val = cell.value
            
            if expected_val is not None:
                # Format lại data (Không dùng .strip() để bảo toàn nguyên vẹn khoảng trắng bạn đã nhập)
                if isinstance(expected_val, float) and expected_val.is_integer():
                    expected_str = str(int(expected_val))
                elif isinstance(expected_val, datetime.datetime):
                    expected_str = expected_val.strftime("%Y%m%d")
                else:
                    expected_str = str(expected_val)
                
                # Encode và tính độ dài byte thực tế
                expected_bytes = expected_str.encode(args.encoding, errors='replace')
                field_len = len(expected_bytes)
                
                field_data_list.append({'cell': cell, 'bytes': expected_bytes, 'str': expected_str, 'has_data': True, 'name': field['name'], 'layout_len': field['length']})
            else:
                # Nếu ô thực sự không có data (None), độ dài sẽ tính là 0 (Không lấy từ dòng 2 nữa)
                field_len = 0
                field_data_list.append({'cell': cell, 'has_data': False, 'name': field['name'], 'layout_len': field['length']})
                
            row_total_bytes += field_len

        # Trích xuất đoạn byte chunk tương ứng với dòng data này
        if use_lines:
            if record_index < len(lines): 
                chunk_bytes = lines[record_index]
            else:
                chunk_bytes = b"" # Output không đủ dòng
        else:
            # Luôn cắt block theo chiều dài chuẩn của layout để đảm bảo không bị lệch Record
            slice_length = layout_total_bytes
            start_pos = current_offset
            end_pos = start_pos + slice_length
            chunk_bytes = output_bytes[start_pos:end_pos]
            current_offset = end_pos # Cộng dồn cho record tiếp theo
            
        # Loại bỏ ký tự SOSI (0x0E, 0x0F) ra khỏi chunk để việc decode và match chuỗi Shift-JIS không bị gãy đoạn
        chunk_clean = chunk_bytes.replace(b'\x0E', b'').replace(b'\x0F', b'')

        print(current_offset)
        print(f" - Đang kiểm tra dòng {row_idx} (Record {record_index + 1}) | Chiều dài tự tính: {row_total_bytes} bytes")
        
        # In ra cảnh báo chi tiết các field có độ dài nhập vào khác với layout
        if row_total_bytes != layout_total_bytes and row_total_bytes > 0:
            print(f"   -> [GỢI Ý] Chiều dài data bạn nhập ({row_total_bytes}) đang lệch so với Layout ({layout_total_bytes}). Nguyên nhân từ các ô Excel sau:")
            for f_data in field_data_list:
                if f_data['has_data']:
                    actual_len = len(f_data['bytes'])
                    # if actual_len != f_data['layout_len']:
                    #     print(f"      + Field '{f_data['name']}': Layout = {f_data['layout_len']} byte | Thực tế nhập = {actual_len} byte -> '{f_data['str']}'")

        if not chunk_bytes:
            print(f"   -> Cảnh báo: File output không có đủ data cho record {record_index + 1}")
        else:
            display_chunk = chunk_clean.decode(args.encoding, errors='replace')
            print(f"   [Data Output]: '{display_chunk}'")
        
        for f_data in field_data_list:
            if f_data['has_data']:
                cell = f_data['cell']
                expected_bytes = f_data['bytes']
                expected_str = f_data['str']

                if expected_bytes in chunk_clean:
                    cell.fill = fill_match
                else:
                    cell.fill = fill_diff
                    
                    # In ra log chuỗi thực tế trong file output (đã giải mã) để dễ debug
                    chunk_str_for_debug = chunk_bytes.decode(args.encoding, errors='replace')
                    # Chỉ in 100 ký tự để log không bị quá dài
                    snippet = chunk_str_for_debug[:100].replace('\r', '').replace('\n', '')
                    # print(f"      [DIFF] Excel: '{expected_str}' | Thực tế file Output có: '{snippet}'...")
    
    print(f"4. Đang lưu kết quả ra: {args.out_excel}")
    wb.save(args.out_excel)
    print("--- Hoàn tất ---")

if __name__ == "__main__":
    main()