# -*- coding: utf-8 -*-
import argparse
import sys
import os
import shutil
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

def try_fallback_tools(input_path, output_path, from_enc, to_enc, keep_sosi=False):
    """Tự động tìm và gọi công cụ iconv hoặc java có sẵn trên Windows"""
    from_enc_lower = from_enc.lower()
    to_enc_lower = to_enc.lower()

    if from_enc_lower in ['cp930', 'ibm930']: iconv_from, java_from, mixed_enc = 'IBM930', 'Cp930', 'Cp930'
    elif from_enc_lower in ['cp939', 'ibm939']: iconv_from, java_from, mixed_enc = 'IBM939', 'Cp939', 'Cp939'
    elif from_enc_lower in ['cp20290', 'ibm20290', 'cp290', 'ibm290']: iconv_from, java_from, mixed_enc = 'IBM290', 'Cp290', 'Cp930'
    else: iconv_from, java_from, mixed_enc = from_enc, from_enc, from_enc

    iconv_to = 'SHIFT-JIS' if to_enc_lower in ['shift_jis', 'sjis', 'cp932'] else to_enc

    java_to = 'MS932' if to_enc_lower in ['shift_jis', 'sjis', 'cp932'] else ('UTF-8' if to_enc_lower == 'utf-8' else to_enc)

    # 1. Ưu tiên thử dùng Java trước (do xử lý EBCDIC tốt hơn)
    if shutil.which('java') and shutil.which('javac'):
        print("\n--- Đang tự động dùng Java để thực hiện chuyển đổi thay cho Python... ---")
        if keep_sosi:
            print("--- Kích hoạt chế độ thay thế SOSI bằng space (0x20) ---")
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
                        out.write(0x20); start = i + 1; isDbcs = true;
                    }} else if (input[i] == 0x0F) {{
                        if (i > start) {{
                            byte[] chunk = Arrays.copyOfRange(input, start, i);
                            byte[] wrapped = new byte[chunk.length + 2];
                            wrapped[0] = 0x0E;
                            System.arraycopy(chunk, 0, wrapped, 1, chunk.length);
                            wrapped[wrapped.length - 1] = 0x0F;
                            out.write(new String(wrapped, mixedEnc).getBytes("{java_to}"));
                        }}
                        out.write(0x20); start = i + 1; isDbcs = false;
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
            print("\nChuyển đổi thành công bằng công cụ Java ngầm!")
            print(f"File kết quả đã được lưu tại: {output_path}")
            return True
        except subprocess.CalledProcessError as e: print(f"Lỗi khi chạy Java: {e}", file=sys.stderr)
        finally:
            if os.path.exists("TmpConverter.java"): os.remove("TmpConverter.java")
            if os.path.exists("TmpConverter.class"): os.remove("TmpConverter.class")

    # 2. Thử dùng iconv nếu Java thất bại (hoặc không cài)
    iconv_path = shutil.which('iconv')
    if not iconv_path and os.path.exists(r"C:\Program Files\Git\usr\bin\iconv.exe"):
        iconv_path = r"C:\Program Files\Git\usr\bin\iconv.exe"

    if iconv_path and not keep_sosi:
        print(f"\n--- Phát hiện công cụ phụ trợ iconv tại: {iconv_path} ---")
        print("--- Đang tự động gọi iconv để thực hiện chuyển đổi thay cho Python... ---")
        # Dùng -c để bỏ qua các ký tự rác không hợp lệ (nếu có ở cuối file mainframe)
        cmd = [iconv_path, '-c', '-f', iconv_from, '-t', iconv_to, '-o', output_path, input_path]
        try:
            subprocess.run(cmd, check=True)
            print("\nChuyển đổi thành công bằng công cụ iconv ngầm!")
            print(f"File kết quả đã được lưu tại: {output_path}")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Lỗi khi chạy iconv: {e}", file=sys.stderr)

    return False

def convert_file_encoding(input_path, output_path, from_encoding, to_encoding, keep_sosi=False):
    """
    Chuyển đổi encoding của một file.
    Đọc file input dưới dạng binary, decode bằng from_encoding,
    encode lại bằng to_encoding, và ghi ra file output.
    """
    try:
        print(f"--- Đang đọc file: {input_path} ---")
        with open(input_path, 'rb') as f_in:
            binary_data = f_in.read()

        print(f"--- Đang giải mã từ '{from_encoding}' ---")
        # Giải mã từ encoding nguồn (ví dụ: EBCDIC tiếng Nhật) sang chuỗi Python (Unicode)
        # errors='replace' sẽ thay thế các ký tự không hợp lệ bằng '?'
        decoded_string = binary_data.decode(from_encoding, errors='replace')

        print(f"--- Đang mã hoá sang '{to_encoding}' ---")
        # Mã hoá chuỗi Python sang encoding đích (ví dụ: Shift-JIS)
        output_data = decoded_string.encode(to_encoding, errors='replace')

        print(f"--- Đang ghi ra file: {output_path} ---")
        with open(output_path, 'wb') as f_out:
            f_out.write(output_data)

        print("\nChuyển đổi thành công!")
        print(f"File kết quả đã được lưu tại: {output_path}")

    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file đầu vào '{input_path}'", file=sys.stderr)
        sys.exit(1)
    except (UnicodeDecodeError, UnicodeEncodeError) as e:
        print(f"Lỗi encoding: {e}", file=sys.stderr)
        print("Vui lòng kiểm tra lại encoding nguồn và đích có chính xác không.", file=sys.stderr)
        sys.exit(1)
    except LookupError as e:
        # The exception message is usually "unknown encoding: <encoding_name>"
        # We can extract the name for a cleaner message.
        encoding_name = str(e).replace("unknown encoding: ", "").strip()
        print(f"Lỗi: Không nhận dạng được encoding '{encoding_name}'.", file=sys.stderr)
        print("Encoding này có thể không được tích hợp sẵn trong môi trường Python của bạn.", file=sys.stderr)

        # Tự động chuyển hướng gọi fallback tools thay vì thoát ngay
        if try_fallback_tools(input_path, output_path, from_encoding, to_encoding, keep_sosi):
            sys.exit(0)

        if any(enc in encoding_name for enc in ['cp930', 'cp939', 'cp20290', 'cp290', 'ibm']):
            print("\nLưu ý: Python thuần không hỗ trợ sẵn các bảng mã EBCDIC tiếng Nhật (như IBM930, IBM939, IBM20290, JEF).", file=sys.stderr)
            print("Giải pháp: Bạn nên sử dụng công cụ `iconv` của Git Bash hoặc sử dụng Java (hỗ trợ Cp930).", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Đã xảy ra lỗi không mong muốn: {e}", file=sys.stderr)
        sys.exit(1)

def main():
    """Hàm chính để phân tích tham số dòng lệnh và gọi hàm chuyển đổi."""
    parser = argparse.ArgumentParser(
        description="Công cụ chuyển đổi encoding file, ví dụ từ BCDIC/EBCDIC-JP sang Shift-JIS.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("input_file", help="Đường dẫn đến file nguồn cần chuyển đổi.")
    parser.add_argument("output_file", help="Đường dẫn đến file đích để lưu kết quả.")
    parser.add_argument(
        "--from-enc",
        default="cp930",
        help="""Encoding của file nguồn.
Mặc định là 'cp930'.
Người dùng có thể dùng 'BCDIC' để chỉ các encoding EBCDIC tiếng Nhật.
Các lựa chọn phổ biến cho EBCDIC-JP:
- cp930: EBCDIC Japanese Katakana-Kanji
- cp939: EBCDIC Japanese Latin-Kanji
- cp20290: EBCDIC Japanese Katakana Extended (SBCS/Mixed)
"""
    )
    parser.add_argument(
        "--to-enc",
        default="shift_jis",
        help="""Encoding của file đích.
Mặc định là 'shift_jis'.
Các lựa chọn phổ biến:
- shift_jis (hoặc sjis)
- cp932 (biến thể của Shift-JIS từ Microsoft, tương thích rộng hơn)
- utf-8
"""
    )
    parser.add_argument(
        "--keep-sosi",
        action="store_true",
        help="Thay thế ký tự điều khiển SOSI (0x0E, 0x0F) bằng ký tự space (0x20) để bảo toàn chiều dài."
    )

    args = parser.parse_args()

    convert_file_encoding(args.input_file, args.output_file, args.from_enc, args.to_enc, args.keep_sosi)

if __name__ == "__main__":
    main()