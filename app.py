# -*- coding: utf-8 -*-
import os
import time
import subprocess
from flask import Flask, request, render_template_string, send_from_directory
import json

app = Flask(__name__)

# Thư mục lưu trữ các file tải lên và kết quả
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "web_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

HTML_FORM = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Công cụ So sánh Layout</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f7f6; margin: 40px; }
        .container { max-width: 500px; background: white; margin: auto; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        h2 { text-align: center; color: #333; }
        .form-group { margin-bottom: 15px; }
        label { font-weight: bold; display: block; margin-bottom: 5px; color: #555; }
        input[type="file"], input[type="text"] { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
        button { width: 100%; padding: 12px; background-color: #28a745; color: white; border: none; border-radius: 4px; font-size: 16px; cursor: pointer; margin-top: 10px; }
        button:hover { background-color: #218838; }
    </style>
</head>
<body>
    <div class="container">
        <h2>So Sánh Layout Data</h2>
        <form action="/compare" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label>1. File Excel Testcase (VD: check.xlsx):</label>
                <input type="file" name="excel_file" accept=".xlsx" required>
            </div>
            <div class="form-group">
                <label>2. File Output (Mainframe):</label>
                <input type="file" name="output_file" required>
            </div>
            <div class="form-group">
                <label>Bảng mã gốc (VD: cp20290, cp930):</label>
                <input type="text" name="from_enc" value="cp20290">
            </div>
            <div class="form-group">
                <label style="font-weight: normal; cursor: pointer;">
                    <input type="checkbox" name="keep_sosi" value="yes" checked>
                    Bảo toàn độ dài (Thay SOSI bằng khoảng trắng)
                </label>
            </div>
            <button type="submit">Bắt đầu so sánh</button>
        </form>
    </div>
</body>
</html>
"""

HTML_RESULT = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Kết quả so sánh</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #f4f7f6; margin: 40px; text-align: center; }
        .container { max-width: 95%; background: white; margin: auto; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .btn { display: block; width: 80%; margin: 15px auto; padding: 12px; text-decoration: none; color: white; border-radius: 4px; font-weight: bold; }
        .btn-excel { background-color: #1d6f42; } /* Excel green */
        .btn-txt { background-color: #007bff; } /* Text blue */
        .btn-back { background-color: #6c757d; width: auto; display: inline-block; padding: 8px 15px; margin-top: 20px;}
        .log-box { text-align: left; background: #eee; padding: 10px; border-radius: 4px; font-size: 12px; overflow-x: auto; max-height: 200px; }
        .preview-box { text-align: left; background: #272822; color: #f8f8f2; padding: 15px; border-radius: 6px; font-family: Consolas, "Courier New", monospace; font-size: 13px; overflow-x: auto; white-space: pre; max-height: 400px; border: 1px solid #444; }
        .summary-box { display: flex; justify-content: center; gap: 20px; margin: 25px 0; }
        .summary-item { padding: 15px 25px; border-radius: 8px; color: white; font-weight: bold; font-size: 18px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .summary-match { background-color: #28a745; }
        .summary-diff { background-color: #dc3545; }
        .table-container { overflow-x: auto; overflow-y: auto; margin-top: 20px; border: 1px solid #ddd; max-width: 100%; max-height: 600px; }
        .result-table { min-width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed; }
        .result-table th, .result-table td { border: 1px solid #ddd; padding: 8px; vertical-align: middle; position: relative; width: 120px; max-width: 120px; height: 40px; box-sizing: border-box; }
        .cell-content { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 100%; display: block; }
        .result-table thead { background-color: #f2f2f2; position: sticky; top: 0; z-index: 2;}
        .result-table th { font-weight: bold; text-align: center; }
        .result-table th .length { font-weight: normal; color: #666; font-size: 11px; }
        .result-table .row-header { font-weight: bold; background-color: #f2f2f2; text-align: center; position: sticky; left: 0; z-index: 3; width: 80px; min-width: 80px; max-width: 80px; }
        .result-table thead .row-header { z-index: 4; }
        .result-table td.status-match { background-color: #e6ffed; }
        .result-table td.status-diff { background-color: #ffebee; }
        .result-table td.has-tooltip { cursor: pointer; }
        .actual-tooltip {
            display: none;
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            background-color: #2b2b2b;
            color: #fff;
            padding: 10px 14px;
            border-radius: 6px;
            z-index: 10;
            box-shadow: 0 4px 12px rgba(0,0,0,0.4);
            text-align: left;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
            white-space: nowrap;
        }
        .result-table td.has-tooltip:hover .actual-tooltip {
            display: block;
        }
        .actual-tooltip::after {
            content: "";
            position: absolute;
            top: 100%;
            left: 50%;
            margin-left: -5px;
            border-width: 5px;
            border-style: solid;
            border-color: #2b2b2b transparent transparent transparent;
        }
        .tooltip-row { display: flex; align-items: center; gap: 8px; }
        .tooltip-expected { color: #4ade80; margin-bottom: 8px; padding-bottom: 8px; border-bottom: 1px solid #444; }
        .tooltip-actual { color: #f87171; }
        .tooltip-actual.is-match { color: #4ade80; }
        .tooltip-label { font-size: 12px; color: #aaa; font-family: Arial, sans-serif; width: 115px; flex-shrink: 0; }
        .tooltip-value { background-color: #1a1a1a; padding: 3px 6px; border-radius: 4px; border: 1px solid #444; color: #fff; min-width: 20px; display: inline-block; }
    </style>
</head>
<body>
    <div class="container">
        <h2 style="color: #28a745;">Hoàn tất xử lý!</h2>
        
        {% if summary %}
        <div class="summary-box">
            <div class="summary-item summary-match">✅ Khớp: {{ summary.match }}</div>
            <div class="summary-item summary-diff">❌ Lệch: {{ summary.diff }}</div>
        </div>
        {% endif %}

        {% if detailed_results %}
        <hr style="margin-top:30px; border-top: 1px solid #ccc;">
        <h3 style="text-align: left;">🔎 Chi tiết so sánh:</h3>
        <div class="table-container">
            <table class="result-table">
                <thead>
                    <tr>
                        <th class="row-header">Dòng Excel</th>
                        {% for field in detailed_results.layout %}
                        <th>
                            <div class="cell-content" title="{{ field.name }}">{{ field.name }}</div>
                            <div class="length">({{ field.length }}b)</div>
                        </th>
                        {% endfor %}
                    </tr>
                </thead>
                <tbody>
                    {% for row in detailed_results.results %}
                    <tr>
                        <td class="row-header">{{ row.row_index }}</td>
                        {% for field in row.fields %}
                        <td class="status-{{ field.status }} {% if field.status in ['diff'] %}has-tooltip{% endif %}">
                            <span class="cell-content">{{ field.expected if field.expected else '""' }}</span>
                            {% if field.status in ['diff'] %}
                                <div class="actual-tooltip">
                                    <div class="tooltip-expected tooltip-row">
                                        <span class="tooltip-label">Dự kiến (Excel):</span>
                                        <span class="tooltip-value">{{ field.expected.replace(' ', '␣') if field.expected else '""' }}</span>
                                    </div>
                                    <div class="tooltip-actual tooltip-row {% if field.status == 'match' %}is-match{% endif %}">
                                        <span class="tooltip-label">Thực tế (Output):</span>
                                        <span class="tooltip-value">{{ field.actual.replace(' ', '␣') if field.actual else '""' }}</span>
                                    </div>
                                </div>
                            {% endif %}
                        </td>
                        {% endfor %}
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        {% endif %}

        <p>Vui lòng tải các file kết quả bên dưới:</p>
        
        <a href="/download/{{ req_id }}/Result.xlsx" class="btn btn-excel">📥 Tải file Result.xlsx</a>
        
        {% if has_converted %}
        <a href="/download/{{ req_id }}/output_converted.txt" class="btn btn-txt">📥 Tải file Output đã Convert (Shift-JIS)</a>
        {% endif %}
        
        <a href="/" class="btn btn-back">⬅ Quay lại trang chủ</a>
        
        {% if preview_text %}
        <hr style="margin-top:30px; border-top: 1px solid #ccc;">
        <h3 style="text-align: left;">👀 Preview Dữ liệu Output (5000 byte đầu):</h3>
        <pre class="preview-box">{{ preview_text }}</pre>
        {% endif %}

        <hr style="margin-top:30px; border-top: 1px solid #ccc;">
        <h3 style="text-align: left;">Log xử lý:</h3>
        <pre class="log-box">{{ log }}</pre>
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_FORM)

@app.route('/compare', methods=['POST'])
def compare():
    excel_file = request.files.get('excel_file')
    output_file = request.files.get('output_file')
    from_enc = request.form.get('from_enc', '').strip()
    keep_sosi = request.form.get('keep_sosi')

    if not excel_file or not output_file or excel_file.filename == '' or output_file.filename == '':
        return "Vui lòng chọn đủ 2 file!", 400

    # Tạo một thư mục riêng biệt cho mỗi lần chạy (dựa trên Timestamp)
    req_id = str(int(time.time() * 1000))
    req_dir = os.path.join(UPLOAD_DIR, req_id)
    os.makedirs(req_dir, exist_ok=True)

    excel_path = os.path.join(req_dir, "check.xlsx")
    out_txt_path = os.path.join(req_dir, "output.txt")
    result_excel_path = os.path.join(req_dir, "Result.xlsx")
    
    excel_file.save(excel_path)
    output_file.save(out_txt_path)

    # Gọi script python hiện tại của bạn
    cmd = ["python", "compare_fixed_length.py", "--excel", excel_path, "--output", out_txt_path, "--out_excel", result_excel_path]
    if from_enc:
        cmd.extend(["--from_enc", from_enc])
    if keep_sosi:
        cmd.append("--keep_sosi")

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    process = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR, encoding='utf-8', errors='replace', env=env)

    has_converted = os.path.exists(os.path.join(req_dir, "output_converted.txt"))
    log_output = process.stdout + "\n" + process.stderr
    
    # Đọc nội dung file để preview (ưu tiên file đã convert)
    preview_text = ""
    preview_file_path = os.path.join(req_dir, "output_converted.txt") if has_converted else out_txt_path
    try:
        with open(preview_file_path, 'rb') as f:
            preview_bytes = f.read(5000)
            preview_text = preview_bytes.decode('shift_jis', errors='replace')
    except Exception as e:
        preview_text = f"Không thể tải trước preview: {e}"

    # Parse summary từ log
    summary = None
    for line in log_output.splitlines():
        if line.startswith("[SUMMARY]"):
            try:
                parts = line.replace("[SUMMARY]", "").strip().split(',')
                match_count = int(parts[0].split(':')[1].strip())
                diff_count = int(parts[1].split(':')[1].strip())
                summary = {"match": match_count, "diff": diff_count}
                break # Tìm thấy là dừng
            except (IndexError, ValueError):
                pass # Bỏ qua nếu dòng summary bị lỗi
    
    # Parse detailed results từ log
    detailed_results = None
    for line in log_output.splitlines():
        if line.startswith("[RESULT_JSON]"):
            try:
                json_str = line.replace("[RESULT_JSON]", "", 1)
                detailed_results = json.loads(json_str)
            except json.JSONDecodeError:
                pass # Bỏ qua nếu JSON bị lỗi
            break

    return render_template_string(
        HTML_RESULT, 
        req_id=req_id, 
        has_converted=has_converted,
        log=log_output,
        preview_text=preview_text,
        summary=summary,
        detailed_results=detailed_results
    )

@app.route('/download/<req_id>/<filename>')
def download(req_id, filename):
    # Chỉ cho phép tải 2 file chỉ định, bảo mật an toàn chống path traversal
    if filename not in ["Result.xlsx", "output_converted.txt"]:
        return "File không hợp lệ!", 400
        
    directory = os.path.join(UPLOAD_DIR, req_id)
    if not os.path.exists(os.path.join(directory, filename)):
        return "Không tìm thấy file kết quả. Có thể đã xảy ra lỗi trong lúc chạy.", 404
        
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == '__main__':
    # Mở mạng LAN (0.0.0.0) để người khác cùng IP truy cập được
    print("Đang khởi động Web Server... Truy cập vào http://127.0.0.1:5000 trên trình duyệt.")
    app.run(host='0.0.0.0', port=5000, debug=True)