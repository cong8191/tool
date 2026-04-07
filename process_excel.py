import xlwings as xw
import pandas as pd
import os
import shutil
from datetime import datetime

def process_excel(input_ids=None, target_date=None):
    source_path = os.path.abspath('01.要件定義_インターフェース一覧（STEP3）.xlsx')
    template_path = os.path.abspath('template_FtoF.xlsx')
    design_docs_dir = r'D:\Project\151_ISA_AsteriaWrap\trunk\99_FromJP\10_プロジェクト資材\02.IFAgreement\03.確定'
    output_dir = os.getcwd()

    if not os.path.exists(source_path):
        print(f"Lỗi: Không tìm thấy file {source_path}")
        return

    processing_date = target_date if target_date else datetime.now().strftime('%Y/%m/%d')
    print(f"--- Đang đọc dữ liệu từ Master (Ngày: {processing_date}) ---")
    df = pd.read_excel(source_path, sheet_name='【STEP3】インターフェース一覧', skiprows=20, header=None)

    results = {}
    for index, row in df.iterrows():
        if_id = str(row[65]).strip() if pd.notnull(row[65]) else None
        
        if not if_id or if_id == '-': continue
        if input_ids and if_id not in input_ids: continue

        if if_id not in results: results[if_id] = []
        val_g = str(row[6]).zfill(4) if pd.notnull(row[6]) else ""
        val_h = str(row[7]) if pd.notnull(row[7]) else ""
        results[if_id].append({
            'cotc': row[2], 'cotd': row[3], 'cotj': row[9],
            'coty': row[23], 'cotz': row[24], 'cotaf': row[30], 'cotag': row[31],
            'cotah': row[32], 'cotai': row[33], 'cotaq': row[41], 'cotar': row[42],
            'cotax': row[48], 'cotay': row[49], 'cotaz': row[50],
            'cotba': row[51], 'combined_gh': f"{val_g}-{val_h}"
        })

    if not results: 
        print("Không có dữ liệu nào phù hợp với điều kiện đã cho.") 
        return

    print(f"--- Đang khởi động Excel ---")
    app = xw.App(visible=False)
    app.display_alerts = False
    try: app.api.AskToUpdateLinks = False
    except: pass
    
    try:
        existing_files = os.listdir(design_docs_dir) if os.path.exists(design_docs_dir) else []
        for if_id, items in results.items():
            found = next((f for f in existing_files if if_id in f), None)
            if not found: continue

            gh = items[0]['combined_gh']
            full_id = f"{if_id}_{gh}"
            file_name = f"連携機能設計書（{full_id}）.xlsx"
            out_p = os.path.join(output_dir, file_name)
            src_p = os.path.join(design_docs_dir, found)

            print(f"Đang xử lý: {file_name}")
            shutil.copy(template_path, out_p)
            wb_out = app.books.open(out_p)
            wb_src = app.books.open(src_p, update_links=False, read_only=True)

            try:
                # 1. Update Basic Info
                s_cover = next((s for s in wb_out.sheets if '表紙' in s.name), None)
                if s_cover:
                    s_cover.range('A15').value = f"（{full_id}）"
                    s_cover.range('A26').value = processing_date
                s_history = next((s for s in wb_out.sheets if '改版履歴' in s.name), None)
                if s_history:
                    s_history.range('A7').value = processing_date
                    # Thiết lập in trong 1 trang
                    s_history.api.PageSetup.Zoom = False
                    s_history.api.PageSetup.FitToPagesWide = 1
                    s_history.api.PageSetup.FitToPagesTall = 1

                # Shapes
                t_shapes = {"改版履歴": "正方形/長方形 19", "個別レイアウト情報": "正方形/長方形 37", "個別処理詳細": "正方形/長方形 38"}
                for sn, shn in t_shapes.items():
                    s = next((sh for sh in wb_out.sheets if sn in sh.name), None)
                    if s:
                        for shape in s.shapes:
                            if shape.name == shn:
                                try: shape.api.TextFrame2.TextRange.Text = processing_date
                                except: pass
                            if shape.type == 'group':
                                try:
                                    for sub in shape.api.GroupItems:
                                        if sub.Name == shn: sub.TextFrame2.TextRange.Text = processing_date
                                except: pass

                # 2. Layout Mapping
                s_layout = next((s for s in wb_out.sheets if '個別レイアウト情報' in s.name), None)
                if s_layout:
                    s_layout.range('O10').value = found
                    s_layout.range('O11').value = gh
                    s_src_m = next((s for s in wb_src.sheets if "マッピング定義" in s.name), None)
                    
                    def transform_val(v):
                        v_str = str(v).strip()
                        if v_str == "対象外": return "拡張子なし"
                        return v_str

                    it0 = items[0]
                    if s_src_m: s_layout.range('O22').value = s_src_m.range('B2').value
                    s_layout.range('O24').value = transform_val(it0['coty'])
                    s_layout.range('O25').value = transform_val(it0['cotz'])
                    s_layout.range('O26').value = it0['cotah']
                    s_layout.range('O27').value = it0['cotai']
                    s_layout.range('O28').value = it0['cotaf']
                    s_layout.range('O29').value = it0['cotag']
                    if s_src_m: s_layout.range('O33').value = s_src_m.range('U2').value
                    s_layout.range('O35').value = transform_val(it0['cotaq'])
                    s_layout.range('O36').value = transform_val(it0['cotar'])
                    s_layout.range('O37').value = it0['cotaz']
                    s_layout.range('O38').value = it0['cotba']
                    s_layout.range('O39').value = it0['cotax']
                    s_layout.range('O40').value = it0['cotay']
                    
                    s_layout.range('B15').value = [[i['cotc']] for i in items]
                    s_layout.range('J15').value = [[i['cotd']] for i in items]
                    s_layout.range('AE15').value = [[i['cotj']] for i in items]
                    s_layout.range('AP15').value = [["ApplicationLogMT"]] * len(items)

                # 3. Copy Sheets (Surgical approach to keep Comments & Zoom but kill Links)
                for name in ["IFA_機能概要", "IFA_マッピング定義"]:
                    s_old = next((s for s in wb_out.sheets if name == s.name), None)
                    if s_old: s_old.delete()

                for sn, dn in [("機能概要", "IFA_機能概要"), ("マッピング定義", "IFA_マッピング定義")]:
                    ss = next((s for s in wb_src.sheets if sn in s.name), None)
                    if ss:
                        ss.api.Copy(After=wb_out.sheets[-1].api)
                        new_s = wb_out.sheets[-1]
                        new_s.name = dn
                        
                        # A. Tiêu diệt link trong Cell
                        try: new_s.used_range.value = new_s.used_range.value
                        except: pass
                        
                        # B. Tiêu diệt link trong Data Validation (nguyên nhân chính gây Security Warning)
                        try: new_s.api.Cells.Validation.Delete()
                        except: pass

                # 4. Final Link/Security Clean
                try: 
                    wb_out.api.UpdateLinks = 3
                    wb_out.api.UpdateRemoteReferences = False
                except: pass
                
                # Break ALL external links
                for link_type in [1, 5]: # 1=Excel, 5=OLE
                    links = wb_out.api.LinkSources(link_type)
                    if links:
                        for l in links:
                            try: wb_out.api.BreakLink(l, link_type)
                            except: pass

                for c in wb_out.api.Connections:
                    try: c.Delete()
                    except: pass

                # Xóa Names lỗi/ngoài triệt để
                for n in wb_out.api.Names:
                    try:
                        if "#REF!" in n.Value or "[" in n.Value or "!" in n.Value:
                            n.Delete()
                    except: pass

                # 5. Reset View
                for s in wb_out.sheets:
                    try:
                        s.activate()
                        s.range('A1').select()
                    except: pass
                
                wb_out.sheets[0].activate()
                wb_out.save()
            finally:
                wb_out.close()
                wb_src.close()
    finally:
        app.quit()

if __name__ == "__main__":
    import sys
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('ids', nargs='*')
    parser.add_argument('--date', default=None)
    args = parser.parse_args()
    process_excel(args.ids if args.ids else None, target_date=args.date)
