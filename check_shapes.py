import xlwings as xw
import os

def list_shapes():
    template_path = os.path.abspath('template_FtoF.xlsx')
    if not os.path.exists(template_path):
        print("Không tìm thấy template.")
        return

    app = xw.App(visible=False)
    try:
        wb = app.books.open(template_path)
        for sheet in wb.sheets:
            print(f"Sheet: {sheet.name}")
            for shape in sheet.shapes:
                print(f"  - Shape Name: '{shape.name}', Type: {shape.type}")
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    list_shapes()
