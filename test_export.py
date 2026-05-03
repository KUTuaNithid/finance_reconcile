import openpyxl
import os

FS_CELLS = {
    "เงินสดและรายการเทียบเท่าเงินสด": {"sheet": "สินทรัพย์", "cell": "C9"},
    "ลูกหนี้การค้า": {"sheet": "สินทรัพย์", "cell": "C10"},
    "สินทรัพย์หมุนเวียนอื่น ": {"sheet": "สินทรัพย์", "cell": "C11"},
    "อุปกรณ์-สุทธิ": {"sheet": "สินทรัพย์", "cell": "C15"},
    "เจ้าหนี้อื่น": {"sheet": "สินทรัพย์", "cell": "C42"},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"sheet": "สินทรัพย์", "cell": "C43"},
    "หนี้สินหมุนเวียนอื่น ": {"sheet": "สินทรัพย์", "cell": "C44"},
    "รายได้จากการให้บริการ": {"sheet": "กำไรขาดทุน", "cell": "D8"},
    "รายได้อื่น": {"sheet": "กำไรขาดทุน", "cell": "D9"},
    "ต้นทุนการให้บริการ": {"sheet": "กำไรขาดทุน", "cell": "D13"},
    "ค่าใช้จ่ายในการขายและบริหาร": {"sheet": "กำไรขาดทุน", "cell": "D14"},
}

def export_fs(fs_summary_dict):
    template_path = "../FS.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    for fs_line, val in fs_summary_dict.items():
        if fs_line in FS_CELLS:
            sheet_name = FS_CELLS[fs_line]["sheet"]
            cell_ref = FS_CELLS[fs_line]["cell"]
            wb[sheet_name][cell_ref].value = val
            
    # Test save
    wb.save("test_out.xlsx")
    print("Saved to test_out.xlsx")

test_data = {"เงินสดและรายการเทียบเท่าเงินสด": 500000.50}
export_fs(test_data)

