import pandas as pd
import re

# ===== BƯỚC 0: ĐỌC & CHUẨN HÓA TÊN CỘT =====
df = pd.read_csv("Fact_GopY_NLP_Final.csv", sep=';', encoding='utf-8-sig')

def clean_col(c):
    c = str(c)
    c = c.replace('\xa0', ' ')      # non-breaking space -> space thường
    c = c.replace('\n', ' ').replace('\r', ' ')
    c = re.sub(r'\s+', ' ', c)
    return c.strip()

df.columns = [clean_col(c) for c in df.columns]

id_cols = ["ID", "Tên ĐVKD", "Loại hình ĐVKD", "Mã NV (HDxxxxxx)", "Chức danh", "Trang_Thai_Phat_Sinh"]

q_fields = ["Q1_MoTaLoi", "Q2_BuocPhatSinh", "Q3_ThongTinGD", "Q4_PhanLoai",
            "Q5_ThoiGian", "Q6_XuLyTam", "Q7_RuiRo", "Q8_TanSuat",
            "Q9_DoiTuong", "Q10_NguyenNhanGoc", "Q11_DeXuat",
            "DIM1_NghiepVu_ML", "DIM2_NguyenNhanLoi_ML", "DIM1_Method", "DIM1_Ref"]

# ===== HÀM CHECK_ERROR_SIGNAL (bản đã fix bug xóa dấu tiếng Việt) =====
def check_error_signal(q1_text):
    if pd.isna(q1_text): return 'Không phát sinh'
    t = re.sub(r'\s+', ' ', str(q1_text).lower().strip())

    veto_keywords = ['hệ thống báo lỗi', 'duyệt không thành công', 'vẫn bị trừ tiền',
                      'không cảnh báo', 'ghi có vào tài khoản 2 lần', 'ứng dụng đang tạm thời gặp lỗi']
    if any(v in t for v in veto_keywords): return 'Có phát sinh lỗi'

    strong_rejects = ['không phát sinh nhiều nên chưa gặp lỗi', 'chưa phát sinh lỗi', 'không phát sinh lỗi',
        'không vướng mắc', 'khong vuong mac', 'đvkd không phát sinh', 'đơn vị không phát sinh',
        'không phát sinh từ', 'chưa có phát sinh', 'không có phát sinh', 'không phát sinh giao dịch',
        'không có tình huống', 'không có nghiệp vụ', 'ko có nghiệp vụ', 'đã trả lời', 'vì cũng lâu rồi',
        'như trên', 'không nhớ', 'chưa thấy phát sinh', 'gần đây nhất thì không có',
        'chỉ phát sinh các nghiệp vụ cơ bản', 'không gặp lỗi', 'đơn vị chưa phát sinh',
        'đơn vị không ps', 'đơn vị k ps', 'không có ý kiến', 'không ý kiến',
        'hiện tại, những giao dịch đvkd thực hiện chưa', 'từ 1/2/26 đến nay đơn vị thực hiện giao dịch']
    if any(p in t for p in strong_rejects): return 'Không phát sinh'

    exact_matches = ['không', 'khong', 'k', 'ko', 'chưa', 'chua', 'không có', 'khong co', 'chưa có', 'chua co',
        'không ps', 'khong ps', 'không phát sinh', 'khong phat sinh', 'chưa phát sinh', 'chua phat sinh',
        'ok', 'bình thường', 'tốt', 'na', 'n/a', 'none', 'nan', 'x', '.', '-', '0']
    if t in exact_matches: return 'Không phát sinh'

    if len(t) < 100:
        temp = t
        for w in ['không','khong','có','co','chưa','chua','phát sinh','phat sinh','ps','lỗi','loi','ý kiến','y kien']:
            temp = temp.replace(w, '')
        temp = re.sub(r'[\s\.\,\-]', '', temp)   # FIX: chỉ xóa khoảng trắng/dấu câu, giữ chữ có dấu
        if len(temp) < 5: return 'Không phát sinh'

    return 'Có phát sinh lỗi'

def extract_th(df, prefix):
    cols_map = {f"{prefix}_{f}": f for f in q_fields}
    sub = df[id_cols + list(cols_map.keys())].rename(columns=cols_map).copy()
    sub["Tinh_Huong"] = prefix
    sub["Trang_Thai_TinhHuong"] = sub["Q1_MoTaLoi"].apply(check_error_signal)

    def has_signal(row):
        combined = " ".join([
            str(row.get("Q6_XuLyTam", "")),
            str(row.get("Q10_NguyenNhanGoc", "")),
            str(row.get("Q11_DeXuat", ""))
        ])
        return check_error_signal(combined) == "Có phát sinh lỗi"

    sub["Co_The_Bo_Sot"] = sub.apply(has_signal, axis=1)
    return sub

th1 = extract_th(df, "TH1")
th2 = extract_th(df, "TH2")
fact_gopy_long = pd.concat([th1, th2], ignore_index=True)

fact_gopy_long.to_csv("Fact_GopY_long_TEST.csv", index=False, encoding='utf-8-sig', sep=';')

print(f"Tổng dòng sau unpivot: {len(fact_gopy_long)}")
print(f"Số case có thể bị bỏ sót khỏi AI (Co_The_Bo_Sot=True): {fact_gopy_long['Co_The_Bo_Sot'].sum()}")
print("Đã xuất file: Fact_GopY_long_TEST.csv")