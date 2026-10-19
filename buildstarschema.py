"""
build_star_schema.py
Từ 1 file Excel gốc (survey-dicore-giaodichtaiquay.xlsx) -> star schema cho Power BI.
Chỉ cần chạy:  python build_star_schema.py

Logic xử lý lấy từ clean_data.py và clean_data_gopy.py:
  - Nhận diện 47 nghiệp vụ, mỗi nghiệp vụ lấy 3 cột theo thứ tự (mức hỗ trợ, thời gian xử lý, loại lỗi)
  - Gắn cờ ĐVKD_Flag (Yes/No), Thu_Tu_NV, Thu_Tu_SLA theo sla_order_map
  - Phân loại góp ý: Trang_Thai_TinhHuong (từ Q1), Co_The_Bo_Sot (từ Q6 + Q10 + Q11)
Sau đó tách thành: Fact_NghiepVu, Fact_GopY, Dim_DonVi, Dim_NghiepVu, Dim_MucDoHoTro, Dim_SLA, Dim_LoaiLoi.

Đầu ra nằm trong thư mục star_schema/.
"""
import re
import unicodedata
from pathlib import Path

import pandas as pd

# =====================================================================
# CONFIG
# =====================================================================
INPUT_FILE = "survey-dicore-giaodichtaiquay.xlsx"   # đổi tên cho đúng file của bạn
OUTPUT_DIR = Path("star_schema")

KEEP_MA_NV = True   # False nếu không muốn đưa Mã NV vào output. Email luôn bị loại.

SLA_ORDER_MAP = {'<5 phút': 1, '>5 - 15 phút': 2, '>15 - 30 phút': 3, '>30 - 60 phút': 4,
                 '>60 phút': 5, 'Chưa phát sinh': 6, 'Không có dữ liệu': 99}

# Nhóm nghiệp vụ theo 4 khối trong eForm (TÊN NHÓM DO MÌNH ĐẶT, sửa theo cách công ty gọi)
NHOM_NV = [(1, 10, "Khách hàng & tài khoản"), (11, 17, "Thanh toán & ngoại tệ"),
           (18, 27, "Quỹ & TSĐB"), (28, 47, "Tín dụng")]

# 11 câu hỏi trong mỗi tình huống TH1/TH2 (thứ tự như clean_data_gopy.py)
Q_NAMES = ["Q1_MoTaLoi", "Q2_BuocPhatSinh", "Q3_ThongTinGD", "Q4_PhanLoai", "Q5_ThoiGian",
           "Q6_XuLyTam", "Q7_RuiRo", "Q8_TanSuat", "Q9_DoiTuong", "Q10_NguyenNhanGoc", "Q11_DeXuat"]
# =====================================================================

ID_COLS = {"ID": "Response_ID", "Tên ĐVKD": "Ten_DVKD", "Loại hình ĐVKD": "Loai_Hinh_DVKD",
           "Mã NV (HDxxxxxx)": "Ma_NV", "Chức danh": "Chuc_Danh"}


def clean_col(c):
    c = str(c).replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", c).strip()


def plain(s):
    s = unicodedata.normalize("NFD", str(s).lower())
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").replace("đ", "d")
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------------
# Logic phân loại góp ý (giữ nguyên từ clean_data_gopy.py)
# ---------------------------------------------------------------------
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
        for w in ['không', 'khong', 'có', 'co', 'chưa', 'chua', 'phát sinh', 'phat sinh', 'ps', 'lỗi', 'loi',
                  'ý kiến', 'y kien']:
            temp = temp.replace(w, '')
        temp = re.sub(r'[\s\.\,\-]', '', temp)
        if len(temp) < 5: return 'Không phát sinh'

    return 'Có phát sinh lỗi'


# ---------------------------------------------------------------------
# Logic gắn cờ nghiệp vụ (giữ nguyên từ clean_data.py)
# ---------------------------------------------------------------------
def check_logic_mau_thuan(row):
    col1 = str(row['Muc_Do_Ho_Tro']).lower().strip()
    col2 = str(row['Thoi_Gian_Xu_Ly']).lower().strip()
    col3 = str(row['Loai_Loi']).lower().strip()
    if "chưa phát sinh nghiệp vụ" in col1:
        if ("chưa phát sinh" not in col2) or ("không phát sinh lỗi" not in col3):
            return "Yes"
    return "No"


# ---------------------------------------------------------------------
# Phân nhóm cho dim
# ---------------------------------------------------------------------
def nhom_nv(n):
    return next((g for lo, hi, g in NHOM_NV if lo <= n <= hi), "Khác")


def ho_tro_info(s):
    p = plain(s)
    if p.startswith("chua phat sinh"):
        return "Chưa phát sinh", 90
    if p.startswith("khong co du lieu"):
        return "Không có dữ liệu", 99
    if "khong can ho tro" in p:
        return "Tự thao tác", 1
    if "giai dap" in p or "huong dan" in p:
        return "CNTT hướng dẫn", 2
    if "vnc" in p or "xu ly truc tiep" in p:
        return "CNTT xử lý trực tiếp", 3
    return "Khác", 50


def loi_info(s):
    p = plain(s)
    if p.startswith("khong phat sinh loi"):
        return 0, 0
    if p.startswith("loi"):
        return 1, 1
    if p.startswith("chua phat sinh"):
        return 0, 90
    if p.startswith("khong co du lieu"):
        return 0, 99
    return 0, 50


def main():
    # =================== 1. ĐỌC FILE GỐC ===================
    print(f"[1/6] Đọc file gốc: {INPUT_FILE}")
    df_raw = pd.read_excel(INPUT_FILE)
    df_raw.columns = [clean_col(c) for c in df_raw.columns]
    print(f"      {df_raw.shape[0]} phản hồi x {df_raw.shape[1]} cột")

    missing = [c for c in ID_COLS if c not in df_raw.columns]
    if missing:
        raise KeyError(f"Không thấy cột {missing}. Các cột hiện có: {list(df_raw.columns[:15])}...")
    id_cols = list(ID_COLS)
    assert df_raw["ID"].is_unique, "Cột ID bị trùng"

    # =================== 2. UNPIVOT NGHIỆP VỤ ===================
    print("[2/6] Chuẩn hóa tên cột và unpivot nghiệp vụ")
    idx_chucdanh = df_raw.columns.get_loc('Chức danh')
    unique_nvs = []
    for col in df_raw.columns[idx_chucdanh + 1:]:
        col_str = str(col)
        if col_str.startswith('TH1.') or col_str.startswith('TH2.'):
            break
        base_name = re.sub(r'(\.\d+|\s*\(\d+\)|\s*\d+)$', '', col_str).strip()
        if base_name not in unique_nvs and re.match(r'^\d+\.', base_name):
            unique_nvs.append(base_name)
    print(f"      Nhận diện {len(unique_nvs)} nghiệp vụ")

    parts = []
    for nv in unique_nvs:
        matching = [c for c in df_raw.columns
                    if str(c).startswith(nv) and re.match(r'^[\.\(\)\d]*$', str(c)[len(nv):].strip())]
        if len(matching) < 3:
            print(f"      [Cảnh báo] Nghiệp vụ '{nv}' thiếu cột, chỉ thấy: {matching}")
            continue
        t = df_raw[id_cols + matching[:3]].copy()
        t.columns = id_cols + ['Muc_Do_Ho_Tro', 'Thoi_Gian_Xu_Ly', 'Loai_Loi']
        t['Ten_Nghiep_Vu'] = nv
        m = re.match(r'^(\d+)\.', nv)
        t['Thu_Tu_NV'] = int(m.group(1)) if m else 999
        parts.append(t)
    flat = pd.concat(parts, ignore_index=True).rename(columns=ID_COLS)

    # =================== 3. GẮN CỜ MÂU THUẪN ===================
    print("[3/6] Gắn cờ mâu thuẫn ĐVKD_Flag và thứ tự SLA")
    flat['ĐVKD_Flag'] = flat.apply(check_logic_mau_thuan, axis=1)   # tính trước khi fillna, như bản gốc
    for c in ['Muc_Do_Ho_Tro', 'Thoi_Gian_Xu_Ly', 'Loai_Loi']:
        flat[c] = flat[c].fillna('Không có dữ liệu').astype(str).str.strip()
    flat['Thu_Tu_SLA'] = flat['Thoi_Gian_Xu_Ly'].map(SLA_ORDER_MAP).fillna(99).astype(int)
    unmapped = sorted(set(flat['Thoi_Gian_Xu_Ly']) - set(SLA_ORDER_MAP))
    if unmapped:
        print("      [Cảnh báo] Thoi_Gian_Xu_Ly chưa có trong SLA_ORDER_MAP:", unmapped)
    for c in ["Ten_DVKD", "Loai_Hinh_DVKD", "Chuc_Danh"]:
        flat[c] = flat[c].fillna("Không xác định").astype(str).str.strip().replace("", "Không xác định")
    flat["Ma_NV"] = flat["Ma_NV"].fillna("").astype(str).str.strip()

    # =================== 4. GÓP Ý TH1/TH2 ===================
    print("[4/6] Unpivot góp ý TH1/TH2 và phân loại")
    gy_parts = []
    for p in (1, 2):
        sub = df_raw[id_cols].copy()
        for q in range(1, 12):
            hit = [c for c in df_raw.columns if re.match(rf"^TH{p}\.{q}(?!\d)", str(c))]
            if hit:
                sub[Q_NAMES[q - 1]] = df_raw[hit[0]]
            else:
                print(f"      [Cảnh báo] Không thấy cột TH{p}.{q}")
                sub[Q_NAMES[q - 1]] = pd.NA
        sub["Tinh_Huong"] = f"TH{p}"
        gy_parts.append(sub)
    gopy = pd.concat(gy_parts, ignore_index=True).rename(columns=ID_COLS)
    gopy["Trang_Thai_TinhHuong"] = gopy["Q1_MoTaLoi"].apply(check_error_signal)

    def has_signal(row):
        # chỉ ghép các câu có nội dung (tránh chuỗi "nan nan nan" bị tính là có lỗi)
        combined = " ".join(str(row[c]) for c in ["Q6_XuLyTam", "Q10_NguyenNhanGoc", "Q11_DeXuat"]
                            if pd.notna(row[c]))
        return check_error_signal(combined) == "Có phát sinh lỗi"

    gopy["Co_The_Bo_Sot"] = gopy.apply(has_signal, axis=1)
    for c in ["Ten_DVKD", "Loai_Hinh_DVKD", "Chuc_Danh"]:
        gopy[c] = gopy[c].fillna("Không xác định").astype(str).str.strip().replace("", "Không xác định")
    gopy["Ma_NV"] = gopy["Ma_NV"].fillna("").astype(str).str.strip()

    # =================== 5. TÁCH STAR SCHEMA ===================
    print("[5/6] Tách Dim và gắn khóa ngoại")
    dv_src = pd.concat([flat[["Ten_DVKD", "Loai_Hinh_DVKD"]], gopy[["Ten_DVKD", "Loai_Hinh_DVKD"]]])
    dim_dv = dv_src.drop_duplicates("Ten_DVKD").sort_values("Ten_DVKD").reset_index(drop=True)
    dim_dv.insert(0, "DonVi_ID", range(1, len(dim_dv) + 1))
    dim_dv.insert(1, "Ma_DVKD", dim_dv["Ten_DVKD"].str.extract(r"^(\d+)_", expand=False).fillna(""))
    multi = dv_src.drop_duplicates().groupby("Ten_DVKD").size()
    if (multi > 1).any():
        print("      [Cảnh báo] ĐVKD có nhiều hơn 1 loại hình:", list(multi[multi > 1].index))

    nv = flat[["Thu_Tu_NV", "Ten_Nghiep_Vu"]].drop_duplicates().sort_values("Thu_Tu_NV").reset_index(drop=True)
    assert nv["Thu_Tu_NV"].is_unique, "Có số thứ tự nghiệp vụ trùng tên khác nhau"
    dim_nv = pd.DataFrame({"NghiepVu_ID": nv["Thu_Tu_NV"], "Ten_Nghiep_Vu": nv["Ten_Nghiep_Vu"]})
    dim_nv["Ten_Ngan"] = dim_nv["Ten_Nghiep_Vu"].str.replace(r"^\d+\.\s*", "", regex=True)
    dim_nv["Nhom_Nghiep_Vu"] = dim_nv["NghiepVu_ID"].map(nhom_nv)
    dim_nv["Thu_Tu_NV"] = dim_nv["NghiepVu_ID"]

    dim_ht = pd.DataFrame({"Muc_Do_Ho_Tro": sorted(set(flat["Muc_Do_Ho_Tro"]))})
    info = dim_ht["Muc_Do_Ho_Tro"].map(ho_tro_info)
    dim_ht["Nhom_Ho_Tro"] = info.map(lambda t: t[0])
    dim_ht["Thu_Tu_HoTro"] = info.map(lambda t: t[1])
    dim_ht = dim_ht.sort_values(["Thu_Tu_HoTro", "Muc_Do_Ho_Tro"]).reset_index(drop=True)
    dim_ht.insert(0, "MucHoTro_ID", range(1, len(dim_ht) + 1))
    other = dim_ht.loc[dim_ht["Nhom_Ho_Tro"] == "Khác", "Muc_Do_Ho_Tro"].tolist()
    if other:
        print("      [Lưu ý] Mức hỗ trợ chưa có nhóm riêng (đang gom vào 'Khác'):", other)

    dim_sla = pd.DataFrame({"Thoi_Gian_Xu_Ly": sorted(set(flat["Thoi_Gian_Xu_Ly"]))})
    dim_sla["Thu_Tu_SLA"] = dim_sla["Thoi_Gian_Xu_Ly"].map(SLA_ORDER_MAP).fillna(99).astype(int)
    dim_sla = dim_sla.sort_values(["Thu_Tu_SLA", "Thoi_Gian_Xu_Ly"]).reset_index(drop=True)
    dim_sla.insert(0, "SLA_ID", range(1, len(dim_sla) + 1))

    dim_loi = pd.DataFrame({"Loai_Loi": sorted(set(flat["Loai_Loi"]))})
    info = dim_loi["Loai_Loi"].map(loi_info)
    dim_loi["Co_Loi"] = info.map(lambda t: t[0])
    dim_loi["Thu_Tu_LoaiLoi"] = info.map(lambda t: t[1])
    dim_loi = dim_loi.sort_values(["Thu_Tu_LoaiLoi", "Loai_Loi"]).reset_index(drop=True)
    dim_loi.insert(0, "LoaiLoi_ID", range(1, len(dim_loi) + 1))

    m_dv = dict(zip(dim_dv["Ten_DVKD"], dim_dv["DonVi_ID"]))
    fact_nv = flat.copy()
    fact_nv["DonVi_ID"] = fact_nv["Ten_DVKD"].map(m_dv)
    fact_nv["NghiepVu_ID"] = fact_nv["Thu_Tu_NV"]
    fact_nv["MucHoTro_ID"] = fact_nv["Muc_Do_Ho_Tro"].map(dict(zip(dim_ht["Muc_Do_Ho_Tro"], dim_ht["MucHoTro_ID"])))
    fact_nv["SLA_ID"] = fact_nv["Thoi_Gian_Xu_Ly"].map(dict(zip(dim_sla["Thoi_Gian_Xu_Ly"], dim_sla["SLA_ID"])))
    fact_nv["LoaiLoi_ID"] = fact_nv["Loai_Loi"].map(dict(zip(dim_loi["Loai_Loi"], dim_loi["LoaiLoi_ID"])))
    ma = ["Ma_NV"] if KEEP_MA_NV else []
    fact_nv = fact_nv[["Response_ID", "DonVi_ID"] + ma + ["Chuc_Danh", "NghiepVu_ID", "MucHoTro_ID",
                                                          "SLA_ID", "LoaiLoi_ID", "ĐVKD_Flag"]]

    gopy["DonVi_ID"] = gopy["Ten_DVKD"].map(m_dv)
    fact_gy = gopy[["Response_ID", "DonVi_ID"] + ma + ["Chuc_Danh", "Tinh_Huong"] + Q_NAMES +
                   ["Trang_Thai_TinhHuong", "Co_The_Bo_Sot"]]

    # =================== 6. KIỂM TRA & XUẤT ===================
    print("[6/6] Kiểm tra toàn vẹn và xuất CSV")
    for c in ["DonVi_ID", "NghiepVu_ID", "MucHoTro_ID", "SLA_ID", "LoaiLoi_ID"]:
        assert fact_nv[c].notna().all(), f"Fact_NghiepVu mất khóa ngoại {c}"
    assert fact_gy["DonVi_ID"].notna().all(), "Fact_GopY mất khóa ngoại DonVi_ID"
    assert not fact_nv.duplicated(["Response_ID", "NghiepVu_ID"]).any(), "Trùng grain Fact_NghiepVu"
    assert not fact_gy.duplicated(["Response_ID", "Tinh_Huong"]).any(), "Trùng grain Fact_GopY"
    assert len(fact_nv) == len(df_raw) * len(unique_nvs), "Số dòng fact != số phản hồi x số nghiệp vụ"
    assert len(fact_gy) == len(df_raw) * 2, "Số dòng Fact_GopY != số phản hồi x 2"

    back = pd.DataFrame({
        "Ten_DVKD": fact_nv["DonVi_ID"].map(dict(zip(dim_dv["DonVi_ID"], dim_dv["Ten_DVKD"]))),
        "Ten_Nghiep_Vu": fact_nv["NghiepVu_ID"].map(dict(zip(dim_nv["NghiepVu_ID"], dim_nv["Ten_Nghiep_Vu"]))),
        "Muc_Do_Ho_Tro": fact_nv["MucHoTro_ID"].map(dict(zip(dim_ht["MucHoTro_ID"], dim_ht["Muc_Do_Ho_Tro"]))),
        "Thoi_Gian_Xu_Ly": fact_nv["SLA_ID"].map(dict(zip(dim_sla["SLA_ID"], dim_sla["Thoi_Gian_Xu_Ly"]))),
        "Loai_Loi": fact_nv["LoaiLoi_ID"].map(dict(zip(dim_loi["LoaiLoi_ID"], dim_loi["Loai_Loi"]))),
        "ĐVKD_Flag": fact_nv["ĐVKD_Flag"]})
    for c in back.columns:
        assert (back[c].values == flat[c].values).all(), f"Round-trip lệch ở cột {c}"
    print("      Round-trip OK: ghép Fact + Dim lại khớp 100% với bảng phẳng trước khi tách")

    OUTPUT_DIR.mkdir(exist_ok=True)
    outs = [("Fact_NghiepVu", fact_nv), ("Fact_GopY", fact_gy), ("Dim_DonVi", dim_dv),
            ("Dim_NghiepVu", dim_nv), ("Dim_MucDoHoTro", dim_ht), ("Dim_SLA", dim_sla),
            ("Dim_LoaiLoi", dim_loi)]
    for name, df in outs:
        df.to_csv(OUTPUT_DIR / f"{name}.csv", index=False, encoding="utf-8-sig")
        print(f"      {name}.csv: {len(df)} dòng x {df.shape[1]} cột")

    print("\n--- Số liệu để đối soát Python / Excel / Power BI ---")
    print(f"Số phản hồi (ID duy nhất)         : {df_raw['ID'].nunique()}")
    print(f"Số nghiệp vụ                      : {len(unique_nvs)}")
    print(f"Số dòng Fact_NghiepVu             : {len(fact_nv)}  (= phản hồi x nghiệp vụ)")
    print(f"ĐVKD_Flag = Yes / No              : {int((fact_nv['ĐVKD_Flag']=='Yes').sum())} / {int((fact_nv['ĐVKD_Flag']=='No').sum())}")
    print("Số dòng theo Loại lỗi:")
    print(flat["Loai_Loi"].value_counts().rename_axis(None).to_string())
    print(f"Số dòng Fact_GopY                 : {len(fact_gy)}  (= phản hồi x tình huống)")
    print("Trang_Thai_TinhHuong:")
    print(fact_gy["Trang_Thai_TinhHuong"].value_counts().rename_axis(None).to_string())
    print(f"Co_The_Bo_Sot = True              : {int(fact_gy['Co_The_Bo_Sot'].sum())}")
    print("All checks passed.")


if __name__ == "__main__":
    main()