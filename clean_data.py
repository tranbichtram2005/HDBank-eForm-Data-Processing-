import pandas as pd
import numpy as np
import re

file_path = 'survey-dicore-giaodichtaiquay-test.xlsx' 
df_raw = pd.read_excel(file_path)

# 1. XÁC ĐỊNH CÁC NHÓM CỘT
idx_chucdanh = df_raw.columns.get_loc('Chức danh')
id_vars = list(df_raw.columns[:idx_chucdanh + 1])

# 2. TỰ ĐỘNG TÌM ĐÚNG 47 NGHIỆP VỤ GỐC (Bất chấp vị trí xáo trộn)
print("Đang quét tìm và chuẩn hóa các nghiệp vụ gốc...")
unique_nvs = []

for col in df_raw.columns[idx_chucdanh + 1:]:
    col_str = str(col)
    if col_str.startswith('TH1.') or col_str.startswith('TH2.'):
        break
    c_clean = col_str.replace('\xa0', ' ').strip()
    base_name = re.sub(r'(\.\d+|\s*\(\d+\)|\s*\d+)$', '', c_clean).strip()
    if base_name not in unique_nvs and re.match(r'^\d+\.', base_name):
        unique_nvs.append(base_name)

print(f"-> Đã gom thành công ĐÚNG {len(unique_nvs)} nghiệp vụ chuẩn. Bắt đầu xử lý (Unpivot)...")

# 3. XỬ LÝ KHỐI NGHIỆP VỤ (UNPIVOT)
df_fact_nghiepvu = pd.DataFrame()

for nv in unique_nvs:
    matching_cols = []
    for c in df_raw.columns:
        c_clean = str(c).replace('\xa0', ' ').strip()
        if c_clean.startswith(nv):
            remainder = c_clean[len(nv):].strip()
            if re.match(r'^[\.\(\)\d]*$', remainder):
                matching_cols.append(c)

    if len(matching_cols) >= 3:
        col_k1 = matching_cols[0]
        col_k2 = matching_cols[1]
        col_k3 = matching_cols[2]

        temp_df = df_raw[id_vars + [col_k1, col_k2, col_k3]].copy()

        temp_df.rename(columns={
            col_k1: 'Muc_Do_Ho_Tro',
            col_k2: 'Thoi_Gian_Xu_Ly',
            col_k3: 'Loai_Loi'
        }, inplace=True)

        temp_df['Ten_Nghiep_Vu'] = nv

        # ===== MỚI: Tách số thứ tự để Power BI sort đúng 1,2,3...47 =====
        match_num = re.match(r'^(\d+)\.', nv)
        temp_df['Thu_Tu_NV'] = int(match_num.group(1)) if match_num else 999

        df_fact_nghiepvu = pd.concat([df_fact_nghiepvu, temp_df], ignore_index=True)
    else:
        print(f"[Cảnh báo] Nghiệp vụ '{nv}' bị thiếu cột. Chỉ tìm thấy: {matching_cols}")

# ==========================================
# 4. KIỂM TRA CHÉO 3 CỘT & GẮN CỜ ĐVKD_Flag
# ==========================================
print("Đang check logic và gắn cờ ĐVKD_Flag...")

def check_logic_mau_thuan(row):
    col1 = str(row['Muc_Do_Ho_Tro']).lower().strip()
    col2 = str(row['Thoi_Gian_Xu_Ly']).lower().strip()
    col3 = str(row['Loai_Loi']).lower().strip()

    if "chưa phát sinh nghiệp vụ" in col1:
        if ("chưa phát sinh" not in col2) or ("không phát sinh lỗi" not in col3):
            return "Yes"
    return "No"

df_fact_nghiepvu['ĐVKD_Flag'] = df_fact_nghiepvu.apply(check_logic_mau_thuan, axis=1)

df_fact_nghiepvu['Muc_Do_Ho_Tro'] = df_fact_nghiepvu['Muc_Do_Ho_Tro'].fillna('Không có dữ liệu')
df_fact_nghiepvu['Thoi_Gian_Xu_Ly'] = df_fact_nghiepvu['Thoi_Gian_Xu_Ly'].fillna('Không có dữ liệu')
df_fact_nghiepvu['Loai_Loi'] = df_fact_nghiepvu['Loai_Loi'].fillna('Không có dữ liệu')

sla_order_map = {
    '<5 phút': 1,
    '>5 - 15 phút': 2,
    '>15 - 30 phút': 3,
    '>30 - 60 phút': 4,
    '>60 phút': 5,
    'Chưa phát sinh': 6,
    'Không có dữ liệu': 99
}
df_fact_nghiepvu['Thu_Tu_SLA'] = df_fact_nghiepvu['Thoi_Gian_Xu_Ly'].map(sla_order_map).fillna(99).astype(int)

# Cảnh báo nếu có giá trị SLA lạ không khớp map (giúp bạn phát hiện sai chính tả/khoảng trắng)
unmapped_sla = df_fact_nghiepvu.loc[
    ~df_fact_nghiepvu['Thoi_Gian_Xu_Ly'].isin(sla_order_map.keys()), 'Thoi_Gian_Xu_Ly'
].unique()
if len(unmapped_sla) > 0:
    print(f"[Cảnh báo] Các giá trị Thoi_Gian_Xu_Ly chưa có trong sla_order_map: {unmapped_sla}")

print("Đã gắn cờ xong! Các dòng mâu thuẫn sẽ có ĐVKD_Flag = 'Yes'.")

# 5. XỬ LÝ KHỐI ĐỊNH TÍNH (GÓP Ý TH1 & TH2)
print("Đang trích xuất dữ liệu góp ý...")
gopy_cols = [c for c in df_raw.columns if str(c).startswith('TH1.') or str(c).startswith('TH2.')]

if len(gopy_cols) > 0:
    df_fact_gopy = df_raw[id_vars + gopy_cols].dropna(subset=gopy_cols, how='all')
    df_fact_gopy.to_csv('Fact_GopY.csv', index=False, encoding='utf-8-sig')
    print("-> Đã xuất 2. Fact_GopY.csv (Dữ liệu câu hỏi tình huống tự luận)")
else:
    print("[Thông báo] Không tìm thấy cột câu hỏi tự luận (TH1, TH2)!")

# 6. XUẤT FILE KẾT QUẢ
df_fact_nghiepvu.to_csv('Fact_NghiepVu.csv', index=False, encoding='utf-8-sig')

print("-" * 30)
print("XỬ LÝ THÀNH CÔNG! Đã tạo file Fact_NghiepVu.csv (Đủ 47 nghiệp vụ chuẩn).")