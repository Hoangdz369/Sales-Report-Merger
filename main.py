import pandas as pd

# Đọc file excel
df_shopee = pd.read_excel("Shopee.xlsx") #Thay tên file excel của bạn vào đây
df_facebook = pd.read_excel("Facebook.xlsx") #Thay tên file excel của bạn vào đây
df_tiktok = pd.read_excel("TikTok.xlsx") #Thay tên file excel của bạn vào đây

# Tạo cột mẫu cho bảng df shopee mới
mapping_shopee = {
    "Ma_don": "Ma_don",
    "Ten_KH": "Khach_hang",
    "SL": "So_luong",
    "Gia_ban": "Don_gia",
    "Ngay": "Ngay",
}
# Tạo cột mẫu cho bảng df facebook mới
mapping_facebook = {
    "Order ID": "Ma_don",
    "Customer": "Khach_hang",
    "Quantity": "So_luong",
    "Price": "Don_gia",
    "Date": "Ngay",
}
# Tạo cột mẫu cho bảng df tiktok mới
mapping_tiktok = {
    "id_dh": "Ma_don",
    "khach_hang": "Khach_hang",
    "so_luong": "So_luong",
    "don_gia": "Don_gia",
    "ngay_dat": "Ngay",
}

# Hàm xử lý
def xu_ly_chung(df_shopee, df_facebook, df_tiktok, mapping_shopee, mapping_facebook, mapping_tiktok):
    # Set cột mặc định cần lấy từ dữ liệu gốc
    df_shopee = df_shopee[mapping_shopee.keys()]
    df_facebook = df_facebook[mapping_facebook.keys()]
    df_tiktok = df_tiktok[mapping_tiktok.keys()]

    # Xử lý làm chung các bảng có cùng tên cột
    df_shopee = df_shopee.rename(columns=mapping_shopee)
    df_facebook = df_facebook.rename(columns=mapping_facebook)
    df_tiktok = df_tiktok.rename(columns=mapping_tiktok)

    # ///// Xử lý làm sach dữ liệu
    # Xử lý định dạng ngày
    # Shopee
    date_shopee = pd.to_datetime(
        df_shopee["Ngay"],
        format="%Y/%m/%d",
        errors = "coerce",
    ).notna().all()

    if not date_shopee:
        df_shopee["Ngay"] = pd.to_datetime(df_shopee["Ngay"], dayfirst=True)
    else: 
        df_shopee["Ngay"] = pd.to_datetime(df_shopee["Ngay"])

    # Facebook
    date_facebook = pd.to_datetime(
        df_facebook["Ngay"],
        format="%Y-%m-%d",
        errors="coerce"
    ).notna().all()

    if not date_facebook:
        df_facebook["Ngay"] = pd.to_datetime(df_facebook["Ngay"], dayfirst=True)
    else: 
        df_facebook["Ngay"] = pd.to_datetime(df_facebook["Ngay"])

    # Tiktok
    date_tiktok = pd.to_datetime(
        df_tiktok["Ngay"],
        format="%Y/%m/%d",
        errors="coerce"
    ).notna().all()

    if not date_tiktok:
        df_tiktok["Ngay"] = pd.to_datetime(df_tiktok["Ngay"], dayfirst=True)
    else: 
        df_tiktok["Ngay"] = pd.to_datetime(df_tiktok["Ngay"])

    # Thêm cột kênh tmđt
    df_shopee["Kenh"] = "Shopee"
    df_facebook["Kenh"] = "Facebook"
    df_tiktok["Kenh"] = "Tiktok"

    # Gộp ba bảng với nhau
    df_all = pd.concat([df_shopee, df_facebook, df_tiktok], axis=0).reset_index(drop=True)

    # Xử lý dữ liệu thừa & đổi kiểu dữ liệu 
    df_main_price = []
    df_price = [price for price in df_all["Don_gia"]]

    for each_price in df_price:
        if str(each_price).isdigit():
            df_main_price.append(each_price)
        else:
            price = [each_word for each_word in each_price if each_word.isdigit()]
            price = "".join(price)
            df_main_price.append(price)

    df_all["Don_gia"] = df_main_price
    df_all["Don_gia"] = pd.to_numeric(df_all["Don_gia"], errors="coerce")

    # Kiểm tra thiếu dữ liệu trong bảng & thêm ô thiếu dữ liệu
    df_null = df_all[['Ma_don', 'Khach_hang', 'So_luong', 'Don_gia', 'Ngay']].isna().any(axis=1)
    # Thêm ô thiếu dữ liệu
    df_all["Thieu_du_lieu"] = df_null.map({True: "Can kiem tra", False: ""})

    # Xử lý thành tiền
    df_all["So_luong"] = df_all["So_luong"].astype("Int64")
    df_all["Thanh_tien"] = df_all["So_luong"] * df_all["Don_gia"]

    return df_all

def export_excel(df_export_excel):
    
    # Xử lý dữ liệu để tạo sheet doanh thu theo kênh
    df_thieu = df_export_excel[df_export_excel["Thieu_du_lieu"] == ""]
    df_export_excel["Ngay"] = df_export_excel["Ngay"].dt.strftime("%Y-%m-%d")
    df_doanh_thu_kenh = df_thieu.groupby("Kenh")["Thanh_tien"].sum()

    # Export excel
    with pd.ExcelWriter("file_export_excel_04.xlsx") as writer:
        df_export_excel.to_excel(writer, sheet_name="Chi tiết doanh thu", index=False)
        df_doanh_thu_kenh.to_excel(writer,sheet_name="Doanh thu theo kênh")


# Data frame chi tiết doanh thu 
df_export = xu_ly_chung(df_shopee, df_facebook, df_tiktok, mapping_shopee, mapping_facebook, mapping_tiktok)
# Xuất file excel
export_excel(df_export_excel=df_export)
