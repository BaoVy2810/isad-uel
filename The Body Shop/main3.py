import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# Function to safely read Excel sheets
def safe_read_excel(file_path, sheet_name=None):
    try:
        if sheet_name:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            return pd.read_excel(file_path)
    except ValueError as e:
        print(f"⚠️ Warning: {e}")
        return None
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {file_path}")
        return None


# Function to list available sheets in Excel file
def list_excel_sheets(file_path):
    try:
        xl_file = pd.ExcelFile(file_path)
        return xl_file.sheet_names
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return []


# Đọc các file dữ liệu với xử lý lỗi
print("📊 Đang đọc dữ liệu...")

karma = safe_read_excel('[FANPAGE KARMA] The Body Shop.xlsx', sheet_name='Metrics Overview')
apify = safe_read_excel('[APIFY] The Body Shop.xlsx')

# Check available sheets in FASTMOSS file
print("\n🔍 Kiểm tra các sheet có sẵn trong file FASTMOSS:")
fastmoss_sheets = list_excel_sheets('[FASTMOSS] The Body Shop.xlsx')
print(f"Available sheets: {fastmoss_sheets}")

fastmoss_video = safe_read_excel('[FASTMOSS] The Body Shop.xlsx', sheet_name='Data Video')
fastmoss_live = safe_read_excel('[FASTMOSS] The Body Shop.xlsx', sheet_name='Data Livestream')

# Try to read product data with different possible sheet names
fastmoss_product = None
possible_product_sheets = ['Data Product ', 'Data Product', 'Product', 'Products', 'Sản phẩm', 'Data Sản phẩm']

for sheet_name in possible_product_sheets:
    if sheet_name in fastmoss_sheets:
        fastmoss_product = safe_read_excel('[FASTMOSS] The Body Shop.xlsx', sheet_name=sheet_name)
        print(f"✅ Found product data in sheet: {sheet_name}")
        break

if fastmoss_product is None:
    print("⚠️ Warning: Product data sheet not found. Continuing without product analysis.")

# ===== CHUẨN HÓA DỮ LIỆU =====
print("\n🔧 Đang chuẩn hóa dữ liệu...")


# Function to safely convert datetime columns
def safe_datetime_convert(df, column_name, unit=None):
    if df is not None and column_name in df.columns:
        try:
            if unit:
                df[column_name] = pd.to_datetime(df[column_name], unit=unit)
            else:
                df[column_name] = pd.to_datetime(df[column_name])
            print(f"✅ Converted {column_name} to datetime")
        except Exception as e:
            print(f"⚠️ Warning: Could not convert {column_name} to datetime: {e}")
    else:
        if df is not None:
            print(f"⚠️ Warning: Column '{column_name}' not found. Available columns: {df.columns.tolist()}")


# Chuyển kiểu ngày tháng (chỉ nếu dữ liệu tồn tại)
# ✅ Kiểm tra Karma sheet nếu toàn cột Unnamed
if karma is not None:
    print(f"Karma columns: {karma.columns.tolist()}")
    if all(col.startswith('Unnamed') for col in karma.columns):
        print("⚠️ Warning: All columns in Karma are unnamed. Trying with header=1...")
        try:
            karma = pd.read_excel('[FANPAGE KARMA] The Body Shop.xlsx', sheet_name='Metrics Overview', header=1)
            print(f"✅ New Karma columns: {karma.columns.tolist()}")
        except Exception as e:
            print(f"❌ Failed to re-read Karma with header=1: {e}")

    # Try common date column names
    date_columns = ['Date', 'date', 'Ngày', 'Time', 'Timestamp']
    for col in date_columns:
        if col in karma.columns:
            safe_datetime_convert(karma, col)
            break

if apify is not None:
    safe_datetime_convert(apify, 'createTime', unit='s')

# ✅ Đổi tên cột đúng cho ngày đăng
if fastmoss_video is not None:
    safe_datetime_convert(fastmoss_video, 'Thời gian phát hành')  # Sửa từ 'Ngày đăng'

if fastmoss_live is not None:
    safe_datetime_convert(fastmoss_live, 'Thời gian bắt đầu Livestream')  # Sửa từ 'Ngày'


# Làm sạch đơn vị tiền tệ (VND)
def clean_currency(value):
    try:
        if pd.isna(value):
            return None
        if isinstance(value, str):
            value = value.replace("₫", "").replace(".", "").replace(",", ".").replace(" ", "").lower()
            if "tr" in value:
                return float(value.replace("tr", "")) * 1_000_000
            elif "k" in value:
                return float(value.replace("k", "")) * 1_000
            else:
                return float(value)
        return float(value)
    except:
        return None


# Áp dụng vào cột doanh số (chỉ nếu dữ liệu tồn tại)
if fastmoss_video is not None and 'Doanh số bán hàng của video' in fastmoss_video.columns:
    fastmoss_video['Doanh số (VND)'] = fastmoss_video['Doanh số bán hàng của video'].apply(clean_currency)

if fastmoss_live is not None and 'Doanh số Livestream' in fastmoss_live.columns:
    fastmoss_live['Doanh số (VND)'] = fastmoss_live['Doanh số Livestream'].apply(clean_currency)

if fastmoss_product is not None and 'Doanh số' in fastmoss_product.columns:
    fastmoss_product['Doanh số (VND)'] = fastmoss_product['Doanh số'].apply(clean_currency)

# ===== 1️⃣ PHÂN TÍCH TƯƠNG TÁC TRÊN FACEBOOK (KARMA) =====
if karma is not None:
    print("\n📈 Đang tạo biểu đồ Facebook engagement...")

    # Find the date column and engagement column
    date_col = None
    engagement_col = None

    # Look for date column
    for col in karma.columns:
        if any(keyword in col.lower() for keyword in ['date', 'ngày', 'time']):
            date_col = col
            break

    # Look for engagement column
    for col in karma.columns:
        if any(keyword in col.lower() for keyword in ['engagement', 'tương tác', 'interact']):
            engagement_col = col
            break

    if date_col and engagement_col:
        plt.figure(figsize=(10, 5))
        sns.lineplot(data=karma, x=date_col, y=engagement_col)
        plt.title("Tương tác tổng thể theo ngày (Facebook)")
        plt.xlabel("Ngày")
        plt.ylabel("Engagement")
        plt.grid(True)
        plt.tight_layout()
        plt.show()
    else:
        print(f"⚠️ Could not find date or engagement columns. Available columns: {karma.columns.tolist()}")
else:
    print("⚠️ Bỏ qua phân tích Facebook - dữ liệu không khả dụng")

# ===== 2️⃣ PHÂN TÍCH TOP VIDEO TIKTOK (APIFY) =====
# ✅ Phân tích TikTok videos
if apify is not None:
    print("🎥 Phân tích TikTok videos...")
    # Sắp xếp video theo lượt xem giảm dần
    if 'playCount' in apify.columns:
        top_tiktok = apify.sort_values(by='playCount', ascending=False).head(10)
        print("🔥 Top 10 video TikTok nhiều lượt xem nhất:")
        # ✅ Kiểm tra các cột có tồn tại trước khi in
        expected_columns = ['desc', 'playCount', 'diggCount', 'shareCount']
        available_columns = [col for col in expected_columns if col in top_tiktok.columns]
        if available_columns:
            print(top_tiktok[available_columns])
        else:
            print(f"⚠️ Không tìm thấy các cột mong đợi trong Apify: {expected_columns}")
            print(f"📋 Cột hiện có: {apify.columns.tolist()}")
    else:
        print("⚠️ Không có cột 'playCount' trong dữ liệu TikTok.")

    # Tính tương tác trung bình theo tuần TikTok
    if 'createTime' in apify.columns:
        apify['week'] = apify['createTime'].dt.to_period("W").astype(str)

        # ✅ Kiểm tra các cột cần thiết cho groupby
        groupby_columns = ['playCount', 'diggCount', 'shareCount', 'commentCount']
        available_groupby_columns = {col: 'sum' for col in groupby_columns if col in apify.columns}

        if available_groupby_columns:
            apify_grouped = apify.groupby('week').agg(available_groupby_columns).reset_index()

            # Tính tỉ lệ tương tác nếu có đủ cột
            if 'playCount' in apify_grouped.columns and 'diggCount' in apify_grouped.columns:
                apify_grouped["Tỉ lệ tương tác (%)"] = (apify_grouped["diggCount"] / apify_grouped["playCount"]) * 100

                # Tính tỉ lệ tương tác tổng
                interaction_cols = ['diggCount', 'shareCount', 'commentCount']
                available_interaction_cols = [col for col in interaction_cols if col in apify_grouped.columns]
                if available_interaction_cols:
                    total_interactions = apify_grouped[available_interaction_cols].sum(axis=1)
                    apify_grouped["Tỉ lệ tương tác tổng (%)"] = (total_interactions / apify_grouped["playCount"]) * 100
        else:
            print(f"⚠️ Không tìm thấy cột cần thiết để tính tương tác: {groupby_columns}")
            apify_grouped = None
    else:
        print("⚠️ Không có cột 'createTime' để phân tích theo tuần")
        apify_grouped = None
else:
    print("⚠️ Bỏ qua phân tích TikTok - dữ liệu không khả dụng")
    apify_grouped = None

# ===== 3️⃣ SO SÁNH VIDEO vs LIVESTREAM (FASTMOSS) =====
if fastmoss_video is not None and fastmoss_live is not None:
    print("\n📊 So sánh Video vs Livestream...")

    # ✅ Tìm cột ngày phù hợp
    video_date_col = None
    live_date_col = None

    # Tìm cột ngày cho video
    for col in ['Thời gian phát hành', 'Ngày đăng', 'Date']:
        if col in fastmoss_video.columns:
            video_date_col = col
            break

    # Tìm cột ngày cho livestream
    for col in ['Thời gian bắt đầu Livestream', 'Ngày', 'Date']:
        if col in fastmoss_live.columns:
            live_date_col = col
            break

    # Kiểm tra cột lượt xem
    video_view_col = 'Lượt xem' if 'Lượt xem' in fastmoss_video.columns else None
    live_view_col = 'Lượt xem' if 'Lượt xem' in fastmoss_live.columns else None

    if video_date_col and live_date_col and video_view_col and live_view_col:
        video_by_date = fastmoss_video.groupby(video_date_col)[video_view_col].sum()
        live_by_date = fastmoss_live.groupby(live_date_col)[live_view_col].sum()

        plt.figure(figsize=(10, 5))
        video_by_date.plot(label='Video thường', marker='o')
        live_by_date.plot(label='Livestream', marker='x')
        plt.legend()
        plt.title('So sánh lượt xem Video thường và Livestream theo ngày')
        plt.ylabel('Lượt xem')
        plt.xlabel('Ngày')
        plt.grid(True)
        plt.tight_layout()
        plt.show()
    else:
        print(f"⚠️ Không tìm thấy cột cần thiết:")
        print(f"Video - Date: {video_date_col}, Views: {video_view_col}")
        print(f"Live - Date: {live_date_col}, Views: {live_view_col}")
        print(f"Video columns: {fastmoss_video.columns.tolist()}")
        print(f"Live columns: {fastmoss_live.columns.tolist()}")
else:
    print("⚠️ Bỏ qua so sánh Video vs Livestream - dữ liệu không đầy đủ")

# ===== 4️⃣ XUẤT FILE EXCEL PHÂN TÍCH =====
print("\n💾 Đang xuất file Excel...")

# Prepare data for export (only if available)
sheets_to_export = {}

if karma is not None:
    sheets_to_export["Facebook Karma"] = karma

if fastmoss_video is not None:
    video_summary = fastmoss_video.copy()

    # ✅ Tìm cột ngày phù hợp
    video_date_col = None
    for col in ['Thời gian phát hành', 'Ngày đăng', 'Date']:
        if col in video_summary.columns:
            video_date_col = col
            break

    if video_date_col:
        video_summary["week"] = video_summary[video_date_col].dt.to_period("W").astype(str)

        # ✅ Kiểm tra các cột cần thiết
        agg_dict = {}
        if 'Lượt xem' in video_summary.columns:
            agg_dict['Lượt xem'] = 'sum'
        if 'Số lượng likes' in video_summary.columns:
            agg_dict['Số lượng likes'] = 'sum'
        if 'Doanh số (VND)' in video_summary.columns:
            agg_dict['Doanh số (VND)'] = 'sum'

        if agg_dict:
            video_grouped = video_summary.groupby("week").agg(agg_dict).reset_index()

            # Tính tỉ lệ tương tác nếu có đủ cột
            if "Số lượng likes" in video_grouped.columns and "Lượt xem" in video_grouped.columns:
                video_grouped["Tỉ lệ tương tác (%)"] = (video_grouped["Số lượng likes"] / video_grouped[
                    "Lượt xem"]) * 100

            sheets_to_export["FASTMOSS Video"] = video_grouped
        else:
            print("⚠️ Không tìm thấy cột dữ liệu cần thiết cho video analysis")

if fastmoss_live is not None:
    live_summary = fastmoss_live.copy()

    # ✅ Tìm cột ngày phù hợp
    live_date_col = None
    for col in ['Thời gian bắt đầu Livestream', 'Ngày', 'Date']:
        if col in live_summary.columns:
            live_date_col = col
            break

    if live_date_col:
        live_summary["week"] = live_summary[live_date_col].dt.to_period("W").astype(str)

        # ✅ Kiểm tra các cột cần thiết
        agg_dict = {}
        if 'Lượt xem' in live_summary.columns:
            agg_dict['Lượt xem'] = 'sum'
        if 'Doanh số (VND)' in live_summary.columns:
            agg_dict['Doanh số (VND)'] = 'sum'

        if agg_dict:
            live_grouped = live_summary.groupby("week").agg(agg_dict).reset_index()
            sheets_to_export["Livestream"] = live_grouped
        else:
            print("⚠️ Không tìm thấy cột dữ liệu cần thiết cho livestream analysis")

if apify_grouped is not None:
    sheets_to_export["TikTok"] = apify_grouped

if fastmoss_product is not None:
    sheets_to_export["Product Data"] = fastmoss_product

# Export to Excel
if sheets_to_export:
    with pd.ExcelWriter("TheBodyShop_Insights.xlsx") as writer:
        for sheet_name, data in sheets_to_export.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

    print("✅ Đã xuất file Excel: TheBodyShop_Insights.xlsx")
    print(f"📋 Các sheet đã xuất: {list(sheets_to_export.keys())}")
else:
    print("❌ Không có dữ liệu để xuất")

print("\n🎉 Hoàn thành phân tích!")