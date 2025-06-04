import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Đọc các file dữ liệu (cập nhật đường dẫn nếu khác)
karma = pd.read_excel('[FANPAGE KARMA] The Body Shop.xlsx', sheet_name='Metrics Overview')
apify = pd.read_excel('[APIFY] The Body Shop.xlsx')
fastmoss_video = pd.read_excel('[FASTMOSS] The Body Shop.xlsx', sheet_name='Data Video')
fastmoss_live = pd.read_excel('[FASTMOSS] The Body Shop.xlsx', sheet_name='Data Livestream')

# CHUẨN HÓA DỮ LIỆU
karma['Date'] = pd.to_datetime(karma['Date'])
apify['createTime'] = pd.to_datetime(apify['createTime'], unit='s')
fastmoss_video['Ngày đăng'] = pd.to_datetime(fastmoss_video['Ngày đăng'])
fastmoss_live['Ngày'] = pd.to_datetime(fastmoss_live['Ngày'])

# 1️⃣ Phân tích TƯƠNG TÁC TRÊN FACEBOOK (FANPAGE KARMA)
plt.figure(figsize=(10,5))
sns.lineplot(data=karma, x='Date', y='Engagement')
plt.title("Tương tác tổng thể theo ngày (Facebook)")
plt.show()

# 2️⃣ Phân tích TOP VIDEO TIKTOK (APIFY)
top_tiktok = apify.sort_values('playCount', ascending=False).head(10)
print("🔥 Top 10 video nhiều lượt xem nhất:")
print(top_tiktok[['desc', 'playCount', 'diggCount', 'shareCount']])

# 3️⃣ Phân tích VIDEO vs LIVESTREAM (FASTMOSS)
video_by_date = fastmoss_video.groupby('Ngày đăng')['Lượt xem'].sum()
live_by_date = fastmoss_live.groupby('Ngày')['Lượt xem'].sum()

plt.figure(figsize=(10,5))
video_by_date.plot(label='Video', marker='o')
live_by_date.plot(label='Livestream', marker='x')
plt.legend()
plt.title('So sánh lượt xem Video và Livestream')
plt.ylabel('Lượt xem')
plt.show()
