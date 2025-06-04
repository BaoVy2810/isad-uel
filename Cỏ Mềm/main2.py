import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Đọc dữ liệu
apify = pd.read_excel("[APIFY] CỎ MỀM.xlsx")
fastmoss_df = pd.read_excel("[FASTMOSS] Cỏ Mềm.xlsx")
fanpage = pd.read_excel("[FANPAGE KARMA] CỎ MỀM.xlsx", header=4)
print(fanpage.columns)

# Xử lý Apify
apify['timestamp'] = pd.to_datetime(apify['createTimeISO'], errors='coerce')

def extract_hashtags(text):
    if isinstance(text, str):
        return [tag.strip("#") for tag in text.split() if tag.startswith("#")]
    return []

apify['hashtags'] = apify['text'].apply(extract_hashtags)
all_hashtags = sum(apify['hashtags'].dropna(), [])
top_hashtags = pd.Series(all_hashtags).value_counts().head(10)
apify['hour'] = apify['timestamp'].dt.hour
post_by_hour = apify['hour'].value_counts().sort_index()

# Hàm chuyển đổi số từ FastMoss
def convert_number(s):
    if isinstance(s, str):
        s = s.replace('.', '').replace(',', '.').lower().strip()
        if 'tr' in s or 'triệu' in s or 'm' in s:
            s = s.replace('tr', '').replace('triệu', '').replace('m', '')
            return float(s) * 1_000_000
        elif 'k' in s or 'nghìn' in s:
            s = s.replace('k', '').replace('nghìn', '')
            return float(s) * 1_000
        else:
            return float(s)
    return s

# Xử lý FastMoss
fastmoss = fastmoss_df.copy()
fastmoss['Lượt xem'] = fastmoss['Lượt xem'].apply(convert_number)
fastmoss['[90 ngày gần đây]Lượt thích'] = fastmoss['[90 ngày gần đây]Lượt thích'].apply(convert_number)
fastmoss['Thời gian đăng'] = pd.to_datetime(fastmoss['Thời gian đăng'].str.extract(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')[0], errors='coerce')
fastmoss['hour'] = fastmoss['Thời gian đăng'].dt.hour
top_categories = fastmoss.groupby('Phân loại KOC/KOL')['Lượt xem'].mean().sort_values(ascending=False).head(10)

# Vẽ biểu đồ
plt.figure(figsize=(16, 10))

plt.subplot(2, 2, 1)
sns.barplot(x=post_by_hour.index, y=post_by_hour.values, palette='viridis')
plt.title("Post Frequency by Hour (TikTok)")
plt.xlabel("Hour of Day")
plt.ylabel("Number of Posts")

plt.subplot(2, 2, 2)
sns.barplot(x=top_hashtags.values, y=top_hashtags.index, palette='magma')
plt.title("Top 10 Trending Hashtags")
plt.xlabel("Frequency")
plt.ylabel("Hashtag")

plt.subplot(2, 1, 2)
sns.barplot(x=top_categories.values, y=top_categories.index, palette="coolwarm")
plt.title("Average Views per KOL Category (FastMoss)")
plt.xlabel("Average Views")
plt.ylabel("KOC/KOL Category")

# Sắp xếp theo tỷ lệ tương tác
top_fanpages = fanpage.sort_values(by='Post interaction rate', ascending=False).head(10)

plt.figure(figsize=(12, 6))
sns.barplot(x='Post interaction rate', y='Profile', data=top_fanpages, palette='crest')
plt.title("Top 10 Fanpages by Post Interaction Rate")
plt.xlabel("Interaction Rate (%)")
plt.ylabel("Fanpage")

plt.tight_layout()
plt.show()
