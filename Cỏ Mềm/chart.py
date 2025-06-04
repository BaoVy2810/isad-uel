import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# Set up plotting style
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")


def load_and_process_data():
    """Load and process all data sources"""
    try:
        # Load data files
        apify = pd.read_excel("[APIFY] CỎ MỀM.xlsx")
        fastmoss_df = pd.read_excel("[FASTMOSS] Cỏ Mềm.xlsx")
        fanpage = pd.read_excel("[FANPAGE KARMA] CỎ MỀM.xlsx", header=4)

        print("Data loaded successfully!")
        print(f"Apify records: {len(apify)}")
        print(f"FastMoss records: {len(fastmoss_df)}")
        print(f"Fanpage records: {len(fanpage)}")

        return apify, fastmoss_df, fanpage

    except FileNotFoundError as e:
        print(f"File not found: {e}")
        return None, None, None
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None


def process_apify_data(apify):
    """Process TikTok data from Apify"""
    if apify is None:
        return None, None, None

    # Convert timestamp
    apify['timestamp'] = pd.to_datetime(apify['createTimeISO'], errors='coerce')

    # Extract hashtags
    def extract_hashtags(text):
        if isinstance(text, str):
            return [tag.strip("#").lower() for tag in text.split() if tag.startswith("#")]
        return []

    apify['hashtags'] = apify['text'].apply(extract_hashtags)

    # Get top hashtags
    all_hashtags = [tag for sublist in apify['hashtags'].dropna() for tag in sublist]
    top_hashtags = pd.Series(all_hashtags).value_counts().head(10)

    # Post frequency by hour
    apify['hour'] = apify['timestamp'].dt.hour
    post_by_hour = apify['hour'].value_counts().sort_index()

    # Additional metrics
    apify['day_of_week'] = apify['timestamp'].dt.day_name()
    posts_by_day = apify['day_of_week'].value_counts()

    return apify, top_hashtags, post_by_hour, posts_by_day


def convert_vietnamese_numbers(s):
    """Convert Vietnamese number formats to numeric values"""
    if pd.isna(s) or s == '':
        return 0

    if isinstance(s, (int, float)):
        return s

    if isinstance(s, str):
        s = s.replace('.', '').replace(',', '.').lower().strip()

        # Handle Vietnamese abbreviations
        multipliers = {
            'tỷ': 1_000_000_000,
            'tr': 1_000_000,
            'triệu': 1_000_000,
            'm': 1_000_000,
            'k': 1_000,
            'nghìn': 1_000
        }

        for abbr, mult in multipliers.items():
            if abbr in s:
                try:
                    number = float(s.replace(abbr, '').strip())
                    return number * mult
                except ValueError:
                    continue

        # Try to convert directly
        try:
            return float(s)
        except ValueError:
            return 0

    return 0


def process_fastmoss_data(fastmoss_df):
    """Process FastMoss influencer data"""
    if fastmoss_df is None:
        return None, None

    fastmoss = fastmoss_df.copy()

    # Convert numeric columns
    numeric_columns = ['Lượt xem', '[90 ngày gần đây]Lượt thích', 'Lượt theo dõi']
    for col in numeric_columns:
        if col in fastmoss.columns:
            fastmoss[col] = fastmoss[col].apply(convert_vietnamese_numbers)

    # Process timestamps
    if 'Thời gian đăng' in fastmoss.columns:
        # Extract datetime from Vietnamese format
        fastmoss['Thời gian đăng'] = pd.to_datetime(
            fastmoss['Thời gian đăng'].str.extract(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')[0],
            errors='coerce'
        )
        fastmoss['hour'] = fastmoss['Thời gian đăng'].dt.hour

    # Calculate engagement rate
    if 'Lượt xem' in fastmoss.columns and '[90 ngày gần đây]Lượt thích' in fastmoss.columns:
        fastmoss['engagement_rate'] = (fastmoss['[90 ngày gần đây]Lượt thích'] /
                                       fastmoss['Lượt xem'] * 100).fillna(0)

    # Top categories by views
    if 'Phân loại KOC/KOL' in fastmoss.columns:
        top_categories = (fastmoss.groupby('Phân loại KOC/KOL')['Lượt xem']
                          .mean().sort_values(ascending=False).head(10))
    else:
        top_categories = pd.Series()

    return fastmoss, top_categories


def process_fanpage_data(fanpage):
    """Process Fanpage Karma data"""
    if fanpage is None:
        return None

    # Clean interaction rate column
    if 'Post interaction rate' in fanpage.columns:
        fanpage['Post interaction rate'] = pd.to_numeric(
            fanpage['Post interaction rate'], errors='coerce'
        )
        top_fanpages = (fanpage.sort_values(by='Post interaction rate', ascending=False)
                        .head(10))
    else:
        print("Available columns:", fanpage.columns.tolist())
        top_fanpages = fanpage.head(10)

    return top_fanpages


def create_enhanced_visualizations(apify, top_hashtags, post_by_hour, posts_by_day,
                                   fastmoss, top_categories, top_fanpages):
    """Create comprehensive visualizations"""

    # Set up the figure with subplots
    fig = plt.figure(figsize=(20, 16))

    # 1. Post frequency by hour (TikTok)
    plt.subplot(3, 3, 1)
    if post_by_hour is not None and not post_by_hour.empty:
        bars = plt.bar(post_by_hour.index, post_by_hour.values,
                       color=plt.cm.viridis(np.linspace(0, 1, len(post_by_hour))))
        plt.title("📱 TikTok Posts by Hour", fontsize=12, fontweight='bold')
        plt.xlabel("Hour of Day")
        plt.ylabel("Number of Posts")
        plt.xticks(range(0, 24, 2))

        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontsize=8)

    # 2. Top hashtags
    plt.subplot(3, 3, 2)
    if top_hashtags is not None and not top_hashtags.empty:
        colors = plt.cm.magma(np.linspace(0.2, 0.8, len(top_hashtags)))
        bars = plt.barh(range(len(top_hashtags)), top_hashtags.values, color=colors)
        plt.title("🔥 Top Trending Hashtags", fontsize=12, fontweight='bold')
        plt.xlabel("Frequency")
        plt.ylabel("Hashtag")
        plt.yticks(range(len(top_hashtags)), [f"#{tag}" for tag in top_hashtags.index])

        # Add value labels
        for i, (bar, value) in enumerate(zip(bars, top_hashtags.values)):
            plt.text(value + max(top_hashtags.values) * 0.01, i,
                     f'{value}', va='center', fontsize=8)

    # 3. Posts by day of week
    plt.subplot(3, 3, 3)
    if posts_by_day is not None and not posts_by_day.empty:
        day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        posts_by_day_ordered = posts_by_day.reindex(day_order, fill_value=0)
        colors = plt.cm.Set3(np.linspace(0, 1, 7))
        bars = plt.bar(range(7), posts_by_day_ordered.values, color=colors)
        plt.title("📅 Posts by Day of Week", fontsize=12, fontweight='bold')
        plt.xlabel("Day")
        plt.ylabel("Number of Posts")
        plt.xticks(range(7), [day[:3] for day in day_order], rotation=45)

        # Add value labels
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontsize=8)

    # 4. KOL Categories by average views
    plt.subplot(3, 2, 3)
    if top_categories is not None and not top_categories.empty:
        colors = plt.cm.coolwarm(np.linspace(0.2, 0.8, len(top_categories)))
        bars = plt.barh(range(len(top_categories)), top_categories.values, color=colors)
        plt.title("👑 KOL Categories by Avg Views", fontsize=12, fontweight='bold')
        plt.xlabel("Average Views")
        plt.ylabel("Category")
        plt.yticks(range(len(top_categories)), top_categories.index)

        # Format large numbers
        for i, (bar, value) in enumerate(zip(bars, top_categories.values)):
            if value >= 1_000_000:
                label = f'{value / 1_000_000:.1f}M'
            elif value >= 1_000:
                label = f'{value / 1_000:.1f}K'
            else:
                label = f'{int(value)}'
            plt.text(value + max(top_categories.values) * 0.01, i,
                     label, va='center', fontsize=8)

    # 5. Top fanpages by interaction rate
    plt.subplot(3, 2, 4)
    if top_fanpages is not None and len(top_fanpages) > 0:
        if 'Post interaction rate' in top_fanpages.columns and 'Profile' in top_fanpages.columns:
            colors = sns.color_palette("crest", n_colors=len(top_fanpages))
            bars = plt.barh(range(len(top_fanpages)),
                            top_fanpages['Post interaction rate'].fillna(0), color=colors)
            plt.title("📊 Top Fanpages by Interaction Rate", fontsize=12, fontweight='bold')
            plt.xlabel("Interaction Rate (%)")
            plt.ylabel("Fanpage")
            plt.yticks(range(len(top_fanpages)),
                       [name[:20] + "..." if len(str(name)) > 20 else str(name)
                        for name in top_fanpages['Profile']])

            # Add value labels
            for i, (bar, value) in enumerate(zip(bars, top_fanpages['Post interaction rate'].fillna(0))):
                plt.text(value + max(top_fanpages['Post interaction rate'].fillna(0)) * 0.01, i,
                         f'{value:.1f}%', va='center', fontsize=8)

    # 6. Engagement rate distribution (FastMoss)
    plt.subplot(3, 2, 5)
    if fastmoss is not None and 'engagement_rate' in fastmoss.columns:
        engagement_data = fastmoss['engagement_rate'].dropna()
        if not engagement_data.empty:
            plt.hist(engagement_data, bins=20, alpha=0.7, color='skyblue', edgecolor='black')
            plt.title("💫 Engagement Rate Distribution", fontsize=12, fontweight='bold')
            plt.xlabel("Engagement Rate (%)")
            plt.ylabel("Frequency")
            plt.axvline(engagement_data.mean(), color='red', linestyle='--',
                        label=f'Mean: {engagement_data.mean():.1f}%')
            plt.legend()

    # 7. Summary statistics
    plt.subplot(3, 2, 6)
    plt.axis('off')

    # Create summary text
    summary_text = "📈 CAMPAIGN SUMMARY\n\n"

    if apify is not None:
        summary_text += f"🎬 TikTok Posts: {len(apify):,}\n"
        if 'timestamp' in apify.columns:
            date_range = f"{apify['timestamp'].min().strftime('%Y-%m-%d')} to {apify['timestamp'].max().strftime('%Y-%m-%d')}"
            summary_text += f"📅 Date Range: {date_range}\n"

    if fastmoss is not None:
        summary_text += f"👥 Influencers: {len(fastmoss):,}\n"
        if 'Lượt xem' in fastmoss.columns:
            total_views = fastmoss['Lượt xem'].sum()
            if total_views >= 1_000_000:
                summary_text += f"👀 Total Views: {total_views / 1_000_000:.1f}M\n"
            else:
                summary_text += f"👀 Total Views: {total_views:,.0f}\n"

    if top_fanpages is not None and len(top_fanpages) > 0:
        summary_text += f"📱 Top Fanpages: {len(top_fanpages)}\n"
        if 'Post interaction rate' in top_fanpages.columns:
            avg_interaction = top_fanpages['Post interaction rate'].mean()
            summary_text += f"💬 Avg Interaction Rate: {avg_interaction:.1f}%\n"

    # Best performing content insights
    summary_text += "\n🏆 KEY INSIGHTS:\n"

    if post_by_hour is not None and not post_by_hour.empty:
        peak_hour = post_by_hour.idxmax()
        summary_text += f"⏰ Peak posting hour: {peak_hour}:00\n"

    if top_hashtags is not None and not top_hashtags.empty:
        top_hashtag = top_hashtags.index[0]
        summary_text += f"🔥 Top hashtag: #{top_hashtag}\n"

    if top_categories is not None and not top_categories.empty:
        best_category = top_categories.index[0]
        summary_text += f"👑 Best KOL category: {best_category}\n"

    plt.text(0.05, 0.95, summary_text, transform=plt.gca().transAxes,
             fontsize=10, verticalalignment='top', fontfamily='monospace',
             bbox=dict(boxstyle="round,pad=0.5", facecolor="lightgray", alpha=0.8))

    plt.tight_layout(pad=3.0)
    plt.suptitle("🌱 CỎ MỀM - Social Media Analytics Dashboard",
                 fontsize=16, fontweight='bold', y=0.98)

    return fig


def main():
    """Main execution function"""
    print("🌱 Starting Cỏ Mềm Social Media Analytics...")

    # Load data
    apify, fastmoss_df, fanpage = load_and_process_data()

    if apify is None and fastmoss_df is None and fanpage is None:
        print("❌ No data could be loaded. Please check your file paths.")
        return

    # Process each data source
    print("\n📊 Processing data...")

    apify_processed, top_hashtags, post_by_hour, posts_by_day = process_apify_data(apify)
    fastmoss_processed, top_categories = process_fastmoss_data(fastmoss_df)
    top_fanpages = process_fanpage_data(fanpage)

    # Create visualizations
    print("\n🎨 Creating visualizations...")

    fig = create_enhanced_visualizations(
        apify_processed, top_hashtags, post_by_hour, posts_by_day,
        fastmoss_processed, top_categories, top_fanpages
    )

    plt.show()

    print("\n✅ Analysis complete! Dashboard generated successfully.")

    # Optional: Save the figure
    # fig.savefig('co_mem_social_media_dashboard.png', dpi=300, bbox_inches='tight')
    # print("📁 Dashboard saved as 'co_mem_social_media_dashboard.png'")


if __name__ == "__main__":
    main()