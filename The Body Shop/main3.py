import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# Set Vietnamese locale for matplotlib
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS', 'Tahoma']
plt.style.use('seaborn-v0_8')


class BodyShopAnalytics:
    def __init__(self):
        self.karma = None
        self.apify = None
        self.fastmoss_video = None
        self.fastmoss_live = None
        self.fastmoss_product = None
        self.analysis_results = {}

    def safe_read_excel(self, file_path, sheet_name=None):
        """Safely read Excel files with error handling"""
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

    def list_excel_sheets(self, file_path):
        """List available sheets in Excel file"""
        try:
            xl_file = pd.ExcelFile(file_path)
            return xl_file.sheet_names
        except Exception as e:
            print(f"❌ Error reading {file_path}: {e}")
            return []

    def load_data(self):
        """Load all data sources"""
        print("📊 Loading data sources...")

        # Load Karma (TikTok) data
        self.karma = self.safe_read_excel('[FANPAGE KARMA] The Body Shop.xlsx',
                                          sheet_name='Metrics Overview')

        # Load Apify (TikTok) data
        self.apify = self.safe_read_excel('[APIFY] The Body Shop.xlsx')

        # Load FASTMOSS data
        fastmoss_sheets = self.list_excel_sheets('[FASTMOSS] The Body Shop.xlsx')
        print(f"🔍 Available FASTMOSS sheets: {fastmoss_sheets}")

        self.fastmoss_video = self.safe_read_excel('[FASTMOSS] The Body Shop.xlsx',
                                                   sheet_name='Data Video')
        self.fastmoss_live = self.safe_read_excel('[FASTMOSS] The Body Shop.xlsx',
                                                  sheet_name='Data Livestream')

        # Try to load product data
        possible_product_sheets = ['Data Product ', 'Data Product', 'Product',
                                   'Products', 'Sản phẩm', 'Data Sản phẩm']

        for sheet_name in possible_product_sheets:
            if sheet_name in fastmoss_sheets:
                self.fastmoss_product = self.safe_read_excel('[FASTMOSS] The Body Shop.xlsx',
                                                             sheet_name=sheet_name)
                print(f"✅ Found product data in sheet: {sheet_name}")
                break

        if self.fastmoss_product is None:
            print("⚠️ Warning: Product data sheet not found.")

    def safe_datetime_convert(self, df, column_name, unit=None):
        """Safely convert datetime columns"""
        if df is not None and column_name in df.columns:
            try:
                if unit:
                    df[column_name] = pd.to_datetime(df[column_name], unit=unit)
                else:
                    df[column_name] = pd.to_datetime(df[column_name])
                print(f"✅ Converted {column_name} to datetime")
                return True
            except Exception as e:
                print(f"⚠️ Warning: Could not convert {column_name} to datetime: {e}")
                return False
        return False

    def clean_currency(self, value):
        """Clean Vietnamese currency format"""
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

    def normalize_data(self):
        """Normalize and clean all data"""
        print("\n🔧 Normalizing data...")

        # Handle Karma (TikTok) data
        if self.karma is not None:
            print(f"Karma columns: {self.karma.columns.tolist()}")

            # Check if all columns are unnamed (header issue)
            if all(col.startswith('Unnamed') for col in self.karma.columns):
                print("⚠️ Trying to re-read Karma with header=1...")
                try:
                    self.karma = pd.read_excel('[FANPAGE KARMA] The Body Shop.xlsx',
                                               sheet_name='Metrics Overview', header=1)
                    print(f"✅ New Karma columns: {self.karma.columns.tolist()}")
                except Exception as e:
                    print(f"❌ Failed to re-read Karma: {e}")

            # Convert date columns
            date_columns = ['Date', 'date', 'Ngày', 'Time', 'Timestamp']
            for col in date_columns:
                if col in self.karma.columns:
                    self.safe_datetime_convert(self.karma, col)
                    break

        # Handle Apify (TikTok) data
        if self.apify is not None:
            self.safe_datetime_convert(self.apify, 'createTime', unit='s')

        # Handle FASTMOSS video data
        if self.fastmoss_video is not None:
            self.safe_datetime_convert(self.fastmoss_video, 'Thời gian phát hành')
            if 'Doanh số bán hàng của video' in self.fastmoss_video.columns:
                self.fastmoss_video['Doanh số (VND)'] = self.fastmoss_video['Doanh số bán hàng của video'].apply(
                    self.clean_currency)

        # Handle FASTMOSS livestream data
        if self.fastmoss_live is not None:
            self.safe_datetime_convert(self.fastmoss_live, 'Thời gian bắt đầu Livestream')
            if 'Doanh số Livestream' in self.fastmoss_live.columns:
                self.fastmoss_live['Doanh số (VND)'] = self.fastmoss_live['Doanh số Livestream'].apply(
                    self.clean_currency)

        # Handle product data
        if self.fastmoss_product is not None and 'Doanh số' in self.fastmoss_product.columns:
            self.fastmoss_product['Doanh số (VND)'] = self.fastmoss_product['Doanh số'].apply(self.clean_currency)

    def analyze_tiktok_engagement(self):
        """Analyze TikTok engagement patterns"""
        if self.karma is None:
            print("⚠️ Skipping TikTok analysis - data not available")
            return None

        print("\n📈 Analyzing TikTok engagement...")

        # Find date and engagement columns
        date_col = None
        engagement_col = None

        for col in self.karma.columns:
            if any(keyword in col.lower() for keyword in ['date', 'ngày', 'time']):
                date_col = col
                break

        for col in self.karma.columns:
            if any(keyword in col.lower() for keyword in ['engagement', 'tương tác', 'interact']):
                engagement_col = col
                break

        if date_col and engagement_col:
            # Create engagement over time plot
            plt.figure(figsize=(12, 6))
            sns.lineplot(data=self.karma, x=date_col, y=engagement_col)
            plt.title("TikTok Engagement Over Time", fontsize=14, fontweight='bold')
            plt.xlabel("Date")
            plt.ylabel("Engagement")
            plt.xticks(rotation=45)
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            plt.show()

            # Calculate basic statistics
            fb_stats = {
                'total_engagement': self.karma[engagement_col].sum(),
                'avg_daily_engagement': self.karma[engagement_col].mean(),
                'peak_engagement': self.karma[engagement_col].max(),
                'peak_date': self.karma.loc[self.karma[engagement_col].idxmax(), date_col]
            }

            self.analysis_results['tiktok'] = fb_stats
            return fb_stats
        else:
            print(f"⚠️ Required columns not found. Available: {self.karma.columns.tolist()}")
            return None

    def analyze_tiktok_performance(self):
        """Analyze TikTok video performance"""
        if self.apify is None:
            print("⚠️ Skipping TikTok analysis - data not available")
            return None

        print("\n🎵 Analyzing TikTok performance...")

        # Top performing videos
        if 'playCount' in self.apify.columns:
            top_videos = self.apify.nlargest(10, 'playCount')
            print("🔥 Top 10 TikTok videos by views:")

            display_cols = ['desc', 'playCount', 'diggCount', 'shareCount']
            available_cols = [col for col in display_cols if col in top_videos.columns]
            if available_cols:
                print(top_videos[available_cols].to_string(index=False))

            # Weekly performance analysis
            if 'createTime' in self.apify.columns:
                self.apify['week'] = self.apify['createTime'].dt.isocalendar().week
                self.apify['weekday'] = self.apify['createTime'].dt.day_name()
                self.apify['hour'] = self.apify['createTime'].dt.hour

                # Weekly aggregation
                weekly_stats = self.apify.groupby('week').agg({
                    'playCount': ['sum', 'mean', 'count'],
                    'diggCount': 'sum',
                    'shareCount': 'sum',
                    'commentCount': 'sum'
                }).round(2)

                # Calculate engagement rate
                total_interactions = (self.apify['diggCount'].fillna(0) +
                                      self.apify['shareCount'].fillna(0) +
                                      self.apify['commentCount'].fillna(0))
                self.apify['engagement_rate'] = (total_interactions / self.apify['playCount'] * 100).round(2)

                # Posting time analysis
                self.plot_posting_patterns()

                # Performance metrics
                tiktok_stats = {
                    'total_videos': len(self.apify),
                    'total_views': self.apify['playCount'].sum(),
                    'avg_views_per_video': self.apify['playCount'].mean(),
                    'avg_engagement_rate': self.apify['engagement_rate'].mean(),
                    'best_posting_hour': self.apify.groupby('hour')['playCount'].mean().idxmax()
                }

                self.analysis_results['tiktok'] = tiktok_stats
                return tiktok_stats

        return None

    def plot_posting_patterns(self):
        """Plot TikTok posting patterns"""
        if self.apify is None or 'hour' not in self.apify.columns:
            return

        fig, axes = plt.subplots(1, 2, figsize=(15, 5))

        # Hourly posting pattern
        hourly_views = self.apify.groupby('hour')['playCount'].mean()
        axes[0].bar(hourly_views.index, hourly_views.values, alpha=0.7, color='skyblue')
        axes[0].set_title('Average Views by Posting Hour')
        axes[0].set_xlabel('Hour of Day')
        axes[0].set_ylabel('Average Views')
        axes[0].grid(True, alpha=0.3)

        # Weekly posting pattern
        if 'weekday' in self.apify.columns:
            weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday',
                             'Friday', 'Saturday', 'Sunday']
            weekday_views = self.apify.groupby('weekday')['playCount'].mean().reindex(weekday_order)
            axes[1].bar(range(len(weekday_views)), weekday_views.values, alpha=0.7, color='lightcoral')
            axes[1].set_title('Average Views by Day of Week')
            axes[1].set_xlabel('Day of Week')
            axes[1].set_ylabel('Average Views')
            axes[1].set_xticks(range(len(weekday_order)))
            axes[1].set_xticklabels([day[:3] for day in weekday_order], rotation=45)
            axes[1].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.show()

    def compare_video_vs_livestream(self):
        """Compare video vs livestream performance"""
        if self.fastmoss_video is None or self.fastmoss_live is None:
            print("⚠️ Skipping video vs livestream comparison - insufficient data")
            return None

        print("\n📊 Comparing Video vs Livestream performance...")

        # Find date columns
        video_date_col = self.find_date_column(self.fastmoss_video,
                                               ['Thời gian phát hành', 'Ngày đăng', 'Date'])
        live_date_col = self.find_date_column(self.fastmoss_live,
                                              ['Thời gian bắt đầu Livestream', 'Ngày', 'Date'])

        if not video_date_col or not live_date_col:
            print("⚠️ Date columns not found for comparison")
            return None

        # Aggregate by date
        video_metrics = self.aggregate_by_date(self.fastmoss_video, video_date_col)
        live_metrics = self.aggregate_by_date(self.fastmoss_live, live_date_col)

        if video_metrics is not None and live_metrics is not None:
            self.plot_video_vs_live_comparison(video_metrics, live_metrics)

            # Calculate comparison stats
            comparison_stats = {
                'video_total_views': video_metrics['Lượt xem'].sum() if 'Lượt xem' in video_metrics.columns else 0,
                'live_total_views': live_metrics['Lượt xem'].sum() if 'Lượt xem' in live_metrics.columns else 0,
                'video_avg_revenue': video_metrics[
                    'Doanh số (VND)'].mean() if 'Doanh số (VND)' in video_metrics.columns else 0,
                'live_avg_revenue': live_metrics[
                    'Doanh số (VND)'].mean() if 'Doanh số (VND)' in live_metrics.columns else 0
            }

            self.analysis_results['comparison'] = comparison_stats
            return comparison_stats

        return None

    def find_date_column(self, df, possible_cols):
        """Find the first available date column"""
        for col in possible_cols:
            if col in df.columns:
                return col
        return None

    def aggregate_by_date(self, df, date_col):
        """Aggregate metrics by date"""
        if date_col not in df.columns:
            return None

        agg_dict = {}
        if 'Lượt xem' in df.columns:
            agg_dict['Lượt xem'] = 'sum'
        if 'Doanh số (VND)' in df.columns:
            agg_dict['Doanh số (VND)'] = 'sum'
        if 'Số lượng likes' in df.columns:
            agg_dict['Số lượng likes'] = 'sum'

        if agg_dict:
            return df.groupby(date_col).agg(agg_dict).reset_index()
        return None

    def plot_video_vs_live_comparison(self, video_data, live_data):
        """Plot comparison between video and livestream"""
        fig, axes = plt.subplots(2, 2, figsize=(15, 10))

        # Views comparison
        if 'Lượt xem' in video_data.columns and 'Lượt xem' in live_data.columns:
            axes[0, 0].plot(video_data.index, video_data['Lượt xem'],
                            label='Video', marker='o', alpha=0.7)
            axes[0, 0].plot(live_data.index, live_data['Lượt xem'],
                            label='Livestream', marker='s', alpha=0.7)
            axes[0, 0].set_title('Views: Video vs Livestream')
            axes[0, 0].set_ylabel('Views')
            axes[0, 0].legend()
            axes[0, 0].grid(True, alpha=0.3)

        # Revenue comparison
        if 'Doanh số (VND)' in video_data.columns and 'Doanh số (VND)' in live_data.columns:
            axes[0, 1].plot(video_data.index, video_data['Doanh số (VND)'],
                            label='Video Revenue', marker='o', alpha=0.7)
            axes[0, 1].plot(live_data.index, live_data['Doanh số (VND)'],
                            label='Livestream Revenue', marker='s', alpha=0.7)
            axes[0, 1].set_title('Revenue: Video vs Livestream')
            axes[0, 1].set_ylabel('Revenue (VND)')
            axes[0, 1].legend()
            axes[0, 1].grid(True, alpha=0.3)

        # Total comparison bars
        if 'Lượt xem' in video_data.columns and 'Lượt xem' in live_data.columns:
            categories = ['Total Views', 'Total Revenue']
            video_totals = [video_data['Lượt xem'].sum(),
                            video_data['Doanh số (VND)'].sum() if 'Doanh số (VND)' in video_data.columns else 0]
            live_totals = [live_data['Lượt xem'].sum(),
                           live_data['Doanh số (VND)'].sum() if 'Doanh số (VND)' in live_data.columns else 0]

            x = np.arange(len(categories))
            width = 0.35

            axes[1, 0].bar(x - width / 2, video_totals, width, label='Video', alpha=0.7)
            axes[1, 0].bar(x + width / 2, live_totals, width, label='Livestream', alpha=0.7)
            axes[1, 0].set_title('Total Performance Comparison')
            axes[1, 0].set_xticks(x)
            axes[1, 0].set_xticklabels(categories)
            axes[1, 0].legend()
            axes[1, 0].grid(True, alpha=0.3)

        # Efficiency metrics
        if ('Doanh số (VND)' in video_data.columns and 'Lượt xem' in video_data.columns and
                'Doanh số (VND)' in live_data.columns and 'Lượt xem' in live_data.columns):
            video_efficiency = video_data['Doanh số (VND)'] / video_data['Lượt xem']
            live_efficiency = live_data['Doanh số (VND)'] / live_data['Lượt xem']

            axes[1, 1].hist([video_efficiency.dropna(), live_efficiency.dropna()],
                            bins=20, alpha=0.7, label=['Video', 'Livestream'])
            axes[1, 1].set_title('Revenue Efficiency (VND per View)')
            axes[1, 1].set_xlabel('VND per View')
            axes[1, 1].set_ylabel('Frequency')
            axes[1, 1].legend()
            axes[1, 1].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.show()

    def generate_insights_report(self):
        """Generate comprehensive insights report"""
        print("\n📋 Generating insights report...")

        report = {
            'analysis_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'data_sources': {
                '': self.karma is not None,
                'tiktok': self.apify is not None,
                'video_content': self.fastmoss_video is not None,
                'livestream': self.fastmoss_live is not None,
                'product_data': self.fastmoss_product is not None
            },
            'key_insights': [],
            'recommendations': []
        }

        # Add insights based on analysis results
        if 'tiktok' in self.analysis_results:
            tiktok_stats = self.analysis_results['tiktok']
            report['key_insights'].append(
                f"TikTok Performance: {tiktok_stats['total_videos']} videos with "
                f"{tiktok_stats['total_views']:,.0f} total views and "
                f"{tiktok_stats['avg_engagement_rate']:.2f}% average engagement rate"
            )

            if tiktok_stats['best_posting_hour']:
                report['recommendations'].append(
                    f"Optimal posting time: {tiktok_stats['best_posting_hour']}:00 "
                    f"based on highest average views"
                )

        if 'tiktok' in self.analysis_results:
            fb_stats = self.analysis_results['tiktok']
            report['key_insights'].append(
                f"Tiktok Engagement: Peak engagement of {fb_stats['peak_engagement']:,.0f} "
                f"on {fb_stats['peak_date']}"
            )

        if 'comparison' in self.analysis_results:
            comp_stats = self.analysis_results['comparison']
            if comp_stats['video_total_views'] > comp_stats['live_total_views']:
                report['recommendations'].append(
                    "Regular videos show higher total views than livestreams - "
                    "consider increasing video content frequency"
                )
            else:
                report['recommendations'].append(
                    "Livestreams show strong performance - "
                    "consider increasing livestream frequency"
                )

        return report

    def export_results(self):
        """Export analysis results to Excel"""
        print("\n💾 Exporting results to Excel...")

        sheets_to_export = {}

        # Prepare data for export
        if self.karma is not None:
            sheets_to_export["TikTok_Data"] = self.karma

        if self.apify is not None:
            # Add calculated metrics to TikTok data
            apify_export = self.apify.copy()
            if 'engagement_rate' not in apify_export.columns and 'playCount' in apify_export.columns:
                total_interactions = (apify_export['diggCount'].fillna(0) +
                                      apify_export['shareCount'].fillna(0) +
                                      apify_export['commentCount'].fillna(0))
                apify_export['engagement_rate'] = (total_interactions / apify_export['playCount'] * 100).round(2)

            sheets_to_export["TikTok_Data"] = apify_export

        if self.fastmoss_video is not None:
            sheets_to_export["Video_Data"] = self.fastmoss_video

        if self.fastmoss_live is not None:
            sheets_to_export["Livestream_Data"] = self.fastmoss_live

        if self.fastmoss_product is not None:
            sheets_to_export["Product_Data"] = self.fastmoss_product

        # Add summary sheet
        if self.analysis_results:
            summary_data = []
            for platform, stats in self.analysis_results.items():
                for metric, value in stats.items():
                    summary_data.append({
                        'Platform': platform.title(),
                        'Metric': metric.replace('_', ' ').title(),
                        'Value': value
                    })

            if summary_data:
                sheets_to_export["Analysis_Summary"] = pd.DataFrame(summary_data)

        # Export to Excel with formatting
        if sheets_to_export:
            with pd.ExcelWriter("TheBodyShop_Enhanced_Analytics.xlsx",
                                engine='xlsxwriter') as writer:
                for sheet_name, df in sheets_to_export.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Get workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    # Add formatting
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'fg_color': '#D7E4BC',
                        'border': 1
                    })

                    # Apply header formatting
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, len(str(value)) + 5)

            print("✅ Enhanced analytics exported: TheBodyShop_Enhanced_Analytics.xlsx")
            print(f"📋 Sheets exported: {list(sheets_to_export.keys())}")
        else:
            print("❌ No data available for export")

    def run_complete_analysis(self):
        """Run the complete analysis pipeline"""
        print("🚀 Starting The Body Shop Social Media Analytics")
        print("=" * 50)

        # Load and normalize data
        self.load_data()
        self.normalize_data()

        # Run analyses
        self.analyze_tiktok_engagement()
        self.analyze_tiktok_performance()
        self.compare_video_vs_livestream()

        # Generate report and export
        report = self.generate_insights_report()
        self.export_results()

        # Print final summary
        print("\n🎉 Analysis Complete!")
        print("=" * 50)
        print("📊 Data Sources Processed:")
        for source, available in report['data_sources'].items():
            status = "✅" if available else "❌"
            print(f"  {status} {source.replace('_', ' ').title()}")

        print("\n💡 Key Insights:")
        for insight in report['key_insights']:
            print(f"  • {insight}")

        print("\n🎯 Recommendations:")
        for rec in report['recommendations']:
            print(f"  • {rec}")

        return report


# Execute the analysis
if __name__ == "__main__":
    analyzer = BodyShopAnalytics()
    final_report = analyzer.run_complete_analysis()