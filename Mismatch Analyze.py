import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
from pathlib import Path
import io
import base64
import warnings
from PIL import Image
import plotly.io as pio
warnings.filterwarnings('ignore')

# تبدیل تاریخ
try:
    import jdatetime
    JALALI_AVAILABLE = True
except ImportError:
    JALALI_AVAILABLE = False

# تنظیمات صفحه
st.set_page_config(
    page_title="سامانه تحلیل مغایرت‌ها",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# استایل CSS کامل
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700&display=swap');
    
    * {
        font-family: 'Vazirmatn', 'Tahoma', sans-serif !important;
    }
    
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 10px;
        font-weight: 600;
        font-size: 14px;
        transition: all 0.3s;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
    }
    
    .stButton>button:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.5);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 30px;
        border-radius: 20px;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        transition: all 0.3s;
        border: 2px solid rgba(255,255,255,0.1);
    }
    
    .metric-card:hover {
        transform: translateY(-8px) scale(1.02);
        box-shadow: 0 15px 40px rgba(0,0,0,0.3);
    }
    
    .info-box {
        background: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        border-left: 5px solid #667eea;
        margin: 10px 0;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: transparent;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: white;
        border-radius: 10px 10px 0 0;
        padding: 12px 24px;
        font-weight: 600;
        font-size: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    .download-section {
        background: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin: 20px 0;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-weight: 700;
    }
    
    .stExpander {
        background: white;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    .highlight-box {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin: 10px 0;
    }
    
    .warning-box {
        background: linear-gradient(135deg, #f39c12 0%, #e74c3c 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin: 10px 0;
    }
    
    .success-box {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)


def gregorian_to_jalali(g_date):
    """تبدیل تاریخ میلادی به شمسی"""
    if not JALALI_AVAILABLE:
        return str(g_date)
    
    try:
        if isinstance(g_date, str):
            g_date = datetime.strptime(g_date, '%Y-%m-%d')
        j_date = jdatetime.date.fromgregorian(date=g_date)
        return j_date.strftime('%Y/%m/%d')
    except:
        return str(g_date)


@st.cache_data
def load_excel_files(uploaded_files):
    """خواندن و ترکیب فایل‌های اکسل"""
    all_data = []
    file_info = []
    
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)
            
            file_name = uploaded_file.name
            parts = Path(file_name).stem.split('_')
            date_part = parts[-1] if len(parts) > 0 else ''
            
            try:
                
                if '-' in date_part:
                    date_obj = datetime.strptime(date_part, '%Y-%m-%d')
                else:
                    date_obj = datetime.strptime(date_part, '%Y%m%d')
                gregorian_date = date_obj.strftime('%Y-%m-%d')
                jalali_date = gregorian_to_jalali(date_obj)
                
                df['تاریخ میلادی'] = gregorian_date
                df['تاریخ شمسی'] = jalali_date
                df['تاریخ_obj'] = date_obj
            except ValueError:
                st.warning(f"⚠️ تاریخ از نام فایل '{file_name}' استخراج نشد. لطفاً فرمت YYYYMMDD را در انتهای نام فایل بررسی کنید.")
                now = datetime.now()
                df['تاریخ میلادی'] = now.strftime('%Y-%m-%d')
                df['تاریخ شمسی'] = gregorian_to_jalali(now)
                df['تاریخ_obj'] = now
            
            df['نام فایل'] = file_name
            all_data.append(df)
            
            file_info.append({
                'نام فایل': file_name,
                'تعداد ردیف': len(df),
                'تعداد ستون': len(df.columns),
                'تاریخ میلادی': df['تاریخ میلادی'].iloc[0] if len(df) > 0 else 'نامشخص',
                'تاریخ شمسی': df['تاریخ شمسی'].iloc[0] if len(df) > 0 else 'نامشخص'
            })
            
        except Exception as e:
            st.error(f"خطا در خواندن {uploaded_file.name}: {str(e)}")
            continue
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df = combined_df.sort_values('تاریخ_obj')
        return combined_df, pd.DataFrame(file_info)
    
    return None, None


def detect_columns(df):
    """شناسایی خودکار ستون‌های مهم"""
    columns = {
        'province': None,
        'site': None,
        'issue': None,
        'comment': None
    }
    
    for col in df.columns:
        col_str = str(col).lower()
        if 'استان' in col_str or 'province' in col_str:
            columns['province'] = col
        elif ('سایت' in col_str or 'site' in col_str or 'کد' in col_str) and not columns['site']:
            columns['site'] = col
        elif 'ستون' in col_str and 'مغایرت' in col_str:
            columns['issue'] = col
        elif ('کامنت' in col_str or 'عنوان' in col_str or 'توضیح' in col_str) and not columns['comment']:
            columns['comment'] = col
    
    return columns


def create_unique_key(df, cols):
    """ایجاد کلید منحصر به فرد برای شناسایی مغایرت‌های تکراری"""
    if not cols['site'] or cols['site'] not in df.columns:
        return None
    
    key_parts = [df[cols['site']].astype(str)]
    
    if cols['issue'] and cols['issue'] in df.columns:
        key_parts.append(df[cols['issue']].astype(str))
    
    if cols['comment'] and cols['comment'] in df.columns:
        key_parts.append(df[cols['comment']].astype(str))
    
    return pd.Series('||'.join(part) for part in zip(*key_parts))


def calculate_summary_stats(df, cols):
    """محاسبه آمار خلاصه"""
    stats = {
        'total_issues': len(df),
        'unique_sites': df[cols['site']].nunique() if cols['site'] and cols['site'] in df.columns else 0,
        'unique_provinces': df[cols['province']].nunique() if cols['province'] and cols['province'] in df.columns else 0,
        'total_dates': df['تاریخ شمسی'].nunique() if 'تاریخ شمسی' in df.columns else 0,
        'date_range': f"{df['تاریخ شمسی'].min()} تا {df['تاریخ شمسی'].max()}" if 'تاریخ شمسی' in df.columns and df['تاریخ شمسی'].nunique() > 0 else 'نامشخص',
        'files_count': df['نام فایل'].nunique() if 'نام فایل' in df.columns else 0
    }
    return stats


def calculate_progress(df, cols):
    """محاسبه پیشرفت رفع مغایرت‌ها"""
    if not cols['province'] or cols['province'] not in df.columns:
        return pd.DataFrame()
    
    if 'تاریخ شمسی' not in df.columns:
        return pd.DataFrame()
    
    dates = sorted(df['تاریخ شمسی'].unique())
    if len(dates) < 2:
        return pd.DataFrame()
    
    first_date = dates[0]
    last_date = dates[-1]
    
    df_copy = df.copy()
    df_copy['کلید_منحصر'] = create_unique_key(df_copy, cols)
    
    if 'کلید_منحصر' not in df_copy.columns or df_copy['کلید_منحصر'].isnull().all():
        st.warning("کلید منحصر به فرد برای مقایسه مغایرت‌ها ایجاد نشد. لطفاً ستون‌های مورد نیاز را بررسی کنید.")
        return pd.DataFrame()

    progress_data = []
    
    all_provinces = sorted(df[cols['province']].dropna().unique())
    
    for province in all_provinces:
        province_df = df_copy[df_copy[cols['province']] == province]
        
        first_data = province_df[province_df['تاریخ شمسی'] == first_date]
        last_data = province_df[province_df['تاریخ شمسی'] == last_date]
        
        first_count = len(first_data)
        last_count = len(last_data)
        
        first_issues = set(first_data['کلید_منحصر'])
        last_issues = set(last_data['کلید_منحصر'])
        
        resolved = len(first_issues - last_issues)
        new_issues = len(last_issues - first_issues)
        remaining = len(first_issues & last_issues)
        
        progress_pct = (resolved / first_count * 100) if first_count > 0 else 0
        
        progress_data.append({
            'استان': province,
            'مغایرت اولیه': first_count,
            'مغایرت فعلی': last_count,
            'رفع شده': resolved,
            'باقیمانده': remaining,
            'مغایرت جدید': new_issues,
            'درصد پیشرفت': round(progress_pct, 2),
            'تاریخ اول': first_date,
            'تاریخ آخر': last_date,
            'وضعیت': '🟢 عالی' if progress_pct >= 75 else ('🟡 خوب' if progress_pct >= 50 else ('🟠 متوسط' if progress_pct >= 25 else '🔴 ضعیف'))
        })
    
    result_df = pd.DataFrame(progress_data)
    if not result_df.empty:
        result_df = result_df.sort_values('درصد پیشرفت', ascending=False)
    
    return result_df


def find_repeated_issues(df, cols):
    """شناسایی مغایرت‌های تکراری بر اساس (کد سایت + نوع مغایرت + عنوان مغایرت)"""
    if not cols['site'] or cols['site'] not in df.columns:
        return pd.DataFrame()

    key_parts = []
    key_parts.append(df[cols['site']].astype(str))
    if cols['issue'] and cols['issue'] in df.columns:
        key_parts.append(df[cols['issue']].astype(str))
    else:
        key_parts.append(pd.Series('', index=df.index))
    if cols['comment'] and cols['comment'] in df.columns:
        key_parts.append(df[cols['comment']].astype(str))
    else:
        key_parts.append(pd.Series('', index=df.index))

    df_copy = df.copy()
    df_copy['کلید_منحصر'] = pd.Series(['||'.join(parts) for parts in zip(*key_parts)], index=df.index)

    df_copy = df_copy.drop_duplicates(subset=['کلید_منحصر', 'تاریخ شمسی'])

    # پیدا کردن آخرین تاریخ
    if 'تاریخ شمسی' not in df_copy.columns or df_copy['تاریخ شمسی'].empty:
        last_date = None
    else:
        last_date = df_copy['تاریخ شمسی'].max()

    counts = df_copy.groupby('کلید_منحصر')['تاریخ شمسی'].nunique().reset_index(name='تعداد تکرار')

    extra_info = df_copy.groupby('کلید_منحصر').agg({
        'تاریخ شمسی': ['min', 'max'],
        cols['site']: 'first',
        cols['issue']: 'first',
        cols['comment']: 'first'
    }).reset_index()

    extra_info.columns = ['کلید_منحصر', 'اولین مشاهده', 'آخرین مشاهده', 'کد سایت', 'نوع مغایرت', 'عنوان مغایرت']

    result = pd.merge(counts, extra_info, on='کلید_منحصر')

    if cols['province'] and cols['province'] in df.columns:
        province_map = df.drop_duplicates(subset=[cols['site']]).set_index(cols['site'])[cols['province']]
        result['استان'] = result['کد سایت'].map(province_map)

    repeated = result[result['تعداد تکرار'] > 1].copy()
    if repeated.empty:
        return pd.DataFrame()

    # بررسی وضعیت برطرف شدن
    if last_date is not None:
        last_report_issues = set(df_copy[df_copy['تاریخ شمسی'] == last_date]['کلید_منحصر'].unique())
        
        repeated['وضعیت رفع'] = repeated['کلید_منحصر'].apply(
            lambda x: '❌ برطرف نشده' if x in last_report_issues else '✅ برطرف شده'
        )
        
        repeated['نماد'] = repeated['کلید_منحصر'].apply(
            lambda x: '🔴' if x in last_report_issues else '🟢'
        )
    else:
        repeated['وضعیت رفع'] = 'نامشخص'
        repeated['نماد'] = '⚪'

    repeated['اولویت'] = repeated['تعداد تکرار'].apply(
        lambda x: '🔴 بحرانی' if x >= 5 else ('🟠 مهم' if x >= 3 else '🟡 عادی')
    )

    repeated['مدت تکرار'] = repeated.apply(
        lambda row: f"{row['اولین مشاهده']} تا {row['آخرین مشاهده']}", axis=1
    )

    col_order = ['نماد', 'استان', 'کد سایت', 'نوع مغایرت', 'عنوان مغایرت',
                 'تعداد تکرار', 'اولویت', 'وضعیت رفع', 'اولین مشاهده', 'آخرین مشاهده', 'مدت تکرار']
    
    repeated = repeated[[col for col in col_order if col in repeated.columns]]
    
    # مرتب‌سازی: اول برطرف نشده‌ها
    repeated['sort_key'] = repeated['وضعیت رفع'].apply(lambda x: 0 if 'نشده' in x else 1)
    repeated = repeated.sort_values(['sort_key', 'تعداد تکرار'], ascending=[True, False])
    repeated = repeated.drop('sort_key', axis=1)
    
    return repeated


def find_new_issues(df, cols):
    """شناسایی مغایرت‌های جدید که فقط در آخرین گزارش هستند"""
    if 'تاریخ شمسی' not in df.columns or df['تاریخ شمسی'].nunique() < 2:
        return pd.DataFrame()
    
    dates = sorted(df['تاریخ شمسی'].unique())
    last_date = dates[-1]
    previous_date = dates[-2]
    
    df_copy = df.copy()
    df_copy['کلید_منحصر'] = create_unique_key(df_copy, cols)
    
    if 'کلید_منحصر' not in df_copy.columns:
        return pd.DataFrame()
    
    last_issues = set(df_copy[df_copy['تاریخ شمسی'] == last_date]['کلید_منحصر'])
    previous_issues = set(df_copy[df_copy['تاریخ شمسی'] == previous_date]['کلید_منحصر'])
    
    new_issue_keys = last_issues - previous_issues
    
    if not new_issue_keys:
        return pd.DataFrame()
    
    new_issues_df = df_copy[(df_copy['کلید_منحصر'].isin(new_issue_keys)) & 
                             (df_copy['تاریخ شمسی'] == last_date)].copy()
    
    if cols['province'] in new_issues_df.columns:
        result = new_issues_df[[cols['province'], cols['site'], cols['issue'], cols['comment']]].copy()
        result.columns = ['استان', 'کد سایت', 'نوع مغایرت', 'عنوان مغایرت']
    else:
        result = new_issues_df[[cols['site'], cols['issue'], cols['comment']]].copy()
        result.columns = ['کد سایت', 'نوع مغایرت', 'عنوان مغایرت']
    
    result['تاریخ ظهور'] = last_date
    result['اولویت بررسی'] = '🔴 فوری'
    
    return result


def analyze_issue_types(df, cols):
    """تحلیل توزیع انواع مغایرت (Pareto Analysis)"""
    if not cols['issue'] or cols['issue'] not in df.columns:
        return pd.DataFrame()
    
    issue_counts = df[cols['issue']].value_counts().reset_index()
    issue_counts.columns = ['نوع مغایرت', 'تعداد']
    
    total = issue_counts['تعداد'].sum()
    issue_counts['درصد'] = (issue_counts['تعداد'] / total * 100).round(2)
    issue_counts['درصد تجمعی'] = issue_counts['درصد'].cumsum().round(2)
    
    issue_counts['دسته‌بندی'] = issue_counts['درصد تجمعی'].apply(
        lambda x: '🔴 بحرانی (80%)' if x <= 80 else '🟡 مهم (95%)' if x <= 95 else '🟢 کم‌اهمیت'
    )
    
    return issue_counts


def calculate_benchmark(progress_df):
    """محاسبه Benchmark و مقایسه با میانگین کشوری"""
    if progress_df.empty:
        return pd.DataFrame()
    
    national_avg = progress_df['درصد پیشرفت'].mean()
    national_median = progress_df['درصد پیشرفت'].median()
    
    benchmark_df = progress_df.copy()
    benchmark_df['میانگین کشوری'] = national_avg
    benchmark_df['میانه کشوری'] = national_median
    benchmark_df['انحراف از میانگین'] = (benchmark_df['درصد پیشرفت'] - national_avg).round(2)
    benchmark_df['عملکرد نسبی'] = benchmark_df['انحراف از میانگین'].apply(
        lambda x: '⭐ بالاتر از میانگین' if x > 10 else ('✅ نزدیک به میانگین' if x >= -10 else '⚠️ پایین‌تر از میانگین')
    )
    
    return benchmark_df


def compare_two_provinces(df, cols, province1, province2):
    """مقایسه دقیق دو استان"""
    if not cols['province'] or cols['province'] not in df.columns:
        return None, None
    
    p1_data = df[df[cols['province']] == province1]
    p2_data = df[df[cols['province']] == province2]
    
    comparison = {
        'معیار': [
            'مجموع مغایرت‌ها',
            'تعداد سایت‌های درگیر',
            'میانگین مغایرت به ازای هر سایت',
            'تعداد انواع مغایرت'
        ],
        province1: [
            len(p1_data),
            p1_data[cols['site']].nunique() if cols['site'] in df.columns else 0,
            len(p1_data) / max(p1_data[cols['site']].nunique(), 1) if cols['site'] in df.columns else 0,
            p1_data[cols['issue']].nunique() if cols['issue'] in df.columns else 0
        ],
        province2: [
            len(p2_data),
            p2_data[cols['site']].nunique() if cols['site'] in df.columns else 0,
            len(p2_data) / max(p2_data[cols['site']].nunique(), 1) if cols['site'] in df.columns else 0,
            p2_data[cols['issue']].nunique() if cols['issue'] in df.columns else 0
        ]
    }
    
    comparison_df = pd.DataFrame(comparison)
    
    # پیدا کردن مغایرت‌های مشترک
    if cols['issue'] in df.columns:
        common_issues = set(p1_data[cols['issue']].unique()) & set(p2_data[cols['issue']].unique())
        common_df = pd.DataFrame({
            'مغایرت مشترک': list(common_issues),
            f'تعداد در {province1}': [len(p1_data[p1_data[cols['issue']] == issue]) for issue in common_issues],
            f'تعداد در {province2}': [len(p2_data[p2_data[cols['issue']] == issue]) for issue in common_issues]
        })
    else:
        common_df = pd.DataFrame()
    
    return comparison_df, common_df


def compare_reports(df, cols):
    """مقایسه کامل گزارش‌ها با یکدیگر"""
    if 'تاریخ شمسی' not in df.columns:
        return None
    
    dates = sorted(df['تاریخ شمسی'].unique())
    comparison_data = []
    
    for date in dates:
        date_df = df[df['تاریخ شمسی'] == date]
        
        comparison_data.append({
            'تاریخ': date,
            'تعداد مغایرت': len(date_df),
            'تعداد سایت': date_df[cols['site']].nunique() if cols['site'] and cols['site'] in df.columns else 0,
            'تعداد استان': date_df[cols['province']].nunique() if cols['province'] and cols['province'] in df.columns else 0,
            'نام فایل': date_df['نام فایل'].iloc[0] if len(date_df) > 0 and 'نام فایل' in date_df.columns else 'نامشخص'
        })
    
    result_df = pd.DataFrame(comparison_data)
    
    if len(result_df) > 1:
        result_df['تغییر از قبل'] = result_df['تعداد مغایرت'].diff().fillna(0).astype(int)
        result_df['درصد تغییر'] = (result_df['تعداد مغایرت'].pct_change() * 100).round(2)
        result_df['روند'] = result_df['تغییر از قبل'].apply(
            lambda x: '⬇️ کاهش' if x < 0 else ('⬆️ افزایش' if x > 0 else '➡️ بدون تغییر')
        )
        result_df.loc[0, 'روند'] = '-'
    
    return result_df


def calculate_province_timeline(df, cols, province):
    """محاسبه روند زمانی برای یک استان خاص"""
    if not cols['province'] or cols['province'] not in df.columns or 'تاریخ شمسی' not in df.columns:
        return None
    
    all_dates = sorted(df['تاریخ شمسی'].unique())
    
    province_df = df[df[cols['province']] == province]
    
    timeline_counts = province_df.groupby('تاریخ شمسی').size().reset_index(name='تعداد مغایرت')
    
    master_timeline = pd.DataFrame({'تاریخ شمسی': all_dates})
    
    full_timeline = pd.merge(master_timeline, timeline_counts, on='تاریخ شمسی', how='left')
    
    full_timeline['تعداد مغایرت'] = full_timeline['تعداد مغایرت'].fillna(0).astype(int)
    
    return full_timeline


def create_trend_chart(df):
    """نمودار روند کلی مغایرت‌ها در کل کشور"""
    if 'تاریخ شمسی' not in df.columns:
        return None
    
    daily_counts = df.groupby('تاریخ شمسی').size().reset_index(name='تعداد')
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=daily_counts['تاریخ شمسی'],
        y=daily_counts['تعداد'],
        mode='lines+markers+text',
        name='تعداد مغایرت',
        line=dict(color='#667eea', width=4),
        marker=dict(size=12, color='#764ba2', line=dict(color='white', width=3)),
        fill='tozeroy',
        fillcolor='rgba(102, 126, 234, 0.15)',
        text=daily_counts['تعداد'],
        textposition="top center",
        textfont=dict(size=12, color='#764ba2'),
        hovertemplate='<b>تاریخ:</b> %{x}<br><b>تعداد کل:</b> %{y:,}<extra></extra>'
    ))
    
    # اضافه کردن خط روند (Trend Line)
    if len(daily_counts) > 2:
        z = np.polyfit(range(len(daily_counts)), daily_counts['تعداد'], 1)
        p = np.poly1d(z)
        
        fig.add_trace(go.Scatter(
            x=daily_counts['تاریخ شمسی'],
            y=p(range(len(daily_counts))),
            mode='lines',
            name='خط روند',
            line=dict(color='red', width=2, dash='dash'),
            hovertemplate='<b>روند:</b> %{y:.0f}<extra></extra>'
        ))
    
    fig.update_layout(
        title={
            'text': '📈 روند تغییرات مغایرت‌ها در کل کشور',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50'}
        },
        xaxis_title='تاریخ شمسی',
        yaxis_title='تعداد مغایرت',
        hovermode='x unified',
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=12),
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    
    return fig


def create_province_chart(df, cols):
    """نمودار توزیع استان‌ها"""
    if not cols['province'] or cols['province'] not in df.columns:
        return None
    
    province_counts = df.groupby(cols['province']).size().reset_index(name='تعداد')
    province_counts = province_counts.sort_values('تعداد', ascending=True).tail(20)
    

    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=province_counts['تعداد'],
        y=province_counts[cols['province']],
        orientation='h',
        marker=dict(
            color=province_counts['تعداد'],
            colorscale='Plasma',
            showscale=True,
            colorbar=dict(title='تعداد', thickness=15)
        ),
        text=province_counts['تعداد'].apply(lambda x: f'{x:,}'),
        textposition='outside',
        textfont=dict(size=13, weight='bold'),
        hovertemplate='<b>%{y}</b><br>تعداد: %{x:,}<extra></extra>'
    ))
    
    fig.update_layout(
        title={
            'text': '🗺️ استان‌های با مغایرت بیشتر (20 استان برتر)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50'}
        },
        xaxis_title='تعداد مغایرت',
        yaxis_title='',
        template='plotly_white',
        height=700,
        font=dict(family='Vazirmatn, Tahoma', size=12),
        showlegend=False,
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    
    return fig


def create_province_progress_chart(df, cols, province):
    """نمودار خطی پیشرفت برای یک استان"""
    timeline = calculate_province_timeline(df, cols, province)
    
    if timeline is None or timeline.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=timeline['تاریخ شمسی'],
        y=timeline['تعداد مغایرت'],
        mode='lines+markers+text',
        name=province,
        line=dict(color='#e74c3c', width=3),
        marker=dict(size=10, color='#c0392b', line=dict(color='white', width=2)),
        fill='tozeroy',
        fillcolor='rgba(231, 76, 60, 0.1)',
        text=timeline['تعداد مغایرت'],
        textposition="top center",
        hovertemplate='<b>تاریخ:</b> %{x}<br><b>تعداد:</b> %{y:,}<extra></extra>'
    ))
    
    fig.update_layout(
        title={
            'text': f'روند مغایرت‌ها در استان {province}',
            'x': 0.5,
            'xanchor': 'center',
            
            'font': {'size': 16, 'color': '#2c3e50'}
        },
        xaxis_title='تاریخ',
        yaxis_title='تعداد',
        template='plotly_white',
        height=400,
        font=dict(family='Vazirmatn, Tahoma', size=11),
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    
    return fig


def create_comparison_chart(comparison_df):
    """نمودار مقایسه گزارش‌ها"""
    if comparison_df is None or comparison_df.empty:
        return None
    
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=('تعداد مغایرت‌ها در هر گزارش', 'تغییرات نسبت به گزارش قبل'),
        vertical_spacing=0.15,
        row_heights=[0.6, 0.4]
    )
    
    fig.add_trace(
        go.Bar(
            x=comparison_df['تاریخ'],
            y=comparison_df['تعداد مغایرت'],
            name='تعداد مغایرت',
            marker_color='#3498db',
            text=comparison_df['تعداد مغایرت'].apply(lambda x: f'{x:,}'),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>تعداد: %{y:,}<extra></extra>'
        ),
        row=1, col=1
    )
    
    if 'تغییر از قبل' in comparison_df.columns:
        colors = ['#2ecc71' if x < 0 else ('#e74c3c' if x > 0 else '#95a5a6') 
                  for x in comparison_df['تغییر از قبل']]
        
        fig.add_trace(
            go.Bar(
                x=comparison_df['تاریخ'],
                y=comparison_df['تغییر از قبل'],
                name='تغییر',
                marker_color=colors,
                text=comparison_df['روند'],
                textposition='outside',
                hovertemplate='<b>%{x}</b><br>تغییر: %{y:+,}<extra></extra>'
            ),
            row=2, col=1
        )
    
    fig.update_layout(
        title={
            'text': '📊 مقایسه کامل گزارش‌ها با یکدیگر',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50'}
        },
        height=700,
        template='plotly_white',
        font=dict(family='Vazirmatn, Tahoma', size=11),
        showlegend=False,
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    return fig


def create_progress_bar_chart(progress_df):
    """نمودار میله‌ای پیشرفت استان‌ها"""
    if progress_df.empty:
        return None
    
    fig = go.Figure()
    
    colors = ['#2ecc71' if x >= 75 else ('#f39c12' if x >= 50 else ('#e67e22' if x >= 25 else '#e74c3c')) 
              for x in progress_df['درصد پیشرفت']]
    
    fig.add_trace(go.Bar(
        x=progress_df['استان'],
        y=progress_df['درصد پیشرفت'],
        marker=dict(
            color=colors,
            line=dict(color='white', width=2)
        ),
        text=progress_df['درصد پیشرفت'].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>پیشرفت: %{y:.1f}%<extra></extra>'
    ))
    
    fig.add_hline(y=50, line_dash="dash", line_color="gray", 
                  annotation_text="هدف: 50%", annotation_position="right")
    
    fig.update_layout(
        title={
            'text': '📊 درصد پیشرفت رفع مغایرت به تفکیک استان',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50'}
        },
        xaxis_title='استان',
        yaxis_title='درصد پیشرفت (%)',
        yaxis_range=[0, max(110, progress_df['درصد پیشرفت'].max() * 1.1)],
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=11),
        showlegend=False,
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    fig.update_xaxes(showgrid=False, tickangle=-45)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    
    return fig


def create_comparison_bar_chart(progress_df):
    """نمودار مقایسه‌ای گروهی"""
    if progress_df.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='مغایرت اولیه',
        x=progress_df['استان'],
        y=progress_df['مغایرت اولیه'],
        marker_color='#e74c3c',
        text=progress_df['مغایرت اولیه'],
        textposition='auto',
    ))
    
    fig.add_trace(go.Bar(
        name='مغایرت فعلی',
        x=progress_df['استان'],
        y=progress_df['مغایرت فعلی'],
        marker_color='#3498db',
        text=progress_df['مغایرت فعلی'],
        textposition='auto',
    ))
    
    fig.add_trace(go.Bar(
        name='رفع شده',
        x=progress_df['استان'],
        y=progress_df['رفع شده'],
        marker_color='#2ecc71',
        text=progress_df['رفع شده'],
        textposition='auto',
    ))
    
    fig.update_layout(
        title={
            'text': '📊 مقایسه جامع مغایرت‌ها در استان‌ها',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20, 'color': '#2c3e50'}
        },
        xaxis_title='استان',
        yaxis_title='تعداد',
        barmode='group',
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=11),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
        plot_bgcolor='rgba(248, 249, 250, 0.8)',
        paper_bgcolor='white',
    )
    
    fig.update_xaxes(showgrid=False, tickangle=-45)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.08)')
    
    return fig


def create_pie_chart(df, cols):
    """نمودار دایره‌ای توزیع"""
    if not cols['province'] or cols['province'] not in df.columns:
        return None
    
    province_counts = df.groupby(cols['province']).size().reset_index(name='تعداد')
    province_counts = province_counts.sort_values('تعداد', ascending=False).head(10)
    
    fig = go.Figure(data=[go.Pie(
        labels=province_counts[cols['province']],
        values=province_counts['تعداد'],
        hole=0.4,
        marker=dict(
            colors=px.colors.qualitative.Set3,
            line=dict(color='white', width=2)
        ),
        textinfo='label+percent',
        hovertemplate='<b>%{label}</b><br>تعداد: %{value:,}<br>درصد: %{percent}<extra></extra>'
    )])
    
    fig.update_layout(
        title={
            'text': '🎯 توزیع درصدی مغایرت‌ها (10 استان برتر)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#2c3e50'}
        },
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=11),
    )
    
    return fig


def create_heatmap(df, cols):
    """نقشه حرارتی مغایرت‌ها"""
    if not cols['province'] or cols['province'] not in df.columns or 'تاریخ شمسی' not in df.columns:
        return None
    
    pivot_data = df.groupby([cols['province'], 'تاریخ شمسی']).size().reset_index(name='تعداد')
    pivot_table = pivot_data.pivot(index=cols['province'], columns='تاریخ شمسی', values='تعداد').fillna(0)
    
    fig = go.Figure(data=go.Heatmap(
        z=pivot_table.values,
        x=pivot_table.columns,
        y=pivot_table.index,
        colorscale='YlOrRd',
        hovertemplate='استان: %{y}<br>تاریخ: %{x}<br>تعداد: %{z:,}<extra></extra>',
        colorbar=dict(title='تعداد')
    ))
    
    fig.update_layout(
        title={
            'text': '🔥 نقشه حرارتی مغایرت‌ها (استان × تاریخ)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#2c3e50'}
        },
        xaxis_title='تاریخ',
        yaxis_title='',
        template='plotly_white',
        height=600,
        font=dict(family='Vazirmatn, Tahoma', size=13),
    )
    
    return fig

def predict_future_trend(df, cols, periods=3):
    """پیش‌بینی روند آینده با استفاده از Linear Regression"""
    if 'تاریخ شمسی' not in df.columns or df['تاریخ شمسی'].nunique() < 3:
        return None, None
    
    daily_counts = df.groupby('تاریخ شمسی').size().reset_index(name='تعداد')
    daily_counts = daily_counts.sort_values('تاریخ شمسی')
    
    # تبدیل به اعداد برای regression
    X = np.arange(len(daily_counts)).reshape(-1, 1)
    y = daily_counts['تعداد'].values
    
    # Linear Regression ساده
    z = np.polyfit(range(len(daily_counts)), y, 1)
    p = np.poly1d(z)
    
    # پیش‌بینی برای دوره‌های آینده
    future_X = np.arange(len(daily_counts), len(daily_counts) + periods)
    predictions = p(future_X)
    
    # ساخت تاریخ‌های آینده (فرضی)
    last_date = daily_counts['تاریخ شمسی'].iloc[-1]
    future_dates = [f"{last_date} + {i+1}" for i in range(periods)]
    
    prediction_df = pd.DataFrame({
        'دوره': future_dates,
        'پیش‌بینی تعداد مغایرت': predictions.astype(int),
        'روند': ['📉 کاهشی' if z[0] < 0 else '📈 افزایشی' if z[0] > 0 else '➡️ ثابت'] * periods,
        'شیب روند': [round(z[0], 2)] * periods
    })
    
    # ساخت نمودار
    fig = go.Figure()
    
    # داده‌های واقعی
    fig.add_trace(go.Scatter(
        x=list(range(len(daily_counts))),
        y=daily_counts['تعداد'],
        mode='lines+markers',
        name='داده واقعی',
        line=dict(color='#3498db', width=3),
        marker=dict(size=8)
    ))
    
    # خط روند
    fig.add_trace(go.Scatter(
        x=list(range(len(daily_counts))),
        y=p(range(len(daily_counts))),
        mode='lines',
        name='خط روند',
        line=dict(color='red', width=2, dash='dash')
    ))
    
    # پیش‌بینی
    fig.add_trace(go.Scatter(
        x=list(future_X),
        y=predictions,
        mode='lines+markers',
        name='پیش‌بینی',
        line=dict(color='green', width=3, dash='dot'),
        marker=dict(size=10, symbol='star')
    ))
    
    # خط جداکننده
    fig.add_vline(x=len(daily_counts)-0.5, line_dash="solid", line_color="gray", 
                  annotation_text="آخرین داده", annotation_position="top")
    
    fig.update_layout(
        title={
            'text': f'📊 پیش‌بینی روند {periods} دوره آینده',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#2c3e50'}
        },
        xaxis_title='دوره زمانی',
        yaxis_title='تعداد مغایرت',
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=11),
        hovermode='x unified'
    )
    
    return prediction_df, fig


def create_pareto_chart(issue_types_df):
    """نمودار Pareto برای انواع مغایرت"""
    if issue_types_df.empty:
        return None
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Bar(
            x=issue_types_df['نوع مغایرت'],
            y=issue_types_df['تعداد'],
            name='تعداد',
            marker_color='#3498db',
            text=issue_types_df['تعداد'],
            textposition='outside'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(
            x=issue_types_df['نوع مغایرت'],
            y=issue_types_df['درصد تجمعی'],
            name='درصد تجمعی',
            mode='lines+markers',
            line=dict(color='red', width=3),
            marker=dict(size=8)
        ),
        secondary_y=True
    )
    
    fig.add_hline(y=80, line_dash="dash", line_color="gray", secondary_y=True,
                  annotation_text="قانون 80/20", annotation_position="right")
    
    fig.update_xaxes(title_text="نوع مغایرت", tickangle=-45)
    fig.update_yaxes(title_text="تعداد", secondary_y=False)
    fig.update_yaxes(title_text="درصد تجمعی (%)", secondary_y=True, range=[0, 105])
    
    fig.update_layout(
        title={
            'text': '📊 تحلیل Pareto انواع مغایرت (قانون 80/20)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#2c3e50'}
        },
        template='plotly_white',
        height=500,
        font=dict(family='Vazirmatn, Tahoma', size=11),
        hovermode='x unified'
    )
    
    return fig


def download_chart_as_html(fig, filename):
    """دانلود نمودار به صورت HTML"""
    if fig is None:
        return None
    
    buffer = io.StringIO()
    fig.write_html(buffer)
    html_bytes = buffer.getvalue().encode()
    
    b64 = base64.b64encode(html_bytes).decode()
    href = f'<a href="data:text/html;base64,{b64}" download="{filename}.html" style="text-decoration: none;"><button style="background: #667eea; color: white; border: none; padding: 8px 16px; border-radius: 5px; cursor: pointer;">📥 دانلود نمودار</button></a>'
    return href


def save_chart_as_image(fig, width=1600, height=900, scale=3):
    """ذخیره نمودار به صورت تصویر با zoom out"""
    if fig is None:
        return None
    
    try:
        fig.update_layout(
            width=width,
            height=height,
            font=dict(size=14)
        )
        
        img_bytes = pio.to_image(fig, format='png', width=width, height=height, scale=scale)
        return img_bytes
    except Exception as e:
        st.warning(f"خطا در ذخیره تصویر: {str(e)}")
        return None


def create_excel_with_images(df, files_info, comparison_df, progress_df, repeated_df, new_issues_df, 
                             issue_types_df, benchmark_df, stats, all_charts):
    """ساخت فایل Excel با تصاویر نمودارها"""
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='All_Data', index=False)
        files_info.to_excel(writer, sheet_name='Files_Info', index=False)
        
        if comparison_df is not None and not comparison_df.empty:
            comparison_df.to_excel(writer, sheet_name='Reports_Comparison', index=False)
        
        if not progress_df.empty:
            progress_df.to_excel(writer, sheet_name='Provinces_Progress', index=False)
        
        if not repeated_df.empty:
            repeated_df.to_excel(writer, sheet_name='Repeated_Issues', index=False)
        
        if not new_issues_df.empty:
            new_issues_df.to_excel(writer, sheet_name='New_Issues', index=False)
        
        if not issue_types_df.empty:
            issue_types_df.to_excel(writer, sheet_name='Issue_Types_Pareto', index=False)
        
        if not benchmark_df.empty:
            benchmark_df.to_excel(writer, sheet_name='Benchmark_Analysis', index=False)
        
        pd.DataFrame([stats]).to_excel(writer, sheet_name='Summary_Stats', index=False)
        
        workbook = writer.book
        
        chart_sheet = workbook.create_sheet('Charts')
        
        row_position = 1
        
        for chart_name, img_bytes in all_charts.items():
            if img_bytes:
                try:
                    img_temp = io.BytesIO(img_bytes)
                    img = XLImage(img_temp)
                    
                    img.width = int(img.width * 0.4)
                    img.height = int(img.height * 0.4)
                    
                    # اضافه کردن عنوان با Font صحیح
                    cell = chart_sheet.cell(row=row_position, column=1, value=chart_name)
                    cell.font = Font(size=14, bold=True, color='0066CC')
                    
                    # اضافه کردن تصویر
                    img.anchor = f'A{row_position + 1}'
                    chart_sheet.add_image(img)
                    
                    row_position += 45
                    
                except Exception as e:
                    st.warning(f"خطا در اضافه کردن {chart_name}: {str(e)}")
                    continue
    
    return output.getvalue()


def main():
    PLOTLY_CONFIG = {
        'displayModeBar': True,
        'displaylogo': False,
        'modeBarButtonsToRemove': ['pan2d', 'lasso2d', 'select2d'],
        'toImageButtonOptions': {
            'format': 'png',
            'filename': 'chart',
            'height': 800,
            'width': 1200,
            'scale': 2
        }
    }
    
    st.markdown("""
        <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 50px; border-radius: 25px; margin-bottom: 30px; 
                    box-shadow: 0 15px 35px rgba(0,0,0,0.3);'>
            <h1 style='color: white; text-align: center; margin: 0; font-size: 48px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);'>
                📊 سامانه کامل تحلیل مغایرت‌ها
            </h1>
            <p style='color: rgba(255,255,255,0.95); text-align: center; margin-top: 20px; font-size: 20px;'>
                تحلیل هوشمند • مقایسه دقیق • گزارش‌گیری جامع • داشبورد اجرایی
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    if not JALALI_AVAILABLE:
        st.warning("⚠️ برای نمایش تاریخ شمسی: `pip install jdatetime`")
    
    with st.sidebar:
        st.markdown("### 📁 آپلود فایل‌های Excel")
        
        uploaded_files = st.file_uploader(
            "فایل‌های مغایرت را انتخاب کنید",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="برای مقایسه دقیق، حداقل 2 فایل آپلود کنید"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} فایل بارگذاری شده")
            st.markdown("---")
            st.markdown("### ⚙️ تنظیمات نمایش")
            show_raw_data = st.checkbox("📋 نمایش داده‌های خام", value=False)
            show_advanced = st.checkbox("🔬 نمودارهای پیشرفته", value=True)
            
            st.markdown("---")
            st.markdown("### 🎯 فیلترهای پیشرفته")
            
            # فیلتر بازه زمانی
            if uploaded_files:
                st.markdown("#### 📅 فیلتر زمانی")
                filter_date = st.checkbox("فعال‌سازی فیلتر تاریخ", value=False)
            
            st.markdown("---")
            st.markdown("### 📊 آمار سریع")
    
    if not uploaded_files:
        col1, col2, col3, col4 = st.columns(4)
        features = [
            ("🚀", "شروع سریع", "آپلود چند فایل Excel"),
            ("📊", "تحلیل قدرتمند", "نمودارهای تعاملی متنوع"),
            ("📅", "تاریخ شمسی", "تبدیل خودکار تاریخ‌ها"),
            ("💾", "دانلود کامل", "تمام نمودارها و گزارش‌ها")
        ]
        
        for col, (icon, title, desc) in zip([col1, col2, col3, col4], features):
            with col:
                st.markdown(f"""
                    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                                padding: 30px; border-radius: 15px; color: white; text-align: center; height: 200px;'>
                        <h1 style='font-size: 48px; margin: 0;'>{icon}</h1>
                        <h3 style='margin: 15px 0 10px 0;'>{title}</h3>
                        <p style='font-size: 14px; opacity: 0.9;'>{desc}</p>
                    </div>
                """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        with st.expander("📖 راهنمای کامل استفاده", expanded=True):
            st.markdown("""
            ### 🎯 ویژگی‌های کلیدی:
            - **داشبورد اجرایی**: خلاصه‌ای جامع از وضعیت با KPI های کلیدی
            - **تحلیل تکراری‌ها**: شناسایی هوشمند مغایرت‌های تکراری با وضعیت برطرف شدن
            - **پشتیبانی از تاریخ شمسی**: تبدیل خودکار تاریخ از نام فایل
            - **مقایسه گزارش‌ها**: تحلیل روند با خط Trend
            - **Benchmark استان‌ها**: مقایسه با میانگین کشوری
            - **تحلیل Pareto**: شناسایی 20% مغایرت‌هایی که 80% مشکل را ایجاد می‌کنند
            - **مغایرت‌های جدید**: شناسایی مغایرت‌های جدید در آخرین گزارش
            - **مقایسه دو استان**: تحلیل تفصیلی و یافتن مغایرت‌های مشترک
            - **دانلود گزارش کامل**: Excel با تمام تحلیل‌ها و تصاویر نمودارها

            ### 📋 فرمت مورد نیاز فایل‌ها:
            - نام فایل باید حاوی تاریخ میلادی با فرمت **YYYYMMDD** باشد
            - مثال: `Planning_Mismatch_20250831.xlsx`
            - فایل باید دارای ستون‌های **استان**، **کد سایت**، و ستون‌های مربوط به **نوع و عنوان مغایرت** باشد

            ### 🔑 نکات مهم:
            1. برای استفاده از تمام قابلیت‌ها، حداقل **2 فایل** با تاریخ‌های متفاوت آپلود کنید
            2. اطمینان حاصل کنید که نام ستون‌ها در تمام فایل‌ها **یکسان** است
            3. برای نصب کتابخانه‌های لازم: `pip install jdatetime kaleido openpyxl pillow`
            """)
        
        return
    
    with st.spinner('🔄 در حال بارگذاری و پردازش فایل‌ها...'):
        df, files_info = load_excel_files(uploaded_files)
    
    if df is None or df.empty:
        st.error("❌ خطا در خواندن فایل‌ها یا فایل‌ها خالی هستند")
        st.stop()
    
    # اعمال فیلتر زمانی اگر فعال باشد
    df_filtered = df.copy()
    if 'filter_date' in locals() and filter_date and 'تاریخ شمسی' in df.columns:
        dates_available = sorted(df['تاریخ شمسی'].unique())
        with st.sidebar:
            selected_dates = st.multiselect(
                "انتخاب تاریخ‌ها",
                options=dates_available,
                default=dates_available
            )
            if selected_dates:
                df_filtered = df[df['تاریخ شمسی'].isin(selected_dates)]
    
    cols = detect_columns(df_filtered)
    
    with st.sidebar:
        with st.expander("🔍 ستون‌های شناسایی شده"):
            for key, value in cols.items():
                icon = "✅" if value else "❌"
                st.write(f"{icon} **{key}:** {value or 'یافت نشد'}")
    
    stats = calculate_summary_stats(df_filtered, cols)
    progress_df = calculate_progress(df_filtered, cols)
    repeated_df = find_repeated_issues(df_filtered, cols)
    new_issues_df = find_new_issues(df_filtered, cols)
    issue_types_df = analyze_issue_types(df_filtered, cols)
    benchmark_df = calculate_benchmark(progress_df)
    comparison_df = compare_reports(df_filtered, cols)
    
    with st.sidebar:
        st.metric("📊 مجموع مغایرت‌ها", f"{stats['total_issues']:,}")
        st.metric("🏢 سایت‌های منحصر", f"{stats['unique_sites']:,}")
        st.metric("🗺️ استان‌های منحصر", f"{stats['unique_provinces']:,}")
        st.metric("📅 تعداد گزارش‌ها", f"{stats['total_dates']:,}")

    col1, col2, col3, col4 = st.columns(4)
    metrics = [
        ("📋 مجموع مغایرت‌ها", stats['total_issues'], "مغایرت"),
        ("🏢 تعداد سایت‌ها", stats['unique_sites'], "سایت"),
        ("🗺️ تعداد استان‌ها", stats['unique_provinces'], "استان"),
        ("📊 تعداد گزارش‌ها", stats['total_dates'], "گزارش")
    ]
    
    for col, (title, value, unit) in zip([col1, col2, col3, col4], metrics):
        with col:
            st.markdown(f"""
                <div class='metric-card'>
                    <h3 style='margin: 0; font-size: 15px; opacity: 0.95;'>{title}</h3>
                    <h1 style='margin: 20px 0 5px 0; font-size: 48px;'>{value:,}</h1>
                    <p style='margin: 0; font-size: 13px; opacity: 0.8;'>{unit}</p>
                </div>
            """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
        "🎯 داشبورد اجرایی",
        "📈 روند و تحلیل",
        "🔄 مقایسه گزارش‌ها",
        "📊 پیشرفت استان‌ها",
        "🔁 مغایرت‌های تکراری",
        "🆕 مغایرت‌های جدید",
        "📉 تحلیل Pareto",
        "🔮 پیش‌بینی روند",
        "🎨 تحلیل‌های پیشرفته",
        "💾 دانلود و گزارش"
    ])
    
    all_charts = {}
    
    with tab1:
        st.markdown("## 📊 داشبورد اجرایی - خلاصه وضعیت")
        
        # KPI های اصلی
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        
        with kpi1:
            st.markdown(f"""
                <div class='info-box' style='text-align: center;'>
                    <h4>📋 مجموع مغایرت‌ها</h4>
                    <h2 style='color: #e74c3c; margin: 10px 0;'>{stats['total_issues']:,}</h2>
                </div>
            """, unsafe_allow_html=True)
        
        with kpi2:
            if not progress_df.empty:
                avg_progress = progress_df['درصد پیشرفت'].mean()
                st.markdown(f"""
                    <div class='info-box' style='text-align: center;'>
                        <h4>📈 میانگین پیشرفت</h4>
                        <h2 style='color: #3498db; margin: 10px 0;'>{avg_progress:.1f}%</h2>
                    </div>
                """, unsafe_allow_html=True)
        
        with kpi3:
            if not progress_df.empty:
                total_resolved = progress_df['رفع شده'].sum()
                st.markdown(f"""
                    <div class='info-box' style='text-align: center;'>
                        <h4>✅ رفع شده</h4>
                        <h2 style='color: #2ecc71; margin: 10px 0;'>{total_resolved:,}</h2>
                    </div>
                """, unsafe_allow_html=True)
        
        with kpi4:
            if not repeated_df.empty:
                not_resolved = len(repeated_df[repeated_df['وضعیت رفع'].str.contains('برطرف نشده', na=False)])
                st.markdown(f"""
                    <div class='info-box' style='text-align: center;'>
                        <h4>🔁 تکراری و فعال</h4>
                        <h2 style='color: #f39c12; margin: 10px 0;'>{not_resolved:,}</h2>
                    </div>
                """, unsafe_allow_html=True)
        
        with kpi5:
            if not new_issues_df.empty:
                st.markdown(f"""
                    <div class='info-box' style='text-align: center;'>
                        <h4>🆕 جدید</h4>
                        <h2 style='color: #9b59b6; margin: 10px 0;'>{len(new_issues_df):,}</h2>
                    </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # بهترین و بدترین عملکردها
        col1, col2 = st.columns(2)
        
        with col1:
            if not progress_df.empty and len(progress_df) >= 3:
                top3 = progress_df.head(3)
                st.markdown("""
                    <div class='success-box'>
                        <h3>⭐ برترین عملکردها</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                for idx, row in top3.iterrows():
                    st.markdown(f"""
                        <div class='info-box'>
                            <h4>{row['استان']}</h4>
                            <p>پیشرفت: <strong>{row['درصد پیشرفت']:.1f}%</strong></p>
                            <p>رفع شده: <strong>{row['رفع شده']:,}</strong> از {row['مغایرت اولیه']:,}</p>
                        </div>
                    """, unsafe_allow_html=True)
        
        with col2:
            if not progress_df.empty and len(progress_df) >= 3:
                bottom3 = progress_df.tail(3).iloc[::-1]
                st.markdown("""
                    <div class='warning-box'>
                        <h3>⚠️ نیازمند توجه</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                for idx, row in bottom3.iterrows():
                    st.markdown(f"""
                        <div class='info-box'>
                            <h4>{row['استان']}</h4>
                            <p>پیشرفت: <strong>{row['درصد پیشرفت']:.1f}%</strong></p>
                            <p>باقیمانده: <strong>{row['باقیمانده']:,}</strong></p>
                        </div>
                    """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # هشدارهای مهم
        st.markdown("""
            <div class='highlight-box'>
                <h3>🚨 هشدارهای مهم</h3>
            </div>
        """, unsafe_allow_html=True)
        
        warnings_list = []
        
        if not repeated_df.empty:
            critical_repeated = len(repeated_df[
                (repeated_df['اولویت'] == '🔴 بحرانی') & 
                (repeated_df['وضعیت رفع'].str.contains('برطرف نشده', na=False))
            ])
            if critical_repeated > 0:
                warnings_list.append(f"🔴 {critical_repeated} مغایرت بحرانی تکراری و برطرف نشده")
        
        if not new_issues_df.empty and len(new_issues_df) > 10:
            warnings_list.append(f"🆕 {len(new_issues_df)} مغایرت جدید در آخرین گزارش ظاهر شده")
        
        if not progress_df.empty:
            low_progress = len(progress_df[progress_df['درصد پیشرفت'] < 25])
            if low_progress > 0:
                warnings_list.append(f"⚠️ {low_progress} استان با پیشرفت کمتر از 25%")
        
        if warnings_list:
            for warning in warnings_list:
                st.markdown(f"""
                    <div class='info-box' style='border-left-color: #e74c3c;'>
                        <p style='margin: 0; font-size: 16px;'>{warning}</p>
                    </div>
                """, unsafe_allow_html=True)
        else:
            st.success("✅ هیچ هشدار بحرانی وجود ندارد")
    
    with tab2:
        st.markdown("### 📈 روند کلی مغایرت‌ها در کل کشور")
        trend_fig = create_trend_chart(df_filtered)
        if trend_fig:
            st.plotly_chart(trend_fig, config=PLOTLY_CONFIG)
            st.markdown(download_chart_as_html(trend_fig, "trend_chart_total"), unsafe_allow_html=True)
            all_charts['روند کلی مغایرت‌ها'] = save_chart_as_image(trend_fig)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""
                <div class='info-box'>
                    <h4>📊 خلاصه آماری</h4>
                    <p><strong>بازه زمانی:</strong> {stats['date_range']}</p>
                    <p><strong>مجموع مغایرت‌ها:</strong> {stats['total_issues']:,}</p>
                    <p><strong>تعداد گزارش‌ها:</strong> {stats['total_dates']}</p>
                    <p><strong>تعداد فایل‌ها:</strong> {stats['files_count']}</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col2:
            if not progress_df.empty:
                avg_progress = progress_df['درصد پیشرفت'].mean()
                total_resolved = progress_df['رفع شده'].sum()
                st.markdown(f"""
                    <div class='info-box'>
                        <h4>✅ پیشرفت کلی</h4>
                        <p><strong>میانگین پیشرفت:</strong> {avg_progress:.1f}%</p>
                        <p><strong>مجموع رفع شده:</strong> {total_resolved:,}</p>
                        <p><strong>بهترین استان:</strong> {progress_df.iloc[0]['استان']}</p>
                    </div>
                """, unsafe_allow_html=True)
        
        st.markdown("### 🗺️ توزیع مغایرت‌ها در استان‌ها")
        province_fig = create_province_chart(df_filtered, cols)
        if province_fig:
            st.plotly_chart(province_fig, config=PLOTLY_CONFIG)
            st.markdown(download_chart_as_html(province_fig, "province_chart_distribution"), unsafe_allow_html=True)
            all_charts['توزیع استان‌ها'] = save_chart_as_image(province_fig, height=1000)

    with tab3:
        if comparison_df is not None and not comparison_df.empty and len(comparison_df) > 1:
            st.markdown("### 📊 مقایسه تمام گزارش‌ها")
            comparison_fig = create_comparison_chart(comparison_df)
            if comparison_fig:
                st.plotly_chart(comparison_fig, config=PLOTLY_CONFIG)
                st.markdown(download_chart_as_html(comparison_fig, "comparison_chart"), unsafe_allow_html=True)
                all_charts['مقایسه گزارش‌ها'] = save_chart_as_image(comparison_fig, height=1000)
            
            st.markdown("### 📋 جدول مقایسه تفصیلی")
            st.dataframe(comparison_df)
        else:
            st.info("ℹ️ برای مقایسه، حداقل 2 گزارش با تاریخ‌های مختلف لازم است.")

    with tab4:
        if not progress_df.empty:
            st.markdown("### 📊 درصد پیشرفت استان‌ها")
            progress_fig = create_progress_bar_chart(progress_df)
            if progress_fig:
                st.plotly_chart(progress_fig, config=PLOTLY_CONFIG)
                st.markdown(download_chart_as_html(progress_fig, "progress_chart"), unsafe_allow_html=True)
                all_charts['درصد پیشرفت استان‌ها'] = save_chart_as_image(progress_fig)
            
            st.markdown("### 📊 مقایسه تفصیلی مغایرت‌ها (اولیه، فعلی، رفع شده)")
            comparison_bar_fig = create_comparison_bar_chart(progress_df)
            if comparison_bar_fig:
                st.plotly_chart(comparison_bar_fig, config=PLOTLY_CONFIG)
                st.markdown(download_chart_as_html(comparison_bar_fig, "comparison_bar_chart"), unsafe_allow_html=True)
                all_charts['مقایسه تفصیلی مغایرت‌ها'] = save_chart_as_image(comparison_bar_fig)
            
            st.markdown("---")
            st.markdown("### 🎯 تحلیل Benchmark - مقایسه با میانگین کشوری")
            if not benchmark_df.empty:
                st.dataframe(benchmark_df.style.format({
                    'درصد پیشرفت': '{:.2f}%',
                    'میانگین کشوری': '{:.2f}%',
                    'میانه کشوری': '{:.2f}%',
                    'انحراف از میانگین': '{:.2f}'
                }))
            
            st.markdown("---")
            st.markdown("### 📋 جدول پیشرفت تفصیلی")
            st.dataframe(progress_df.style.format({
                'درصد پیشرفت': '{:.2f}%'
            }))

            st.markdown("---")
            st.markdown("### 📈 نمودارهای روند پیشرفت برای هر استان")
            st.info("نمودارهای زیر روند تعداد مغایرت‌ها را در طول زمان برای هر استان به صورت جداگانه نمایش می‌دهند.")
            
            provinces_with_progress = progress_df['استان'].tolist()
            
            if not provinces_with_progress:
                st.warning("استانی برای نمایش نمودار یافت نشد.")
            else:
                num_columns = 2
                chart_cols = st.columns(num_columns)
                for idx, province in enumerate(provinces_with_progress):
                    with chart_cols[idx % num_columns]:
                        province_fig = create_province_progress_chart(df_filtered, cols, province)
                        if province_fig:
                            st.plotly_chart(province_fig, config=PLOTLY_CONFIG)
                            all_charts[f'روند {province}'] = save_chart_as_image(province_fig, width=1400, height=600)

        else:
            st.info("ℹ️ برای محاسبه پیشرفت، حداقل 2 گزارش با تاریخ‌های مختلف لازم است.")

    with tab5:
        if not repeated_df.empty:
            st.markdown("### 📋 مغایرت‌های تکراری")
            st.info("این جدول مغایرت‌هایی را نشان می‌دهد که در گزارش‌های مختلف تکرار شده‌اند.")
            
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                total_repeated = len(repeated_df)
                st.metric("مجموع تکراری‌ها", f"{total_repeated:,}")
            with col_stat2:
                resolved = len(repeated_df[repeated_df['وضعیت رفع'].str.contains('برطرف شده', na=False)])
                st.metric("✅ برطرف شده", f"{resolved:,}", delta=f"{resolved/total_repeated*100:.1f}%")
            with col_stat3:
                not_resolved = len(repeated_df[repeated_df['وضعیت رفع'].str.contains('برطرف نشده', na=False)])
                st.metric("❌ برطرف نشده", f"{not_resolved:,}", delta=f"-{not_resolved/total_repeated*100:.1f}%", delta_color="inverse")
            
            st.markdown("---")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                min_repeat = st.slider("حداقل تعداد تکرار", 2, int(repeated_df['تعداد تکرار'].max()), 2)
            with col2:
                priority_filter = st.multiselect("فیلتر اولویت", options=repeated_df['اولویت'].unique(), default=repeated_df['اولویت'].unique())
            with col3:
                if 'استان' in repeated_df.columns:
                    province_filter = st.multiselect("فیلتر استان", options=sorted(repeated_df['استان'].dropna().unique()))
            with col4:
                status_filter = st.multiselect(
                    "وضعیت رفع",
                    options=repeated_df['وضعیت رفع'].unique(),
                    default=repeated_df['وضعیت رفع'].unique()
                )
            
            filtered_repeated = repeated_df[repeated_df['تعداد تکرار'] >= min_repeat]
            if priority_filter:
                filtered_repeated = filtered_repeated[filtered_repeated['اولویت'].isin(priority_filter)]
            if province_filter:
                filtered_repeated = filtered_repeated[filtered_repeated['استان'].isin(province_filter)]
            if status_filter:
                filtered_repeated = filtered_repeated[filtered_repeated['وضعیت رفع'].isin(status_filter)]

            st.dataframe(filtered_repeated, height=500)
        else:
            st.success("✅ هیچ مغایرت تکراری بین گزارش‌های مختلف یافت نشد!")
    
    with tab6:
        if not new_issues_df.empty:
            st.markdown("### 🆕 مغایرت‌های جدید در آخرین گزارش")
            st.warning(f"⚠️ {len(new_issues_df)} مغایرت جدید در آخرین گزارش شناسایی شد که در گزارش قبلی وجود نداشت.")
            
            if 'استان' in new_issues_df.columns:
                province_new = new_issues_df['استان'].value_counts().reset_index()
                province_new.columns = ['استان', 'تعداد جدید']
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown("#### 📊 توزیع مغایرت‌های جدید به تفکیک استان")
                    st.dataframe(province_new)
                with col2:
                    st.markdown("#### 📈 آمار")
                    st.metric("مجموع جدید", len(new_issues_df))
                    st.metric("تعداد استان‌ها", len(province_new))
            
            st.markdown("---")
            st.markdown("#### 📋 جزئیات مغایرت‌های جدید")
            st.dataframe(new_issues_df, height=400)
        else:
            st.success("✅ هیچ مغایرت جدیدی در آخرین گزارش نسبت به گزارش قبلی شناسایی نشد!")
    
    with tab7:
        if not issue_types_df.empty:
            st.markdown("### 📉 تحلیل Pareto - قانون 80/20")
            st.info("این تحلیل نشان می‌دهد کدام انواع مغایرت بیشترین تاثیر را دارند. معمولاً 20% از انواع مغایرت، 80% مشکلات را ایجاد می‌کنند.")
            
            pareto_fig = create_pareto_chart(issue_types_df)
            if pareto_fig:
                st.plotly_chart(pareto_fig, config=PLOTLY_CONFIG)
                all_charts['تحلیل Pareto'] = save_chart_as_image(pareto_fig)
            
            st.markdown("---")
            st.markdown("### 📋 جدول تفصیلی انواع مغایرت")
            
            critical_issues = issue_types_df[issue_types_df['دسته‌بندی'] == '🔴 بحرانی (80%)']
            if not critical_issues.empty:
                st.markdown("""
                    <div class='warning-box'>
                        <h4>🎯 مغایرت‌های بحرانی (80% اول)</h4>
                        <p>تمرکز روی این موارد بیشترین تاثیر را در کاهش مغایرت‌ها خواهد داشت</p>
                    </div>
                """, unsafe_allow_html=True)
            
            st.dataframe(issue_types_df.style.format({
                'درصد': '{:.2f}%',
                'درصد تجمعی': '{:.2f}%'
            }), height=400)
        else:
            st.info("ℹ️ داده‌های کافی برای تحلیل Pareto موجود نیست.")
    
    with tab8:
        st.markdown("### 🔮 پیش‌بینی روند آینده")
        st.info("این بخش با استفاده از Linear Regression، روند آینده مغایرت‌ها را پیش‌بینی می‌کند.")
        
        col1, col2 = st.columns([1, 3])
        with col1:
            periods = st.slider("تعداد دوره‌های آینده", 1, 10, 3)
        
        prediction_df, prediction_fig = predict_future_trend(df_filtered, cols, periods)
        
        if prediction_df is not None and prediction_fig is not None:
            with col2:
                if prediction_df['شیب روند'].iloc[0] < -5:
                    st.success(f"📉 روند کاهشی قوی: {prediction_df['شیب روند'].iloc[0]}")
                elif prediction_df['شیب روند'].iloc[0] < 0:
                    st.info(f"📉 روند کاهشی: {prediction_df['شیب روند'].iloc[0]}")
                elif prediction_df['شیب روند'].iloc[0] > 5:
                    st.error(f"📈 روند افزایشی قوی: {prediction_df['شیب روند'].iloc[0]}")
                elif prediction_df['شیب روند'].iloc[0] > 0:
                    st.warning(f"📈 روند افزایشی: {prediction_df['شیب روند'].iloc[0]}")
                else:
                    st.info("➡️ روند تقریباً ثابت")
            
            st.plotly_chart(prediction_fig, config=PLOTLY_CONFIG)
            all_charts['پیش‌بینی روند'] = save_chart_as_image(prediction_fig)
            
            st.markdown("---")
            st.markdown("### 📋 جدول پیش‌بینی")
            st.dataframe(prediction_df)
            
            st.markdown("""
                <div class='info-box'>
                    <h4>⚠️ توجه</h4>
                    <p>این پیش‌بینی بر اساس روند خطی گذشته است و عوامل خارجی را در نظر نمی‌گیرد.</p>
                    <p>برای تصمیم‌گیری مهم، حتماً عوامل دیگر را نیز بررسی کنید.</p>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ برای پیش‌بینی، حداقل 3 گزارش با تاریخ‌های مختلف لازم است.")



    with tab9:
        if show_advanced:
            st.markdown("### 🎯 توزیع درصدی استان‌ها (10 استان برتر)")
            pie_fig = create_pie_chart(df_filtered, cols)
            if pie_fig:
                st.plotly_chart(pie_fig, config=PLOTLY_CONFIG)
                st.markdown(download_chart_as_html(pie_fig, "pie_chart"), unsafe_allow_html=True)
                all_charts['توزیع درصدی استان‌ها'] = save_chart_as_image(pie_fig)
            
            st.markdown("### 🔥 نقشه حرارتی مغایرت‌ها (استان × تاریخ)")
            heatmap_fig = create_heatmap(df_filtered, cols)
            if heatmap_fig:
                st.plotly_chart(heatmap_fig, config=PLOTLY_CONFIG)
                st.markdown(download_chart_as_html(heatmap_fig, "heatmap"), unsafe_allow_html=True)
                all_charts['نقشه حرارتی'] = save_chart_as_image(heatmap_fig, height=800)
            
            st.markdown("---")
            st.markdown("### 🔍 مقایسه دو استان")
            
            if cols['province'] and cols['province'] in df_filtered.columns:
                provinces_list = sorted(df_filtered[cols['province']].unique())
                
                col1, col2 = st.columns(2)
                with col1:
                    province1 = st.selectbox("استان اول", provinces_list, key='prov1')
                with col2:
                    province2 = st.selectbox("استان دوم", provinces_list, key='prov2')
                
                if st.button("🔍 مقایسه استان‌ها"):
                    comparison_result, common_issues = compare_two_provinces(df_filtered, cols, province1, province2)
                    
                    if comparison_result is not None:
                        st.markdown(f"#### 📊 مقایسه {province1} و {province2}")
                        st.dataframe(comparison_result)
                        
                        if not common_issues.empty:
                            st.markdown("---")
                            st.markdown("#### 🔗 مغایرت‌های مشترک")
                            st.dataframe(common_issues)
        else:
            st.info("☑️ برای نمایش نمودارهای پیشرفته، گزینه را از سایدبار فعال کنید.")
    
    with tab10:
        st.markdown("""
            <div class='download-section'>
                <h2 style='text-align: center; color: #2c3e50;'>💾 دانلود گزارش‌ها و نمودارها</h2>
                <p style='text-align: center; color: #7f8c8d;'>گزارش کامل شامل تمام داده‌ها، تحلیل‌ها و تصاویر نمودارها با کیفیت بالا</p>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📊 دانلود Excel ساده (بدون تصاویر)")
            output_simple = io.BytesIO()
            with pd.ExcelWriter(output_simple, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, sheet_name='All_Data', index=False)
                files_info.to_excel(writer, sheet_name='Files_Info', index=False)
                if comparison_df is not None and not comparison_df.empty:
                    comparison_df.to_excel(writer, sheet_name='Reports_Comparison', index=False)
                if not progress_df.empty:
                    progress_df.to_excel(writer, sheet_name='Provinces_Progress', index=False)
                if not repeated_df.empty:
                    repeated_df.to_excel(writer, sheet_name='Repeated_Issues', index=False)
                if not new_issues_df.empty:
                    new_issues_df.to_excel(writer, sheet_name='New_Issues', index=False)
                if not issue_types_df.empty:
                    issue_types_df.to_excel(writer, sheet_name='Issue_Types_Pareto', index=False)
                if not benchmark_df.empty:
                    benchmark_df.to_excel(writer, sheet_name='Benchmark_Analysis', index=False)
                pd.DataFrame([stats]).to_excel(writer, sheet_name='Summary_Stats', index=False)
            
            st.download_button(
                "📥 دانلود Excel ساده",
                data=output_simple.getvalue(),
                file_name=f'Mismatch_Analysis_Simple_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        
        with col2:
            st.markdown("#### 📊 دانلود Excel کامل (با تصاویر نمودارها)")
            
            if st.button("🎨 ساخت فایل کامل با تصاویر", type="primary"):
                with st.spinner('⏳ در حال ساخت فایل کامل... (ممکن است چند دقیقه طول بکشد)'):
                    try:
                        excel_with_images = create_excel_with_images(
                            df_filtered, files_info, comparison_df, progress_df, 
                            repeated_df, new_issues_df, issue_types_df, benchmark_df,
                            stats, all_charts
                        )
                        
                        st.download_button(
                            "📥 دانلود Excel کامل با نمودارها",
                            data=excel_with_images,
                            file_name=f'Mismatch_Analysis_Complete_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            type="primary"
                        )
                        st.success("✅ فایل کامل آماده دانلود است!")
                    except Exception as e:
                        st.error(f"❌ خطا در ساخت فایل: {str(e)}")
                        st.info("💡 اگر خطا مربوط به kaleido است، لطفاً آن را نصب کنید: `pip install kaleido`")
        
        st.markdown("---")
        
        # دانلود جداگانه جداول
        with st.expander("📋 دانلود جداول به صورت جداگانه"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if not repeated_df.empty:
                    csv_repeated = repeated_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 مغایرت‌های تکراری (CSV)",
                        data=csv_repeated,
                        file_name=f'Repeated_Issues_{datetime.now().strftime("%Y%m%d")}.csv',
                        mime='text/csv'
                    )
            
            with col2:
                if not new_issues_df.empty:
                    csv_new = new_issues_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 مغایرت‌های جدید (CSV)",
                        data=csv_new,
                        file_name=f'New_Issues_{datetime.now().strftime("%Y%m%d")}.csv',
                        mime='text/csv'
                    )
            
            with col3:
                if not progress_df.empty:
                    csv_progress = progress_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 پیشرفت استان‌ها (CSV)",
                        data=csv_progress,
                        file_name=f'Progress_{datetime.now().strftime("%Y%m%d")}.csv',
                        mime='text/csv'
                    )
        
        st.markdown("---")
        st.markdown("""
            <div class='info-box'>
                <h4>ℹ️ توضیحات</h4>
                <ul>
                    <li><strong>Excel ساده:</strong> فقط شامل داده‌های جدولی است (حجم کم، سرعت بالا)</li>
                    <li><strong>Excel کامل:</strong> شامل تمام داده‌ها + تصاویر نمودارها با کیفیت بالا و zoom out (حجم بیشتر، پردازش طولانی‌تر)</li>
                    <li>نمودارها با اندازه 1600×900 پیکسل و scale 3x ذخیره می‌شوند</li>
                    <li>تصاویر در شیت مجزای "Charts" قرار می‌گیرند</li>
                    <li><strong>فایل‌های CSV:</strong> قابل باز شدن در Excel و سایر نرم‌افزارها با پشتیبانی کامل از فارسی</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        
        # نمایش جداول داده در تب دانلود
        if show_raw_data:
            st.markdown("---")
            st.markdown("### 📊 اطلاعات فایل‌های بارگذاری شده")
            st.dataframe(files_info)
            
            st.markdown("---")
            st.markdown("### 📋 داده‌های خام تجمیع شده")
            st.dataframe(df_filtered, height=600)

    st.markdown("---")
    st.markdown("""
        <div style='text-align: center; padding: 20px; background: white; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);'>
            <p style='font-size: 12px; color: #95a5a6;'>نسخه 3.0 - سامانه جامع تحلیل مغایرت‌ها با داشبورد اجرایی</p>
            <p style='font-size: 11px; color: #bdc3c7; margin-top: 5px;'>شامل: داشبورد اجرایی • تحلیل Pareto • Benchmark • مغایرت‌های جدید • مقایسه استان‌ها</p>
        </div>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()