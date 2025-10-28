# Sales_Insights_Pro.py
#
# A professional, multi-lingual, multi-file-type Sales
# Dashboard and Forecasting tool.
# Version 3.2: Added Weekly/Monthly/Daily forecast frequency selection.
#
# Author: Sameh Sobhy Attia (Original)
# Refactored by: Gemini (Professional Upgrade)
#
# ---Dependencies---
# To run this app, you need Streamlit and other data
# libraries.
# Install them using pip:
# pip install streamlit pandas numpy plotly openpyxl reportlab lxml pdfplumber
#
# ---To Run---
# Save this file as "Sales_Insights_Pro_v3_2.py"
# In your terminal, run:
# streamlit run Sales_Insights_Pro_v3_2.py
#
# ---Features---
# - Caching for high-performance data processing.
# - Interactive Dashboard: Click/select rows to dynamically
# update charts.
# - Supports Excel, CSV, PDF, and HTML (table extraction)
# file uploads.
# - Fully bilingual (English/Arabic) UI.
# - Robust session state management (data persists across
# interactions).
# - Clean, tabbed interface.
# - Dark mode support.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import pdfplumber  # For reading PDF tables
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, \
    TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from typing import List, Dict, Tuple, Optional, Any, \
    BinaryIO
from lxml import etree  # Used for HTML parsing, openpyxl
# needs it

# ================================================
# 1. APP CONFIGURATION & INITIALIZATION
# ================================================

st.set_page_config(page_title='Sales Insights Pro',
                   layout='wide')

# Initialize session state
if 'lang' not in st.session_state:
    st.session_state['lang'] = 'en'
if 'df' not in st.session_state:
    st.session_state['df'] = None
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = None
if 'pivot_df' not in st.session_state:
    st.session_state['pivot_df'] = None
if 'kpi_cols' not in st.session_state:
    st.session_state['kpi_cols'] = []
if 'date_col' not in st.session_state:
    st.session_state['date_col'] = 'None'


# ================================================
# 2. TRANSLATIONS & LANGUAGE HELPER
# ================================================

TRANSLATIONS = {
    'en': {
        'title':
            'Sales Insights & Forecasting Pro',
        'upload':
            'Upload Data (Excel, CSV, PDF, HTML)',
        'upload_prompt': 'Upload a file to get started. Supported formats: \
Excel, CSV, PDF, HTML (tables).',
        'load_sample':
            'Load Sample Data',
        'data_loaded':
            'Successfully loaded',
        'rows':
            'rows',
        'cols':
            'columns',
        'total_everything': 'Total of all Numeric Columns',
        'grand_total':
            'Grand Total',
        'kpi_selection': 'Select KPI Columns (for totals, forecasting)',
        'date_column':
            'Select Date Column (for time series)',
        'pivot_config': 'Pivot Table Configuration',
        'row_field':
            'Row Field(s)',
        'col_field':
            'Column Field(s)',
        'agg_type':
            'Aggregation Type',
        'value_col':
            'Value Column',
        'generate_pivot': 'Generate Pivot Table',
        'stats_summary': 'Statistics Summary',
        'charts':
            'Charts & Visuals',
        'chart_type':
            'Chart Type',
        'x_axis':
            'X-Axis',
        'y_axis':
            'Y-Axis (multi-select)',
        'plot': 'Plot \
Chart',
        'forecasting':
            'Simple Forecasting (Trend)',
        'forecast_column': 'Select numeric column to forecast',
        'forecast_periods': 'Forecast Periods (steps)',
        'run_forecast': 'Run Forecast',
        'insights':
            'Automated Insights',
        'missing_values': 'Missing Values by Column',
        'correlations': 'Correlation Matrix (Numeric)',
        'download_excel': 'Download Summary as Excel',
        'download_html': 'Download Report as HTML',
        'download_pdf': 'Download Report as PDF',
        'language':
            'Language',
        'theme': 'Dark \
Mode',
        'show_data':
            'Show Raw Data',
        'download_pivot': 'Download Pivot as Excel',
        'config':
            'Column Configuration',
        'kpi_tab':
            'KPIs & Stats',
        'dashboard_tab': 'Interactive Dashboard',
        'pivot_tab':
            'Pivot Table',
        'charts_tab':
            'Manual Charts',
        'forecast_tab': 'Forecasting',
        'insights_tab': 'Data Insights',
        'export_tab':
            'Export Report',
        'selected_kpis': 'Totals for Selected KPIs',
        'no_kpis_selected': 'No KPI columns selected.',
        'no_numeric_stats': 'No numeric columns for statistics.',
        'plot_warn':
            'Please select at least one Y-Axis column.',
        'forecast_warn': 'Please select a numeric column to forecast.',
        'forecast_no_date': 'No date column selected. Forecasting on data \
index.',
        'forecast_no_data': 'Not enough data to forecast (need at least 2 data \
points).',
        'forecast_fail': 'Forecasting failed',
        'forecast_table': 'Forecast Table',
        'actual':
            'Actual',
        'forecast':
            'Forecast',
        'confidence':
            'Confidence Interval',
        'no_corr':
            'Not enough numeric columns for correlation.',
        'file_error':
            'Could not read file. Please ensure it is a valid format.',
        'pdf_warn':
            'PDF parsing found 0 tables. Please check the file.',
        'html_warn':
            'HTML parsing found 0 tables. Please check the file.',
        'footer_credit': 'Created by',
        'dashboard_info': 'Select rows from the table below to dynamically \
generate charts based on your selection.',
        'plot_selection_title': 'Plot for Selected Data',
        'plot_all_title': 'Plot for All Data (No Rows Selected)',
        # NEW STATS
        # TRANSLATIONS
        'stat_metric':
            'Metric',
        'stat_value':
            'Value',
        'stat_count':
            'Count',
        'stat_mean':
            'Average',
        'stat_median':
            'Median',
        'stat_max':
            'Max',
        'stat_min':
            'Min',
        'stat_std':
            'Std. Dev.',
        # NEW INSIGHTS
        # TRANSLATIONS
        'insight_total_revenue': 'Total Revenue',
        'insight_total_discounts': 'Total Discounts',
        'insight_total_tax': 'Total Tax',
        'insight_total_qty': 'Total Quantity',
        'insight_top_branch': 'Top Branch',
        'insight_top_salesman': 'Top Salesman',
        'insight_top_product': 'Top Product',
        # *** NEW V3.2 TRANSLATIONS ***
        'forecast_frequency': 'Forecast Frequency',
        'freq_D': 'Daily',
        'freq_W': 'Weekly',
        'freq_M': 'Monthly',
    },
    'ar': {
        'title': 'تحليلات المبيعات والتنبؤ الاحترافي',
        'upload': 'رفع البيانات (Excel, CSV, PDF, \
HTML)',
        'upload_prompt': 'ارفع ملفاً للبدء. الصيغ المدعومة: \
Excel, CSV, PDF, HTML (جداول).',
        'load_sample':
            'تحميل بيانات عينة',
        'data_loaded':
            'تم تحميل',
        'rows': 'صفوف',
        'cols': 'أعمدة',
        'total_everything': 'مجموع كل الأعمدة الرقمية',
        'grand_total':
            'المجموع الكلي',
        'kpi_selection': 'اختر أعمدة المؤشرات (للإجماليات \
والتنبؤ)',
        'date_column':
            'اختر عمود التاريخ (للسلاسل الزمنية)',
        'pivot_config': 'إعداد الجدول المحوري',
        'row_field': 'حقل (حقول) الصف',
        'col_field': 'حقل (حقول) العمود',
        'agg_type': 'نوع التجميع',
        'value_col': 'عمود القيمة',
        'generate_pivot': 'إنشاء جدول محوري',
        'stats_summary': 'ملخص الإحصائيات',
        'charts': 'المخططات والمرئيات',
        'chart_type':
            'نوع المخطط',
        'x_axis': 'المحور السيني',
        'y_axis': 'المحور الصادي (اختيار متعدد)',
        'plot': 'ارسم المخطط',
        'forecasting':
            'التنبؤ البسيط (الاتجاه)',
        'forecast_column': 'اختر العمود الرقمي للتنبؤ',
        'forecast_periods': 'فترات التنبؤ (خطوات)',
        'run_forecast': 'تشغيل التنبؤ',
        'insights': 'رؤى تلقائية',
        'missing_values': 'القيم المفقودة حسب العمود',
        'correlations': 'مصفوفة الارتباط (رقمي)',
        'download_excel': 'تحميل الملخص كملف \
Excel',
        'download_html': 'تحميل التقرير كملف \
HTML',
        'download_pdf': 'تحميل التقرير كملف \
PDF',
        'language': 'اللغة',
        'theme': 'الوضع الداكن',
        'show_data': 'عرض البيانات الخام',
        'download_pivot': 'تحميل الجدول المحوري كـ \
Excel',
        'config': 'تكوين الأعمدة',
        'kpi_tab': 'المؤشرات والإحصائيات',
        'dashboard_tab': 'لوحة تحكم تفاعلية',
        'pivot_tab': 'الجدول المحوري',
        'charts_tab':
            'مخططات يدوية',
        'forecast_tab': 'التنبؤ',
        'insights_tab': 'تحليلات البيانات',
        'export_tab':
            'تصدير التقرير',
        'selected_kpis': 'إجماليات المؤشرات المحددة',
        'no_kpis_selected': 'لم يتم تحديد أعمدة مؤشرات.',
        'no_numeric_stats': 'لا توجد أعمدة رقمية للإحصاءات.',
        'plot_warn': 'يرجى اختيار عمود واحد على الأقل للمحور الصادي.',
        'forecast_warn': 'يرجى اختيار عمود رقمي للتنبؤ.',
        'forecast_no_date': 'لم يتم تحديد عمود تاريخ. سيتم التنبؤ \
بناءً على تسلسل البيانات.',
        'forecast_no_data': 'لا توجد بيانات كافية للتنبؤ (تحتاج \
نقطتي بيانات على الأقل).',
        'forecast_fail': 'فشل التنبؤ',
        'forecast_table': 'جدول التنبؤ',
        'actual': 'الفعلي',
        'forecast': 'التنبؤ',
        'confidence':
            'نطاق الثقة',
        'no_corr': 'لا توجد أعمدة رقمية كافية للارتباط.',
        'file_error':
            'لا يمكن قراءة الملف. يرجى التأكد من أن الصيغة \
صحيحة.',
        'pdf_warn': 'لم يتم العثور على جداول في ملف \
PDF. يرجى فحص الملف.',
        'html_warn': 'لم يتم العثور على جداول في ملف \
HTML. يرجى فحص الملف.',
        'footer_credit': 'إعداد',
        'dashboard_info': 'اختر صفوفاً من الجدول أدناه لإنشاء \
مخططات ديناميكياً بناءً على اختيارك.',
        'plot_selection_title': 'مخطط للبيانات \
المحددة',
        'plot_all_title': 'مخطط لكل البيانات (لم يتم تحديد صفوف)',
        # NEW STATS
        # TRANSLATIONS
        'stat_metric':
            'المقياس',
        'stat_value':
            'القيمة',
        'stat_count':
            'العدد',
        'stat_mean': 'المتوسط',
        'stat_median':
            'الوسيط',
        'stat_max': 'الأعلى',
        'stat_min': 'الأدنى',
        'stat_std': 'الانحراف المعياري',
        # NEW INSIGHTS
        # TRANSLATIONS
        'insight_total_revenue': 'إجمالي \
الإيرادات',
        'insight_total_discounts': 'إجمالي الخصومات',
        'insight_total_tax': 'إجمالي الضريبة',
        'insight_total_qty': 'إجمالي الكمية',
        'insight_top_branch': 'أفضل فرع',
        'insight_top_salesman': 'أفضل بائع',
        'insight_top_product': 'أفضل منتج',
        # *** NEW V3.2 TRANSLATIONS ***
        'forecast_frequency': 'تردد التنبؤ',
        'freq_D': 'يومي',
        'freq_W': 'أسبوعي',
        'freq_M': 'شهري',
    }
}

def t(key: str) -> str:
    """
    Translation helper
    function.
    Fetches a
    translation string based on the current language in session state.
    """
    lang =
    st.session_state.get('lang', 'en')
    return
    TRANSLATIONS.get(lang, TRANSLATIONS['en']).get(key, key)

# ================================================
# 3. DATA LOADING & PARSING HELPERS (WITH CACHING)
# ================================================

@st.cache_data
def parse_pdf(file_content: bytes) ->
Optional[pd.DataFrame]:
    """Extract tables from a PDF file."""
    all_tables = []
    try:
        with
        io.BytesIO(file_content) as f:
            with
            pdfplumber.open(f) as pdf:
                for
                page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            all_tables.append(pd.DataFrame(table[1:], columns=table[0]))
    except Exception
    as e:
        st.error(f"Error reading PDF: {e}")
        return None
    
    if not all_tables:
        st.warning(t('pdf_warn'))
        return None
    
    df =
    pd.concat(all_tables, ignore_index=True)
    return df

@st.cache_data
def parse_html(file_content: bytes) ->
Optional[pd.DataFrame]:
    """Extract tables from an HTML file."""
    try:
        tables =
        pd.read_html(io.BytesIO(file_content), encoding='utf-8')
        if not tables:
            st.warning(t('html_warn'))
            return
        None
        
        df =
        pd.concat(tables, ignore_index=True)
        return df
    except Exception
    as e:
        st.error(f"{t('file_error')}: {e}")
        return None

@st.cache_data
def parse_excel_csv(file_content: bytes, file_name: str)
-> Optional[pd.DataFrame]:
    """Read and clean Excel/CSV files with smart header
detection."""
    name =
    file_name.lower()
    df = None
    file_like_object =
    io.BytesIO(file_content)
    
    try:
        if
        name.endswith('.csv'):
            df =
            pd.read_csv(file_like_object, header=None, encoding='utf-8', engine='python')
        else:
            df =
            pd.read_excel(file_like_object, header=None, engine='openpyxl')
    except Exception
    as e:
        st.error(f"{t('file_error')}: {e}")
        return None

    # Drop completely
    # empty rows and columns
    df =
    df.dropna(how='all').dropna(axis=1, how='all')
    if df.empty:
        return None

    # Detect header
    # row: pick the row with the most non-null values
    header_row =
    df.notna().sum(axis=1).idxmax()
    df.columns =
    df.iloc[header_row].astype(str).str.strip()
    df =
    df.iloc[header_row + 1:].reset_index(drop=True)

    # Clean column
    # names: replace Unnamed or blanks
    df.columns = [
        col if
        (isinstance(col, str) and col.strip() != "" and not
        col.strip().startswith("Unnamed"))
        else
        f"Column_{i}"
        for i, col in
        enumerate(df.columns)
    ]

    df =
    df.dropna(how="all").reset_index(drop=True)

    # Try converting
    # numeric columns
    for c in
    df.columns:
        df[c] =
        pd.to_numeric(df[c], errors='ignore')

    # Drop duplicated
    # columns
    df = df.loc[:,
        ~df.columns.duplicated()]
    return df

def load_data(uploaded_file: BinaryIO):
    """
    Master function to
    load data from any supported file type.
    This function
    handles the file I/O and session state logic,
    while calling
    cached functions for the actual parsing.
    """
    if uploaded_file
    is None:
        return

    name =
    uploaded_file.name
    file_content =
    uploaded_file.getvalue()
    df = None

    try:
        if
        name.lower().endswith('.pdf'):
            df =
            parse_pdf(file_content)
        elif
        name.lower().endswith(('.html', '.htm')):
            df =
            parse_html(file_content)
        elif
        name.lower().endswith(('.csv', '.xls', '.xlsx')):
            df =
            parse_excel_csv(file_content, name)
        else:
            st.error(f"Unsupported file type: {name}")
            return

        if df is not
        None and not df.empty:
            #
            # Post-processing for all loaded data
            df =
            df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
            for c in
            df.columns:
                df[c]
                = pd.to_numeric(df[c], errors='ignore')
            
            st.session_state['df'] = df
            st.session_state['file_name'] = uploaded_file.name
            # Reset pivot table when new data is loaded
            st.session_state['pivot_df'] = None
            st.success(f"{t('data_loaded')} '{uploaded_file.name}' \
({df.shape[0]} {t('rows')}, {df.shape[1]} {t('cols')})")
        elif df is
        None:
            # Error
            # was already shown by the parsing function
            st.session_state['df'] = None
            st.session_state['file_name'] = None
        elif df is not
        None and df.empty:
            # Warning
            # was already shown by the parsing function
            st.session_state['df'] = None
            st.session_state['file_name'] = None

    except Exception
    as e:
        st.error(f"{t('file_error')}: {e}")
        st.session_state['df'] = None
        st.session_state['file_name'] = None

@st.cache_data
def get_sample_data() -> pd.DataFrame:
    """Generates sample data."""
    df =
    pd.DataFrame({
        'Date':
            pd.date_range(end=pd.Timestamp.today(), periods=24, freq='MS'),
        'Category':
            ['A', 'B', 'C'] * 8,
        'Branch':
            ['North', 'South'] * 12,
        'Sales':
            np.random.randint(100, 1000, 24),
        'Quantity':
            np.random.randint(1, 50, 24),
        'Profit':
            np.random.randint(-50, 300, 24)
    })
    return df

def load_sample_data():
    """Loads sample data into session
state."""
    df =
    get_sample_data()
    st.session_state['df'] = df
    st.session_state['file_name'] = 'Sample_Data.csv'
    # Default KPI and Date columns for sample data
    st.session_state['kpi_cols'] = ['Sales', 'Profit', 'Quantity']
    st.session_state['date_col'] = 'Date'
    st.session_state['pivot_df'] = None
    st.success(f"{t('data_loaded')} 'Sample_Data.csv' ({df.shape[0]} \
{t('rows')}, {df.shape[1]} {t('cols')})")

# ================================================
# 4. ANALYSIS & PLOTTING HELPERS (WITH CACHING)
# ================================================

@st.cache_data
def grand_totals(df: pd.DataFrame) -> Tuple[Dict[str,
                                                float], float]:
    """Calculates totals for all numeric
columns."""
    numeric =
    df.select_dtypes(include=[np.number])
    totals =
    numeric.sum(numeric_only=True)
    grand =
    totals.sum()
    return
    totals.to_dict(), grand

@st.cache_data
def stats_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Generates descriptive statistics."""
    numeric =
    df.select_dtypes(include=[np.number])
    if numeric.empty:
        return
    pd.DataFrame()
    summary =
    numeric.agg(['count', 'mean', 'median', 'max', 'min', 'std']).transpose()
    return summary

@st.cache_data
def generate_pivot(df: pd.DataFrame, rows: List[str], cols:
List[str], values: Optional[str], aggfunc: str) -> Optional[pd.DataFrame]:
    """Generates a pivot table."""
    agg_map = {
        'sum': np.sum,
        'mean': np.mean, 'median': np.median,
        'count':
            'count', 'min': np.min, 'max': np.max, 'std': np.std,
    }
    func =
    agg_map.get(aggfunc, np.sum)
    try:
        pvt =
        pd.pivot_table(df, index=rows if rows else None, 
                       columns=cols if cols else None,
                       values=values if values else None, 
                       aggfunc=func, margins=True, fill_value=0)
        return pvt
    except Exception
    as e:
        st.error(f"Pivot error: {e}")
        return None

def run_forecast(df: pd.DataFrame, date_col: Optional[str],
                 fc_col: str, fc_periods: int, fc_frequency: str = 'D'):
    """
    Runs and plots a
    simple polynomial forecast.
    Not cached as it's
    a quick calculation and should respond to UI changes.
    
    *** UPDATED in V3.2 to include fc_frequency ***
    """
    if not fc_col:
        st.warning(t('forecast_warn'))
        return

    try:
        if date_col:
            # ---
            # Forecasting with a Date Column ---
            tmp =
            df[[date_col, fc_col]].copy()
            tmp[date_col] = pd.to_datetime(tmp[date_col], errors='coerce')
            tmp =
            tmp.dropna(subset=[date_col, fc_col])
            tmp =
            tmp.groupby(date_col, as_index=False)[fc_col].mean().sort_values(date_col)
            tmp_series
            = tmp.set_index(date_col)[fc_col]
            tmp_series
            = tmp_series[~tmp_series.index.duplicated(keep='first')]
            
            # *** NEW V3.2 LOGIC: Resample data based on user choice ***
            if fc_frequency == 'W':
                tmp_series = tmp_series.resample('W').mean()
                freq = 'W'
            elif fc_frequency == 'M':
                tmp_series = tmp_series.resample('MS').mean() # 'MS' = Month Start
                freq = 'MS'
            else: # 'D' or other
                # Try to infer daily frequency if not explicitly daily
                freq = pd.infer_freq(tmp_series.index)
                if freq is None: freq = 'D' # Default to Day
            
            # Interpolate missing values created by resampling
            tmp_series = tmp_series.interpolate(method='linear').fillna(method='bfill').fillna(method='ffill')
            # *** END NEW V3.2 LOGIC ***
            
            # UPDATED:
            # Allow forecast for 2 points (for a straight line)
            if
            tmp_series.shape[0] < 2:
                st.warning(t('forecast_no_data'))
                return

            n =
            tmp_series.shape[0]
            deg = 1 #
            # Always use degree 1 (straight line) if n < 6
            if n >=
            6:
                deg =
                2 # Use degree 2 (curve) if 6 or more points
                
            x =
            np.arange(n)
            coeffs =
            np.polyfit(x, tmp_series.values, deg)
            model =
            np.poly1d(coeffs)

            fitted =
            model(x)
            resid =
            tmp_series.values - fitted
            resid_std
            = np.nanstd(resid)
            ci = 1.96
            * resid_std
            
            last_date
            = tmp_series.index.max()
            # The 'freq' variable is now set by the new resampling logic
            future_index = pd.date_range(start=last_date, periods=int(fc_periods) +
                                         1, freq=freq)[1:]

            future_x =
            np.arange(n, n + int(fc_periods))
            preds =
            model(future_x)
            
            forecast_df = pd.DataFrame({
                date_col: future_index,
                'forecast': preds,
                'lower_band': preds - ci,
                'upper_band': preds + ci
            })
            
            fig =
            go.Figure()
            fig.add_trace(go.Scatter(x=tmp_series.index, y=tmp_series.values,
                                     mode='lines', name=t('actual'), line=dict(color='blue')))
            fig.add_trace(go.Scatter(x=forecast_df[date_col],
                                     y=forecast_df['forecast'],
                                     mode='lines', name=t('forecast'), line=dict(dash='dash', color='red',
                                                                                 width=3)))
            fig.add_trace(go.Scatter(
                x=list(forecast_df[date_col]) + list(forecast_df[date_col][::-1]),
                y=list(forecast_df['upper_band']) +
                list(forecast_df['lower_band'][::-1]),
                fill='toself', fillcolor='rgba(255,0,0,0.15)',
                line=dict(color='rgba(255,255,255,0)'),
                hoverinfo="skip", showlegend=True, name=t('confidence')
            ))
            fig.update_layout(title=f"{fc_col} - {t('forecast')} ({t(f'freq_{fc_frequency}')})",
                              xaxis_title=date_col, yaxis_title=fc_col)
            st.plotly_chart(fig, use_container_width=True)
            st.subheader(t('forecast_table'))
            st.dataframe(forecast_df.reset_index(drop=True))

        else:
            # --- No
            # date column: forecast on index ---
            # (This section remains unchanged as frequency selection
            # doesn't apply)
            st.info(t('forecast_no_date'))
            series =
            df[fc_col].dropna().astype(float)
            # UPDATED:
            # Allow forecast for 2 points (for a straight line)
            if
            series.shape[0] < 2:
                st.warning(t('forecast_no_data'))
                return

            n =
            series.shape[0]
            deg = 1 #
            # Always use degree 1 (straight line) if n < 6
            if n >=
            6:
                deg =
                2 # Use degree 2 (curve) if 6 or more points
                
            x =
            np.arange(n)
            coeffs =
            np.polyfit(x, series.values, deg)
            model =
            np.poly1d(coeffs)
            
            fitted =
            model(x)
            resid =
            series.values - fitted
            resid_std
            = np.nanstd(resid)
            ci = 1.96
            * resid_std
            
            future_x =
            np.arange(n, n + int(fc_periods))
            preds =
            model(future_x)
            
            forecast_df = pd.DataFrame({
                'index': future_x,
                'forecast': preds,
                'lower_band': preds - ci,
                'upper_band': preds + ci
            })

            fig =
            go.Figure()
            fig.add_trace(go.Scatter(x=x, y=series.values, mode='lines',
                                     name=t('actual')))
            fig.add_trace(go.Scatter(x=future_x, y=preds, mode='lines',
                                     name=t('forecast'), line=dict(dash='dash', color='red', width=3)))
            fig.add_trace(go.Scatter(
                x=list(future_x) + list(future_x[::-1]),
                y=list(preds + ci) + list(preds - ci)[::-1],
                fill='toself', fillcolor='rgba(255,0,0,0.15)',
                line=dict(color='rgba(255,255,255,0)'),
                hoverinfo="skip", showlegend=True, name=t('confidence')
            ))
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(forecast_df)

    except Exception
    as e:
        st.error(f"{t('forecast_fail')}: {e}")

# ================================================
# 5. EXPORTING HELPERS
# ================================================

def df_to_excel_bytes(sheets: Dict[str, pd.DataFrame]) ->
bytes:
    """Converts a dictionary of DataFrames to an Excel file
in bytes."""
    out = io.BytesIO()
    with
    pd.ExcelWriter(out, engine='openpyxl') as writer:
        for name,
        df_sheet in sheets.items():
            if not
            isinstance(df_sheet, pd.DataFrame):
                continue
            safe_name
            = str(name)[:31]  # Excel sheet name
            # limit
            df_sheet.to_excel(writer, sheet_name=safe_name,
                              index=isinstance(df_sheet.index, pd.MultiIndex) or df_sheet.index.name is not None)
    out.seek(0)
    return
    out.getvalue()

def create_html_report(df: pd.DataFrame, insights:
List[str]) -> bytes:
    """Generates a simple HTML report."""
    html =
    f'<html><head><meta \
charset="utf-8"><title>{t("title")}</title></head><body>'
    html +=
    f'<h1>{t("title")}</h1>'
    html +=
    f'<p>Generated: {datetime.now().strftime("%Y-%m-%d \
%H:%M:%S")}</p>'
    html +=
    f'<h2>Dataset</h2><p>{t("rows")}: {df.shape[0]} | \
{t("cols")}: {df.shape[1]}</p>'
    html +=
    f'<h3>{t("insights")}</h3><ul>'
    for ins in
    insights:
        html +=
        f'<li>{ins}</li>'
    html +=
    '</ul>'
    html +=
    f'<h3>{t("show_data")}</h3>'
    html +=
    df.head(100).to_html(classes='table', border=1, justify='center')
    html +=
    '</body></html>'
    return
    html.encode('utf-8')

def generate_pdf_report(df: pd.DataFrame, stats:
pd.DataFrame, insights: List[str]) -> bytes:
    """Generates a professional PDF report with
tables."""
    buffer =
    io.BytesIO()
    doc =
    SimpleDocTemplate(buffer, pagesize=A4, rightMargin=72, leftMargin=72,
                      topMargin=72, bottomMargin=18)
    styles =
    getSampleStyleSheet()
    story = []

    # Title
    story.append(Paragraph(t('title'), styles['h1']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Report Generated: \
{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 24))

    # Insights
    story.append(Paragraph(t('insights'), styles['h2']))
    for ins in
    insights:
        story.append(Paragraph(f"• {ins}", styles['Normal']))
    story.append(Spacer(1, 24))

    # Statistics
    if not
    stats.empty:
        story.append(Paragraph(t('stats_summary'), styles['h2']))
        # UPDATED: Use
        # translated key for the index column
        stats_df_reset
        = stats.reset_index().rename(columns={'index': t('stat_metric')})
        stats_data =
        [stats_df_reset.columns.to_list()] + stats_df_reset.values.tolist()
        
        # Format
        # numbers in data
        for i in
        range(1, len(stats_data)):
            for j in
            range(1, len(stats_data[i])):
                try:
                    stats_data[i][j] = f"{stats_data[i][j]:.2f}"
                except
                (TypeError, ValueError):
                    pass
        
        t_style =
        TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN',
            (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID',
            (0, 0), (-1, -1), 1, colors.black)
        ])
        
        stats_table =
        Table(stats_data, colWidths=[1.5*inch] +
              [0.8*inch]*(len(stats_df_reset.columns)-1))
        stats_table.setStyle(t_style)
        story.append(stats_table)
        story.append(Spacer(1, 24))

    # Raw Data
    # (Preview)
    story.append(Paragraph(t('show_data') + " (Top 50 rows)",
                           styles['h2']))
    
    # Truncate data if
    # too wide
    max_cols = 8
    df_preview =
    df.head(50)
    if
    df_preview.shape[1] > max_cols:
        df_preview =
        df_preview.iloc[:, :max_cols]
        story.append(Paragraph(f"(Showing first {max_cols} columns)",
                               styles['Italic']))

    data =
    [df_preview.columns.to_list()] + df_preview.astype(str).values.tolist()
    
    t_style_data =
    TableStyle([
        ('BACKGROUND',
        (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR',
        (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0,
                  0), (-1, -1), 'LEFT'),
        ('FONTNAME',
        (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND',
        (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0,
                   0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE',
        (0, 0), (-1, -1), 7),
    ])
    
    data_table =
    Table(data)
    data_table.setStyle(t_style_data)
    story.append(data_table)

    doc.build(story)
    buffer.seek(0)
    return
    buffer.getvalue()

# ================================================
# 6. AUTOMATED INSIGHTS FUNCTION (WITH CACHING)
# ================================================

@st.cache_data
def get_automated_insights(df: pd.DataFrame) -> \
Tuple[List[Tuple[str, str, str]], Dict[str, str], Optional[str], \
      Optional[str]]:
    """Generates a list of textual insights based on column
names."""
    # UPDATED:
    # Insights is now a list of tuples (emoji, key, value)
    insights: \
    List[Tuple[str, str, str]] = []
    insights_dict = {}

    def safe_find(df: \
                  pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """Finds a column name matching possible names (case-insensitive, trimmed)."""
        for name in \
        possible_names:
            for col in \
            df.columns:
                if \
                str(col).strip().lower() == str(name).strip().lower():
                    return col
        return None

    # 1. Find key numeric columns
    rev_col = safe_find(df, ['sales', 'revenue', 'amount', 'total', 'price'])
    discount_col = safe_find(df, ['discount', 'rebate'])
    tax_col = safe_find(df, ['tax', 'vat'])
    qty_col = safe_find(df, ['quantity', 'qty', 'units'])

    # 2. Find key categorical columns for top metrics
    branch_col = safe_find(df, ['branch', 'store', 'location'])
    salesman_col = safe_find(df, ['salesman', 'agent', 'employee'])
    product_col = safe_find(df, ['product', 'item', 'sku'])

    # 3. Calculate Insights
    if rev_col and df[rev_col].dtype in [np.int64, np.float64, np.int32, np.float32]:
        total_rev = df[rev_col].sum()
        insights.append(('💰', t('insight_total_revenue'), f"{total_rev:,.2f}"))
        insights_dict['Total Revenue'] = f"{total_rev:,.2f}"

    if discount_col and df[discount_col].dtype in [np.int64, np.float64, np.int32, np.float32]:
        total_disc = df[discount_col].sum()
        insights.append(('📉', t('insight_total_discounts'), f"{total_disc:,.2f}"))
        insights_dict['Total Discounts'] = f"{total_disc:,.2f}"

    if tax_col and df[tax_col].dtype in [np.int64, np.float64, np.int32, np.float32]:
        total_tax = df[tax_col].sum()
        insights.append(('🏛️', t('insight_total_tax'), f"{total_tax:,.2f}"))
        insights_dict['Total Tax'] = f"{total_tax:,.2f}"

    if qty_col and df[qty_col].dtype in [np.int64, np.float64, np.int32, np.float32]:
        total_qty = df[qty_col].sum()
        insights.append(('📦', t('insight_total_qty'), f"{total_qty:,.0f}"))
        insights_dict['Total Quantity'] = f"{total_qty:,.0f}"

    if branch_col and rev_col and not df[branch_col].empty:
        try:
            top_branch = df.groupby(branch_col)[rev_col].sum().sort_values(ascending=False).index[0]
            insights.append(('🏢', t('insight_top_branch'), str(top_branch)))
            insights_dict['Top Branch'] = str(top_branch)
        except:
             pass

    if salesman_col and rev_col and not df[salesman_col].empty:
        try:
            top_salesman = df.groupby(salesman_col)[rev_col].sum().sort_values(ascending=False).index[0]
            insights.append(('👨‍💼', t('insight_top_salesman'), str(top_salesman)))
            insights_dict['Top Salesman'] = str(top_salesman)
        except:
            pass

    if product_col and qty_col and not df[product_col].empty:
        try:
            top_product = df.groupby(product_col)[qty_col].sum().sort_values(ascending=False).index[0]
            insights.append(('🏷️', t('insight_top_product'), str(top_product)))
            insights_dict['Top Product'] = str(top_product)
        except:
            pass

    return insights, insights_dict, rev_col, qty_col

# ================================================
# 7. MAIN APPLICATION LOGIC
# ================================================

def main():
    st.title(t('title'))

    # --- SIDEBAR CONFIGURATION ---
    # Language Selector
    lang_choice = st.sidebar.radio(t('language'), options=['en', 'ar'],
                                    index=0 if st.session_state['lang'] == 'en' else 1,
                                    format_func=lambda x: 'English' if x == 'en' else 'العربية',
                                    key='lang_radio')
    st.session_state['lang'] = lang_choice

    # Dark Mode switch (simple hack for Streamlit theme control)
    st.sidebar.markdown("---")
    if st.sidebar.checkbox(t('theme'), value=True):
        st.markdown(
            """<style>
            .main { background-color: #0E1117; color: white; }
            .st-bd { color: #AFAFAF; }
            </style>""",
            unsafe_allow_html=True
        )

    # File Uploader
    st.sidebar.title(t('upload'))
    uploaded_file = st.sidebar.file_uploader(t('upload_prompt'),
                                             type=['csv', 'xlsx', 'xls', 'pdf', 'html', 'htm'],
                                             key="file_uploader")

    # Handle file upload change
    if uploaded_file and uploaded_file.name != st.session_state.get('last_uploaded_file'):
         load_data(uploaded_file)
         st.session_state['last_uploaded_file'] = uploaded_file.name

    if st.sidebar.button(t('load_sample')):
        load_sample_data()

    df = st.session_state['df']

    # --- DATA CONFIGURATION ---
    if df is not None:
        all_cols = df.columns.to_list()
        numeric_cols = df.select_dtypes(include=[np.number]).columns.to_list()
        
        st.sidebar.markdown("---")
        st.sidebar.header(t('config'))
        
        # KPI Selection
        st.session_state['kpi_cols'] = st.sidebar.multiselect(
            t('kpi_selection'),
            options=numeric_cols,
            default=st.session_state.get('kpi_cols', [c for c in ['Sales', 'Profit', 'Quantity'] if c in numeric_cols])
        )

        # Date Column Selection
        date_cols = [c for c in all_cols if df[c].dtype == 'datetime64[ns]' or pd.to_datetime(df[c], errors='coerce').notna().any()]
        
        st.session_state['date_col'] = st.sidebar.selectbox(
            t('date_column'),
            options=['None'] + date_cols,
            index=0 if st.session_state.get('date_col') == 'None' or st.session_state.get('date_col') is None else (date_cols.index(st.session_state['date_col']) + 1) if st.session_state['date_col'] in date_cols else 0,
            key='date_selector'
        )
        
        if st.session_state['date_col'] != 'None':
            # Ensure the selected date column is indeed datetime type
            try:
                df[st.session_state['date_col']] = pd.to_datetime(df[st.session_state['date_col']], errors='coerce')
            except:
                st.warning("Selected date column could not be converted to date format.")
                st.session_state['date_col'] = 'None'


    # --- MAIN DASHBOARD TABS ---
    if df is not None:
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            t('kpi_tab'), t('dashboard_tab'), t('pivot_tab'), t('charts_tab'), 
            t('forecast_tab'), t('insights_tab'), t('export_tab')
        ])

        # === TAB 1: KPIs & Stats ===
        with tab1:
            st.header(t('selected_kpis'))
            kpi_cols = st.session_state['kpi_cols']
            if kpi_cols:
                totals = df[kpi_cols].sum(numeric_only=True)
                cols = st.columns(len(kpi_cols))
                for i, col in enumerate(kpi_cols):
                    with cols[i]:
                        st.metric(label=col, value=f"{totals[col]:,.2f}")
                
                st.markdown("---")
            else:
                st.info(t('no_kpis_selected'))

            st.header(t('stats_summary'))
            summary_df = stats_summary(df)
            if not summary_df.empty:
                # Rename columns using translations
                summary_df = summary_df.rename(columns={
                    'count': t('stat_count'), 'mean': t('stat_mean'), 
                    'median': t('stat_median'), 'max': t('stat_max'),
                    'min': t('stat_min'), 'std': t('stat_std')
                })
                st.dataframe(summary_df)
            else:
                st.info(t('no_numeric_stats'))

        # === TAB 2: Interactive Dashboard ===
        with tab2:
            st.header(t('dashboard_tab'))
            st.info(t('dashboard_info'))

            # Display editable/selectable dataframe
            edited_df = st.data_editor(
                df,
                key="dashboard_data_editor",
                use_container_width=True,
                num_rows="dynamic"
            )
            
            # Extract selected rows (Streamlit Data Editor hack for selected rows)
            try:
                # This complex logic is used to correctly identify selected rows from st.data_editor
                selection_indices = st.session_state["dashboard_data_editor"]["selection"]["rows"]
                selected_rows = df.iloc[selection_indices]
            except (KeyError, IndexError):
                # If no rows are selected, selection key might not exist or be empty
                selected_rows = pd.DataFrame() 

            
            plot_df = selected_rows if not selected_rows.empty else edited_df
            plot_title = t('plot_selection_title') if not selected_rows.empty else t('plot_all_title')
            
            st.subheader(plot_title)
            
            # Dashboard plot configuration (pre-defined simple scatter)
            numeric_cols_dashboard = plot_df.select_dtypes(include=[np.number]).columns.to_list()
            
            if len(numeric_cols_dashboard) >= 2:
                x_col = numeric_cols_dashboard[0]
                y_col = numeric_cols_dashboard[1]
                
                cat_cols_dashboard = plot_df.select_dtypes(exclude=[np.number, np.datetime64]).columns.to_list()
                
                fig_dash = px.scatter(
                    plot_df, 
                    x=x_col, 
                    y=y_col, 
                    color=cat_cols_dashboard[0] if cat_cols_dashboard else None,
                    title=f"{x_col} vs {y_col}"
                )
                st.plotly_chart(fig_dash, use_container_width=True)
            elif not numeric_cols_dashboard:
                st.warning(t('no_numeric_stats'))
            else:
                st.info("Select more columns or rows to enable scatter plot visualization.")

        # === TAB 3: Pivot Table ===
        with tab3:
            st.header(t('pivot_tab'))
            
            all_cols_pivot = df.columns.to_list()
            numeric_cols_pivot = df.select_dtypes(include=[np.number]).columns.to_list()

            col_pvt1, col_pvt2, col_pvt3 = st.columns(3)
            with col_pvt1:
                rows_select = st.multiselect(t('row_field'), options=all_cols_pivot, key='pivot_rows')
            with col_pvt2:
                cols_select = st.multiselect(t('col_field'), options=all_cols_pivot, key='pivot_cols')
            with col_pvt3:
                value_select = st.selectbox(t('value_col'), options=['None'] + numeric_cols_pivot, key='pivot_values')
                agg_select = st.selectbox(t('agg_type'), options=['sum', 'mean', 'count', 'max', 'min', 'std'], key='pivot_agg')
                
            if st.button(t('generate_pivot')):
                if value_select != 'None':
                    pvt_df = generate_pivot(df, rows_select, cols_select, value_select, agg_select)
                    if pvt_df is not None:
                        st.session_state['pivot_df'] = pvt_df
                else:
                    st.warning("Please select a value column.")

            if st.session_state.get('pivot_df') is not None:
                st.dataframe(st.session_state['pivot_df'])
                
                # Download button for Pivot
                excel_bytes_pvt = df_to_excel_bytes({'PivotTable': st.session_state['pivot_df']})
                st.download_button(
                    label=f"⬇️ {t('download_pivot')}",
                    data=excel_bytes_pvt,
                    file_name=f"{st.session_state['file_name'].split('.')[0]}_Pivot.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        # === TAB 4: Manual Charts ===
        with tab4:
            st.header(t('charts'))
            
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                chart_type = st.selectbox(t('chart_type'), 
                                          options=['line', 'bar', 'scatter', 'hist', 'box', 'area'], 
                                          key='chart_type_select')
            
            all_cols_chart = df.columns.to_list()
            numeric_cols_chart = df.select_dtypes(include=[np.number]).columns.to_list()

            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                x_axis = st.selectbox(t('x_axis'), options=['None'] + all_cols_chart, key='x_axis_select')
            with col_chart2:
                y_axis = st.multiselect(t('y_axis'), options=numeric_cols_chart, key='y_axis_select')
            
            if st.button(t('plot'), key='plot_button'):
                if not y_axis:
                    st.warning(t('plot_warn'))
                else:
                    try:
                        # Handle multi-Y axis plotting using melted dataframe for Line, Bar, Area, Scatter
                        if chart_type in ['line', 'bar', 'area', 'scatter']:
                            if x_axis == 'None':
                                plot_df_melt = df[y_axis].reset_index().melt(id_vars='index', value_vars=y_axis, var_name='Metric', value_name='Value')
                                fig = px.line(plot_df_melt, x='index', y='Value', color='Metric', title=f"{chart_type.capitalize()} Chart: {', '.join(y_axis)}")
                            else:
                                plot_df_melt = df[[x_axis] + y_axis].melt(id_vars=x_axis, value_vars=y_axis, var_name='Metric', value_name='Value')
                                
                                fig = go.Figure()
                                for metric in plot_df_melt['Metric'].unique():
                                    subset = plot_df_melt[plot_df_melt['Metric'] == metric]
                                    if chart_type == 'line':
                                        fig.add_trace(go.Scatter(x=subset[x_axis], y=subset['Value'], mode='lines', name=metric))
                                    elif chart_type == 'bar':
                                        fig.add_trace(go.Bar(x=subset[x_axis], y=subset['Value'], name=metric))
                                    elif chart_type == 'area':
                                        fig.add_trace(go.Scatter(x=subset[x_axis], y=subset['Value'], fill='tozeroy', mode='lines', name=metric))
                                    elif chart_type == 'scatter':
                                        fig.add_trace(go.Scatter(x=subset[x_axis], y=subset['Value'], mode='markers', name=metric))

                                fig.update_layout(title=f"{chart_type.capitalize()} Chart: {x_axis} vs {', '.join(y_axis)}", xaxis_title=x_axis, yaxis_title="Value")
                                
                        elif chart_type in ['hist', 'box']:
                            # Histogram/Box plots typically use a single Y-variable
                            fig = getattr(px, chart_type)(df, y=y_axis[0], x=x_axis if x_axis != 'None' else None, title=f"{chart_type.capitalize()} Chart: {y_axis[0]}")

                        st.plotly_chart(fig, use_container_width=True)

                    except Exception as e:
                        st.error(f"Plotting Error: {e}")

        # === TAB 5: Forecasting ===
        with tab5:
            st.header(t('forecasting'))
            
            numeric_fc_cols = df.select_dtypes(include=[np.number]).columns.to_list()
            
            col_fc1, col_fc2, col_fc3 = st.columns(3)
            with col_fc1:
                fc_col = st.selectbox(t('forecast_column'), options=['None'] + numeric_fc_cols, key='fc_col_select')
            with col_fc2:
                fc_periods = st.number_input(t('forecast_periods'), min_value=1, value=5, key='fc_periods_input')
            with col_fc3:
                # NEW V3.2 FEATURE: Frequency Selection
                fc_frequency = st.selectbox(
                    t('forecast_frequency'), 
                    options=['D', 'W', 'M'],
                    format_func=lambda x: t(f'freq_{x}'),
                    key='fc_freq_select',
                    disabled=(st.session_state['date_col'] == 'None')
                )
                
            if st.button(t('run_forecast'), key='run_fc_button'):
                if fc_col != 'None':
                    date_col_for_fc = st.session_state['date_col'] if st.session_state['date_col'] != 'None' else None
                    run_forecast(df, date_col_for_fc, fc_col, fc_periods, fc_frequency)
                else:
                    st.warning(t('forecast_warn'))

        # === TAB 6: Data Insights ===
        with tab6:
            st.header(t('insights'))
            
            # Run the automated insights function
            insights_list, insights_dict, rev_col, qty_col = get_automated_insights(df)
            
            st.subheader("Key Performance Indicators (KPIs)")
            if insights_list:
                cols_ins = st.columns(4)
                for i, (emoji, key, value) in enumerate(insights_list[:4]):
                    with cols_ins[i % 4]:
                        st.metric(label=f"{emoji} {key}", value=value)
                
                st.markdown("---")
            else:
                st.info("No automated insights could be generated. Try renaming columns to 'Sales', 'Quantity', 'Branch', etc.")
            
            st.subheader(t('missing_values'))
            missing_count = df.isnull().sum()
            missing_df = pd.DataFrame({
                'Column': missing_count.index,
                'Missing Count': missing_count.values,
                'Missing Percentage': (missing_count.values / len(df)) * 100
            }).sort_values(by='Missing Count', ascending=False)
            
            st.dataframe(missing_df[missing_df['Missing Count'] > 0])
            
            st.subheader(t('correlations'))
            numeric_for_corr = df.select_dtypes(include=[np.number])
            if numeric_for_corr.shape[1] >= 2:
                corr = numeric_for_corr.corr()
                fig_corr = px.imshow(
                    corr,
                    text_auto=".2f",
                    aspect="auto",
                    title=t('correlations')
                )
                st.plotly_chart(fig_corr, use_container_width=True)
            else:
                st.info(t('no_corr'))

        # === TAB 7: Export Report ===
        with tab7:
            st.header(t('export_tab'))
            
            # Prepare data for export
            sheets_for_excel = {
                'Raw Data (Head)': df.head(100),
                'Stats Summary': stats_summary(df),
            }
            if st.session_state.get('pivot_df') is not None:
                sheets_for_excel['Pivot Table'] = st.session_state['pivot_df']

            # Ensure insights are run for export preparation
            insights_list, _, _, _ = get_automated_insights(df)

            excel_bytes = df_to_excel_bytes(sheets_for_excel)
            html_bytes = create_html_report(df, [f"{key}: {value}" for _, key, value in insights_list])
            pdf_bytes = generate_pdf_report(df, stats_summary(df), [f"{key}: {value}" for _, key, value in insights_list])
            
            col_exp1, col_exp2, col_exp3 = st.columns(3)
            with col_exp1:
                st.download_button(
                    label=f"⬇️ {t('download_excel')}",
                    data=excel_bytes,
                    file_name=f"{st.session_state['file_name'].split('.')[0]}_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_exp2:
                st.download_button(
                    label=f"⬇️ {t('download_html')}",
                    data=html_bytes,
                    file_name=f"{st.session_state['file_name'].split('.')[0]}_Report.html",
                    mime="text/html",
                    use_container_width=True
                )
            with col_exp3:
                st.download_button(
                    label=f"⬇️ {t('download_pdf')}",
                    data=pdf_bytes,
                    file_name=f"{st.session_state['file_name'].split('.')[0]}_Report.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        
        # Show Raw Data toggle (in sidebar but displayed below tabs)
        st.sidebar.markdown("---")
        if st.sidebar.checkbox(t('show_data')):
            st.subheader(t('show_data'))
            st.dataframe(df, use_container_width=True)
    else:
        # Welcome message when no data is loaded
        st.info("Please upload a file or load sample data in the sidebar to begin analysis.")
        
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"*{t('footer_credit')} Gemini*")

if __name__ == '__main__':
    main()
