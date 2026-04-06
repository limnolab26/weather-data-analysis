# app.py — 기상자료 보고서 생성기 Web App v3.0
# 변경사항: 대화형 차트(Plotly), 동적 피벗표, Python 고품질 차트(Matplotlib) 추가
# 실행: streamlit run app.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import seaborn as sns
import io
import os
from datetime import datetime

# 기존 모듈 임포트 (파일이 존재해야 함)
try:
    from data_processor import WeatherDataProcessor
    from excel_generator import generate_excel_report
    from pdf_generator import generate_pdf_report
except ImportError:
    # 모듈이 없을 경우를 대비한 가상 클래스/함수 정의 (테스트용)
    class WeatherDataProcessor:
        def process(self, files): return pd.DataFrame()
    def generate_excel_report(df): return b""
    def generate_pdf_report(df): return b""

# ━━━━━ 환경 설정 및 유틸리티 ━━━━━

def setup_korean_font():
    """Matplotlib 한글 폰트 설정 (시스템 환경별 폴백 적용)"""
    import matplotlib.font_manager as fm
    
    font_list = ['NanumGothic', 'Malgun Gothic', 'AppleGothic', 'DejaVu Sans']
    found = False
    for f in font_list:
        if any(f in font.name for font in fm.font_Manager.ttflist):
            plt.rcParams['font.family'] = f
            found = True
            break
    
    plt.rcParams['axes.unicode_minus'] = False
    return found

def get_chart_bytes(fig) -> bytes:
    """Matplotlib Figure를 PNG 바이트로 변환"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    return buf.read()

def add_season_column(df: pd.DataFrame) -> pd.DataFrame:
    """월 정보를 바탕으로 계절 컬럼 추가"""
    if 'date' in df.columns:
        month = df['date'].dt.month
        conditions = [
            (month.isin([3, 4, 5])),
            (month.isin([6, 7, 8])),
            (month.isin([9, 10, 11])),
            (month.isin([12, 1, 2]))
        ]
        choices = ['봄', '여름', '가을', '겨울']
        df['season'] = np.select(conditions, choices, default='알수없음')
    return df

@st.cache_data
def prepare_chart_data(df: pd.DataFrame, element: str, freq: str) -> pd.DataFrame:
    """차트용 데이터 집계 (freq: 'D', 'ME', 'YE')"""
    temp_df = df.copy()
    # 날짜 인덱스 설정 및 리샘플링
    temp_df = temp_df.set_index('date')
    
    if element == 'temp_group':
        # 기온 그룹(평균, 최고, 최저) 특별 처리
        cols = ['temp_avg', 'temp_max', 'temp_min', 'station_name']
        resampled = temp_df[cols].groupby(['station_name', pd.Grouper(freq=freq)]).mean().reset_index()
    else:
        # 단일 요소 처리
        agg_func = 'sum' if element == 'precipitation' else 'mean'
        resampled = temp_df.groupby(['station_name', pd.Grouper(freq=freq)])[element].agg(agg_func).reset_index()
    
    return resampled

# ━━━━━ 앱 UI 레이아웃 ━━━━━

st.set_page_config(page_title="기상자료 분석 리포터 v3.0", layout="wide")
setup_korean_font()

# 사이드바: 데이터 업로드
with st.sidebar:
    st.header("📁 데이터 업로드")
    uploaded_files = st.file_uploader(
        "기상청 ASOS CSV 파일을 선택하세요", 
        type=['csv'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if 'raw_data' not in st.session_state or st.button("🔄 데이터 새로고침"):
            with st.spinner("데이터를 처리 중입니다..."):
                processor = WeatherDataProcessor()
                # data_processor.py의 로직에 따라 데이터 프레임 생성
                df = processor.process(uploaded_files)
                df = add_season_column(df)
                st.session_state.raw_data = df
                st.success(f"{len(uploaded_files)}개 파일 로드 완료!")
    
    st.divider()
    st.caption("v3.0 - Powered by Streamlit & Plotly")

# 메인 타이틀
st.title("🌡️ 기상자료 분석 및 보고서 생성기")
st.markdown("기상청 ASOS 데이터를 기반으로 대화형 차트와 고품질 보고서를 생성합니다.")

if 'raw_data' not in st.session_state:
    st.info("👈 사이드바에서 기상 자료(CSV)를 먼저 업로드해 주세요.")
    st.stop()

df = st.session_state.raw_data
stations = df['station_name'].unique().tolist()
elements = {
    'temp_avg': '평균기온(°C)',
    'temp_max': '최고기온(°C)',
    'temp_min': '최저기온(°C)',
    'precipitation': '일강수량(mm)',
    'humidity': '평균습도(%)',
    'wind_speed': '평균풍속(m/s)',
    'sunshine': '일조시간(hr)'
}

# 탭 구조 정의
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 대화형 차트", "📋 동적 피벗표", "🐍 Python 차트", "⬇️ 보고서 다운로드", "ℹ️ 사용 방법"
])

# ━━━━━ 탭 1: 대화형 차트 (Plotly) ━━━━━
with tab1:
    st.subheader("대화형 데이터 탐색")
    
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    with c1:
        selected_stations = st.multiselect("관측소 선택", stations, default=stations[:1])
    with c2:
        selected_element = st.selectbox("기상 요소", list(elements.keys()), format_func=lambda x: elements[x])
    with c3:
        freq_opt = st.selectbox("집계 단위", ["일별", "월별", "연별"], index=1)
        freq_map = {"일별": "D", "월별": "ME", "연별": "YE"}
    with c4:
        chart_type = st.radio("차트 유형", ["선형", "막대", "복합(기온+강수)"], horizontal=True)

    plot_df = df[df['station_name'].isin(selected_stations)]
    processed_plot_df = prepare_chart_data(plot_df, selected_element, freq_map[freq_opt])

    if chart_type == "복합(기온+강수)":
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        for s in selected_stations:
            s_data = plot_df[plot_df['station_name'] == s].set_index('date').resample(freq_map[freq_opt]).mean().reset_index()
            fig.add_trace(go.Scatter(x=s_data['date'], y=s_data['temp_avg'], name=f'{s}-기온'), secondary_y=False)
            
            s_rain = plot_df[plot_df['station_name'] == s].set_index('date').resample(freq_map[freq_opt]).sum().reset_index()
            fig.add_trace(go.Bar(x=s_rain['date'], y=s_rain['precipitation'], name=f'{s}-강수량', opacity=0.4), secondary_y=True)
        
        fig.update_yaxes(title_text="기온 (°C)", secondary_y=False)
        fig.update_yaxes(title_text="강수량 (mm)", secondary_y=True)
    else:
        if chart_type == "선형":
            fig = px.line(processed_plot_df, x='date', y=selected_element, color='station_name', markers=True)
        else:
            fig = px.bar(processed_plot_df, x='date', y=selected_element, color='station_name', barmode='group')
    
    fig.update_layout(hovermode="x unified", legend_orientation="h", legend_y=1.1)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True})

# ━━━━━ 탭 2: 동적 피벗표 (Pivot Table) ━━━━━
with tab2:
    st.subheader("데이터 동적 집계 (Pivot Table)")
    
    col_opts = ['year', 'month', 'season', 'station_name']
    df['year'] = df['date'].dt.year
    df['month'] = df['date'].dt.month
    
    pc1, pc2, pc3, pc4 = st.columns(4)
    with pc1: row_sel = st.selectbox("행(Row)", col_opts, index=0)
    with pc2: col_sel = st.selectbox("열(Column)", col_opts, index=1)
    with pc3: val_sel = st.selectbox("값(Value)", list(elements.keys()), index=0, format_func=lambda x: elements[x])
    with pc4: agg_sel = st.selectbox("집계 함수", ['mean', 'sum', 'max', 'min', 'count'], index=0)
    
    if st.button("📊 표 생성하기", use_container_width=True):
        pivot_res = pd.pivot_table(df, values=val_sel, index=row_sel, columns=col_sel, aggfunc=agg_sel)
        st.dataframe(pivot_res.style.format("{:.2f}"), use_container_width=True)
        
        csv = pivot_res.to_csv().encode('utf-8-sig')
        st.download_button("📥 결과 CSV 다운로드", data=csv, file_name=f"pivot_analysis_{datetime.now().strftime('%Y%m%d')}.csv", mime='text/csv')

# ━━━━━ 탭 3: Python 고품질 차트 (Matplotlib/Seaborn) ━━━━━
with tab3:
    st.subheader("출판용 고품질 시각화")
    
    chart_kinds = ["기온 시계열(음영)", "월별 강수량 비교", "월별 기온 분포(Box)", "연도별 히트맵", "기온 vs 강수 산점도"]
    sel_kind = st.selectbox("생성할 차트 종류 선택", chart_kinds)
    
    fig, ax = plt.subplots(figsize=(12, 6), dpi=150)
    sns.set_style("whitegrid")
    
    if sel_kind == "기온 시계열(음영)":
        target_s = st.selectbox("관측소 선택", stations, key="plt_s1")
        pdf = df[df['station_name'] == target_s].sort_values('date')
        ax.plot(pdf['date'], pdf['temp_avg'], color='#E74C3C', label='평균기온', lw=1.5)
        ax.fill_between(pdf['date'], pdf['temp_min'], pdf['temp_max'], alpha=0.2, color='#E74C3C', label='최저/최고 범위')
        ax.set_title(f"[{target_s}] 기온 변화 추이", fontsize=14)
        ax.set_ylabel("기온 (°C)")
        
    elif sel_kind == "월별 강수량 비교":
        pdf = df.groupby(['month', 'station_name'])['precipitation'].sum().reset_index()
        sns.barplot(data=pdf, x='month', y='precipitation', hue='station_name', ax=ax)
        ax.set_title("관측소별 월별 누적 강수량", fontsize=14)
        ax.set_ylabel("강수량 (mm)")

    elif sel_kind == "월별 기온 분포(Box)":
        sns.boxplot(data=df, x='month', y='temp_avg', hue='station_name', ax=ax)
        ax.set_title("월별 평균 기온 분포", fontsize=14)

    elif sel_kind == "연도별 히트맵":
        target_s = st.selectbox("관측소 선택", stations, key="plt_s2")
        h_val = st.selectbox("데이터 선택", ['temp_avg', 'precipitation', 'humidity'], format_func=lambda x: elements[x])
        pdf = df[df['station_name'] == target_s]
        pivot_h = pdf.pivot_table(values=h_val, index=df['date'].dt.year, columns=df['date'].dt.month, aggfunc='mean')
        sns.heatmap(pivot_h, annot=True, fmt='.1f', cmap='RdYlBu_r', ax=ax)
        ax.set_title(f"[{target_s}] 연도/월별 {elements[h_val]} 히트맵", fontsize=14)

    elif sel_kind == "기온 vs 강수 산점도":
        sns.scatterplot(data=df, x='temp_avg', y='precipitation', hue='season', alpha=0.6, ax=ax)
        ax.set_title("기온과 강수량의 상관관계 (계절별)", fontsize=14)

    plt.tight_layout()
    st.pyplot(fig)
    
    st.download_button(
        label="🖼️ 고품질 PNG 다운로드",
        data=get_chart_bytes(fig),
        file_name=f"kma_chart_{datetime.now().strftime('%H%M%S')}.png",
        mime="image/png"
    )

# ━━━━━ 탭 4: 보고서 다운로드 (기존 기능) ━━━━━
with tab4:
    st.subheader("정형 보고서 생성")
    st.info("분석 결과를 엑셀(XLSX) 또는 PDF 형식의 보고서로 다운로드합니다.")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        if st.button("📝 엑셀 보고서 생성", use_container_width=True):
            excel_data = generate_excel_report(df)
            st.download_button("📂 엑셀 파일 받기", excel_data, "기상보고서.xlsx")
            
    with col_dl2:
        if st.button("📄 PDF 보고서 생성", use_container_width=True):
            pdf_data = generate_pdf_report(df)
            st.download_button("📂 PDF 파일 받기", pdf_data, "기상보고서.pdf")

# ━━━━━ 탭 5: 사용 방법 ━━━━━
with tab5:
    st.header("📖 사용 가이드")
    st.markdown("""
    1. **데이터 준비**: 기상자료개방포털(data.kma.go.kr)에서 종관기상관측(ASOS) 일자료를 CSV로 다운로드합니다.
    2. **데이터 업로드**: 왼쪽 사이드바에 CSV 파일들을 드래그앤드롭합니다. (여러 관측소 동시 분석 가능)
    3. **대화형 분석**: '대화형 차트' 탭에서 기상 요소를 선택하고 마우스를 올려 세부 값을 확인하세요.
    4. **피벗 분석**: '동적 피벗표' 탭에서 원하는 기준으로 데이터를 요약 집계하고 CSV로 내보낼 수 있습니다.
    5. **고품질 차트**: 보고서나 발표 자료에 적합한 통계 차트를 생성하고 PNG로 저장하세요.
    """)
    st.warning("주의: 업로드하는 파일은 표준 ASOS 일자료 형식이어야 정상적으로 처리됩니다.")
