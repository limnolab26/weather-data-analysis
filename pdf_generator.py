# pdf_generator.py
# PDF 보고서 생성 모듈
# ReportLab 라이브러리를 사용하여 전문적인 PDF 보고서를 생성합니다.

import pandas as pd
from datetime import datetime
import io

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, HRFlowable)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from data_processor import ELEMENT_LABELS, SUM_ELEMENTS, MODE_ELEMENTS

# ── 한글 폰트 등록 (맑은 고딕이 없을 경우 기본 폰트 사용) ──────────────────
import os, sys

def _register_korean_font():
    """시스템에서 한글 폰트를 찾아 등록"""
    candidates = [
        # Windows
        r'C:\Windows\Fonts\malgun.ttf',
        r'C:\Windows\Fonts\NanumGothic.ttf',
        # macOS
        '/Library/Fonts/NanumGothic.ttf',
        '/System/Library/Fonts/Supplemental/AppleGothic.ttf',
        # Linux
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
        '/usr/share/fonts/truetype/unfonts-core/UnDotum.ttf',
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('KoreanFont', path))
                return 'KoreanFont'
            except Exception:
                continue
    return 'Helvetica'  # 폴백: 영문 기본 폰트

FONT_NAME = _register_korean_font()

# ── 색상 정의 ────────────────────────────────────────────────────────────────
DARK_BLUE   = colors.HexColor('#1F4E79')
MID_BLUE    = colors.HexColor('#2E75B6')
LIGHT_BLUE  = colors.HexColor('#DEEAF1')
ORANGE      = colors.HexColor('#C55A11')
LIGHT_GRAY  = colors.HexColor('#F2F2F2')
WHITE       = colors.white
BLACK       = colors.black


class PDFReportGenerator:
    """PDF 보고서 생성기"""

    def generate(self, df: pd.DataFrame, stats: dict, config: dict,
                 output_path: str, selected_cols: list) -> None:
        """
        PDF 보고서 생성

        Args:
            df            : 필터링된 데이터프레임
            stats         : 통계 딕셔너리
            config        : 보고서 설정
            output_path   : 저장 경로
            selected_cols : 포함할 기상요소 컬럼 목록
        """
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            leftMargin=2*cm, rightMargin=2*cm,
            topMargin=2*cm,  bottomMargin=2*cm,
        )

        styles = self._build_styles()
        story  = []

        # ── 표지 ──
        story += self._cover_page(df, config, styles)
        story.append(PageBreak())

        # ── 관측소별 통계 요약 ──
        for stn, stn_stats in stats.items():
            story += self._station_section(stn, stn_stats, df, styles)
            story.append(Spacer(1, 0.5*cm))

        # ── 월별 상세 통계 ──
        story.append(PageBreak())
        story.append(Paragraph("월별 상세 통계", styles['h1']))
        story.append(Spacer(1, 0.3*cm))

        for stn, stn_stats in stats.items():
            story += self._monthly_tables(stn, stn_stats, styles)

        doc.build(story, onFirstPage=self._page_footer,
                  onLaterPages=self._page_footer)

    # ── 스타일 정의 ──────────────────────────────────────────────────────────
    def _build_styles(self) -> dict:
        base = getSampleStyleSheet()
        return {
            'title': ParagraphStyle(
                'title', fontName=FONT_NAME, fontSize=24,
                textColor=DARK_BLUE, alignment=TA_CENTER,
                spaceAfter=6, leading=30,
            ),
            'subtitle': ParagraphStyle(
                'subtitle', fontName=FONT_NAME, fontSize=11,
                textColor=colors.HexColor('#595959'), alignment=TA_CENTER,
                spaceAfter=4,
            ),
            'h1': ParagraphStyle(
                'h1', fontName=FONT_NAME, fontSize=14,
                textColor=WHITE, backColor=DARK_BLUE,
                spaceAfter=6, spaceBefore=12,
                leftIndent=-0.3*cm, rightIndent=-0.3*cm,
                leading=20, borderPadding=(4, 6, 4, 6),
            ),
            'h2': ParagraphStyle(
                'h2', fontName=FONT_NAME, fontSize=12,
                textColor=WHITE, backColor=MID_BLUE,
                spaceAfter=4, spaceBefore=8,
                leading=18, borderPadding=(3, 4, 3, 4),
            ),
            'body': ParagraphStyle(
                'body', fontName=FONT_NAME, fontSize=10,
                textColor=BLACK, leading=15,
            ),
            'small': ParagraphStyle(
                'small', fontName=FONT_NAME, fontSize=8,
                textColor=colors.HexColor('#595959'),
            ),
            'note': ParagraphStyle(
                'note', fontName=FONT_NAME, fontSize=9,
                textColor=colors.HexColor('#7F7F7F'),
                leftIndent=0.5*cm,
            ),
        }

    # ── 표지 페이지 ──────────────────────────────────────────────────────────
    def _cover_page(self, df, config, styles) -> list:
        elems = [Spacer(1, 3*cm)]

        title = config.get('report_title', '기상 현황 보고서')
        org   = config.get('organization', '')
        s_dt  = df['date'].min().strftime('%Y년 %m월 %d일')
        e_dt  = df['date'].max().strftime('%Y년 %m월 %d일')
        today = datetime.now().strftime('%Y년 %m월 %d일')

        elems.append(Paragraph(title, styles['title']))
        elems.append(Spacer(1, 0.5*cm))
        elems.append(HRFlowable(width='100%', thickness=2, color=DARK_BLUE))
        elems.append(Spacer(1, 1*cm))

        # 정보 테이블
        info_data = [
            ['분석 기간', f'{s_dt}  ~  {e_dt}'],
            ['작성일',   today],
        ]
        if org:
            info_data.insert(0, ['작성 기관', org])

        info_table = Table(info_data, colWidths=[4*cm, 11*cm])
        info_table.setStyle(TableStyle([
            ('FONTNAME',   (0,0), (-1,-1), FONT_NAME),
            ('FONTSIZE',   (0,0), (-1,-1), 11),
            ('TEXTCOLOR',  (0,0), (0,-1),  WHITE),
            ('BACKGROUND', (0,0), (0,-1),  MID_BLUE),
            ('TEXTCOLOR',  (1,0), (1,-1),  BLACK),
            ('BACKGROUND', (1,0), (1,-1),  LIGHT_BLUE),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [None, None]),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
            ('TOPPADDING',    (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        elems.append(info_table)
        elems.append(Spacer(1, 2*cm))

        # 포함 관측소
        if self._stations_str(df):
            elems.append(Paragraph(
                f"포함 관측소: {self._stations_str(df)}", styles['subtitle']))

        return elems

    def _stations_str(self, df) -> str:
        if 'station_name' in df.columns:
            stns = df['station_name'].dropna().unique().tolist()
            return ', '.join(str(s) for s in stns)
        return ''

    # ── 관측소 섹션 ──────────────────────────────────────────────────────────
    def _station_section(self, stn, stn_stats, df, styles) -> list:
        elems = []
        elems.append(Paragraph(f"관측소: {stn}", styles['h1']))

        period = stn_stats.get('period', {})
        sd = period.get('start')
        ed = period.get('end')
        days = period.get('days', 0)
        period_str = (f"관측 기간: {sd.strftime('%Y-%m-%d') if sd else '-'}  ~  "
                      f"{ed.strftime('%Y-%m-%d') if ed else '-'}  (총 {days:,}일)")
        elems.append(Paragraph(period_str, styles['body']))
        elems.append(Spacer(1, 0.3*cm))

        overall = stn_stats.get('overall', {})
        if not overall:
            elems.append(Paragraph('해당 기상요소 데이터 없음', styles['note']))
            return elems

        # 통계 테이블
        table_data = [['기상요소', '통계항목', '값']]
        for elem, values in overall.items():
            first = True
            for stat_name, stat_val in values.items():
                if first:
                    table_data.append([elem, stat_name, str(stat_val)])
                    first = False
                else:
                    table_data.append(['', stat_name, str(stat_val)])

        col_w = [5*cm, 4.5*cm, 5.5*cm]
        tbl   = Table(table_data, colWidths=col_w, repeatRows=1)

        style_cmds = [
            ('FONTNAME',     (0,0), (-1,-1), FONT_NAME),
            ('FONTSIZE',     (0,0), (-1,-1), 9),
            ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING',   (0,0), (-1,-1), 5),
            ('BOTTOMPADDING',(0,0), (-1,-1), 5),
            # 헤더
            ('BACKGROUND',   (0,0), (-1,0),  MID_BLUE),
            ('TEXTCOLOR',    (0,0), (-1,0),  WHITE),
            ('FONTSIZE',     (0,0), (-1,0),  10),
            ('FONTNAME',     (0,0), (-1,0),  FONT_NAME),
            # 데이터 줄무늬
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, LIGHT_BLUE]),
            ('GRID',         (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
        ]
        tbl.setStyle(TableStyle(style_cmds))
        elems.append(tbl)
        return elems

    # ── 월별 통계 테이블 ─────────────────────────────────────────────────────
    def _monthly_tables(self, stn, stn_stats, styles) -> list:
        elems = []
        elems.append(Paragraph(f"관측소: {stn}", styles['h2']))

        monthly_all = stn_stats.get('monthly', {})
        if not monthly_all:
            elems.append(Paragraph('데이터 없음', styles['note']))
            return elems

        for elem_label, monthly_data in monthly_all.items():
            if not monthly_data:
                continue

            elems.append(Spacer(1, 0.2*cm))
            elems.append(Paragraph(f"【 {elem_label} 】", styles['body']))
            elems.append(Spacer(1, 0.1*cm))

            months    = sorted(monthly_data.keys())
            stat_keys = list(list(monthly_data.values())[0].keys())

            # 너무 많은 월은 두 줄로 나눔
            COLS_PER_ROW = 12
            for chunk_start in range(0, len(months), COLS_PER_ROW):
                chunk_months = months[chunk_start:chunk_start + COLS_PER_ROW]

                header = ['통계항목'] + chunk_months
                rows   = [header]
                for stat_key in stat_keys:
                    row = [stat_key] + [
                        str(monthly_data.get(m, {}).get(stat_key, '-'))
                        for m in chunk_months
                    ]
                    rows.append(row)

                n_cols  = len(header)
                avail_w = 17*cm
                col_w   = [3*cm] + [(avail_w - 3*cm) / (n_cols - 1)] * (n_cols - 1)

                tbl = Table(rows, colWidths=col_w, repeatRows=1)
                tbl.setStyle(TableStyle([
                    ('FONTNAME',      (0,0), (-1,-1), FONT_NAME),
                    ('FONTSIZE',      (0,0), (-1,-1), 8),
                    ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING',    (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                    ('BACKGROUND',    (0,0), (-1,0),  MID_BLUE),
                    ('TEXTCOLOR',     (0,0), (-1,0),  WHITE),
                    ('ROWBACKGROUNDS',(0,1), (-1,-1),  [WHITE, LIGHT_BLUE]),
                    ('GRID',          (0,0), (-1,-1), 0.5, colors.HexColor('#BFBFBF')),
                ]))
                elems.append(tbl)
                elems.append(Spacer(1, 0.15*cm))

        return elems

    # ── 페이지 번호 / 바닥글 ─────────────────────────────────────────────────
    def _page_footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 8)
        canvas.setFillColor(colors.HexColor('#595959'))

        # 하단 선
        canvas.setStrokeColor(DARK_BLUE)
        canvas.setLineWidth(0.5)
        canvas.line(2*cm, 1.5*cm, A4[0] - 2*cm, 1.5*cm)

        canvas.drawString(2*cm, 1.1*cm, "기상자료개방포털(data.kma.go.kr) 데이터 기반")
        canvas.drawRightString(A4[0] - 2*cm, 1.1*cm, f"{doc.page} 페이지")
        canvas.restoreState()
