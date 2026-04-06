# data_processor.py
# 기상청 데이터 처리 모듈
# 기상자료개방포털(data.kma.go.kr)에서 다운받은 CSV/Excel 파일을 처리합니다.

import pandas as pd
import numpy as np
import os
import logging

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# 기상청 CSV 컬럼명 → 내부 표준명 매핑
# 우선순위 높은 컬럼이 먼저 오도록 순서 중요!
# (동일한 표준명으로 여러 컬럼이 매핑될 경우, 먼저 등록된 컬럼 우선)
# ─────────────────────────────────────────────
COLUMN_MAPPING = {
    # ── 날짜 ──
    '일시': 'date', '날짜': 'date', '기준일': 'date', '관측일': 'date',

    # ── 기온 (실제 기상청 ASOS 컬럼명) ──
    '평균기온(°C)': 'temp_avg', '평균기온': 'temp_avg', '기온(°C)': 'temp_avg',
    '최고기온(°C)': 'temp_max', '최고기온': 'temp_max', '일최고기온(°C)': 'temp_max',
    '최저기온(°C)': 'temp_min', '최저기온': 'temp_min', '일최저기온(°C)': 'temp_min',

    # ── 강수량: 일강수량(mm) 최우선 ──
    '일강수량(mm)': 'precipitation',
    '강수량(mm)': 'precipitation',
    '강수량': 'precipitation',
    # 아래는 중복 방지를 위해 별도 내부명 사용 (통계에서 제외)
    '10분 최다 강수량(mm)': '_precip_10m',
    '1시간 최다강수량(mm)': '_precip_1h',
    '강수 계속시간(hr)': '_precip_dur',

    # ── 습도: 실제 기상청 컬럼명 '평균 상대습도(%)' ──
    '평균 상대습도(%)': 'humidity',
    '평균상대습도(%)': 'humidity',
    '평균습도(%)': 'humidity',
    '평균습도': 'humidity',
    '최소 상대습도(%)': '_humidity_min',

    # ── 풍속/풍향: 실제 기상청 컬럼명 '평균 풍속(m/s)' (공백 포함) ──
    '평균 풍속(m/s)': 'wind_speed',
    '평균풍속(m/s)': 'wind_speed',
    '평균풍속': 'wind_speed',
    '최대 풍속(m/s)': 'wind_max',
    '최대풍속(m/s)': 'wind_max',
    '최대 순간 풍속(m/s)': '_wind_gust',
    '최대순간풍속(m/s)': '_wind_gust',
    '최다풍향(16방위)': 'wind_dir',
    '최다풍향(deg)': 'wind_dir',
    '최다풍향': 'wind_dir',
    '최대 풍속 풍향(16방위)': '_wind_max_dir',
    '최대 순간 풍속 풍향(16방위)': '_wind_gust_dir',

    # ── 일조/일사: 실제 기상청 컬럼명 '합계 일조시간(hr)', '합계 일사량(MJ/m2)' ──
    '합계 일조시간(hr)': 'sunshine',
    '합계일조시간(hr)': 'sunshine',
    '일조시간(hr)': 'sunshine',
    '일조시간': 'sunshine',
    '가조시간(hr)': '_possible_sunshine',
    '합계 일사량(MJ/m2)': 'solar_rad',
    '합계일사량(MJ/m²)': 'solar_rad',
    '합계일사량(MJ/m2)': 'solar_rad',
    '일사량(MJ/m²)': 'solar_rad',
    '일사량': 'solar_rad',
    '1시간 최다일사량(MJ/m2)': '_solar_1h',

    # ── 적설: '일 최심신적설(cm)' 최우선 ──
    '일 최심신적설(cm)': 'snowfall',
    '최심신적설(cm)': 'snowfall',
    '신적설(cm)': 'snowfall',
    '적설(cm)': 'snowfall',
    '일 최심적설(cm)': '_snowdepth',
    '합계 3시간 신적설(cm)': '_snow_3h',

    # ── 관측소 정보 ──
    '지점': 'station_id', '지점번호': 'station_id',
    '지점명': 'station_name', '관측소명': 'station_name',
}

# 내부용(_로 시작) 컬럼은 통계에서 제외
_INTERNAL_COLS = {v for v in COLUMN_MAPPING.values() if v.startswith('_')}

# 내부 표준명 → 한글 표시명
ELEMENT_LABELS = {
    'temp_avg':     '평균기온(°C)',
    'temp_max':     '최고기온(°C)',
    'temp_min':     '최저기온(°C)',
    'precipitation':'강수량(mm)',
    'humidity':     '평균습도(%)',
    'wind_speed':   '평균풍속(m/s)',
    'wind_max':     '최대풍속(m/s)',
    'wind_dir':     '최다풍향',
    'sunshine':     '일조시간(hr)',
    'solar_rad':    '일사량(MJ/m²)',
    'snowfall':     '적설(cm)',
}

# 설정 키 → 실제 컬럼 그룹 매핑
ELEMENT_GROUPS = {
    'temperature':   ['temp_avg', 'temp_max', 'temp_min'],
    'precipitation': ['precipitation'],
    'humidity':      ['humidity'],
    'wind':          ['wind_speed', 'wind_max', 'wind_dir'],
    'sunshine':      ['sunshine', 'solar_rad'],
    'snowfall':      ['snowfall'],
}

# 강수량/적설처럼 합계로 집계하는 요소
SUM_ELEMENTS   = {'precipitation', 'snowfall'}
# 풍향처럼 최빈값으로 집계하는 요소
MODE_ELEMENTS  = {'wind_dir'}


class WeatherDataProcessor:
    """기상청 데이터 로드 및 처리 클래스"""

    def __init__(self):
        self.dataframes = []   # 불러온 파일 목록 (dict 리스트)
        self.merged_df  = None # 모든 파일을 합친 데이터프레임
        self.stations   = []   # 관측소명 목록

    # ── 파일 불러오기 ──────────────────────────────
    def load_file(self, filepath: str) -> dict:
        """
        CSV 또는 Excel 파일 불러오기.
        반환: {'success': bool, 'message': str, 'info': dict}
        """
        try:
            ext = os.path.splitext(filepath)[1].lower()
            df  = self._read_file(filepath, ext)

            if df is None:
                return {'success': False, 'message': '파일을 읽을 수 없습니다. 기상청 CSV 형식인지 확인하세요.'}

            df = self._standardize_columns(df)

            if 'date' not in df.columns:
                return {
                    'success': False,
                    'message': '날짜 컬럼(일시)을 찾지 못했습니다.\n기상자료개방포털 표준 형식인지 확인하세요.'
                }

            df = self._parse_dates(df)
            df = self._convert_numeric(df)
            df = df.dropna(subset=['date'])
            df = df.sort_values('date').reset_index(drop=True)

            station = self._extract_station_info(df, filepath)

            self.dataframes.append({
                'df':       df,
                'filepath': filepath,
                'filename': os.path.basename(filepath),
                'station':  station,
            })

            found_cols = [ELEMENT_LABELS[c] for c in df.columns if c in ELEMENT_LABELS]
            info = {
                'rows':       len(df),
                'date_range': (f"{df['date'].min().strftime('%Y-%m-%d')} ~"
                               f" {df['date'].max().strftime('%Y-%m-%d')}"),
                'station':    station['name'],
                'columns':    found_cols,
            }
            return {'success': True, 'message': '파일 로드 성공', 'info': info}

        except Exception as e:
            logger.error(f"파일 로드 오류: {e}")
            return {'success': False, 'message': f'오류 발생: {e}'}

    def _read_file(self, filepath: str, ext: str):
        """다양한 인코딩·헤더 행 조합으로 파일 읽기 시도"""
        encodings = ['utf-8-sig', 'euc-kr', 'cp949', 'utf-8']
        skip_candidates = [0, 1, 2, 3, 7, 8]

        if ext in ('.xlsx', '.xls'):
            for skip in skip_candidates:
                try:
                    df = pd.read_excel(filepath, skiprows=skip)
                    if self._looks_like_kma(df):
                        return df
                except Exception:
                    continue
            try:
                return pd.read_excel(filepath)
            except Exception:
                return None
        else:
            for enc in encodings:
                for skip in skip_candidates:
                    try:
                        df = pd.read_csv(filepath, encoding=enc, skiprows=skip)
                        if self._looks_like_kma(df):
                            return df
                    except Exception:
                        continue
            # 마지막 수단
            for enc in encodings:
                try:
                    return pd.read_csv(filepath, encoding=enc)
                except Exception:
                    continue
        return None

    def _looks_like_kma(self, df: pd.DataFrame) -> bool:
        """기상청 데이터처럼 보이는지 확인 (날짜 컬럼 존재 여부)"""
        if df.empty or len(df.columns) < 2:
            return False
        date_kw = ('일시', '날짜', '기준일', '관측일')
        return any(kw in str(c) for c in df.columns for kw in date_kw)

    def _standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        컬럼명을 표준 영문명으로 변환.
        동일한 표준명으로 여러 원본 컬럼이 매핑될 경우,
        가장 먼저 등장하는 컬럼만 사용하고 나머지는 고유한 임시명으로 대체
        (중복 컬럼으로 인한 to_numeric 오류 방지)
        """
        rename = {}
        used_targets = set()  # 이미 사용된 표준명 추적

        for col in df.columns:
            s = str(col).strip()
            s_nospace = s.replace(" ", "")

            # 정확히 일치하는 매핑 먼저
            target = COLUMN_MAPPING.get(s)

            # 없으면 공백 제거 후 부분 일치
            if target is None:
                for kma_key, std_key in COLUMN_MAPPING.items():
                    kma_nospace = kma_key.replace(" ", "")
                    if kma_key in s or kma_nospace in s_nospace:
                        target = std_key
                        break

            if target is None:
                continue  # 매핑 없으면 원래 이름 유지

            if target not in used_targets:
                rename[col] = target
                used_targets.add(target)
            else:
                # 이미 사용된 표준명 → _dup_ 접두사로 처리 (통계에서 자동 제외)
                rename[col] = f"_dup_{col}"

        return df.rename(columns=rename)


    def _parse_dates(self, df: pd.DataFrame) -> pd.DataFrame:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        return df

    def _convert_numeric(self, df: pd.DataFrame) -> pd.DataFrame:
        for col in df.columns:
            if col in ELEMENT_LABELS:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df

    def _extract_station_info(self, df: pd.DataFrame, filepath: str) -> dict:
        """관측소명/번호 추출"""
        name = ''
        sid  = ''

        if 'station_name' in df.columns:
            vals = df['station_name'].dropna().unique()
            if len(vals):
                name = str(vals[0])

        if 'station_id' in df.columns:
            vals = df['station_id'].dropna().unique()
            if len(vals):
                try:
                    sid = str(int(vals[0]))
                except Exception:
                    sid = str(vals[0])

        if not name:
            name = os.path.splitext(os.path.basename(filepath))[0]

        return {'name': name, 'id': sid}

    # ── 데이터 병합 ────────────────────────────────
    def merge_all(self) -> bool:
        """불러온 모든 파일 병합"""
        if not self.dataframes:
            return False

        dfs = []
        for item in self.dataframes:
            df = item['df'].copy()
            if 'station_name' not in df.columns:
                df['station_name'] = item['station']['name']
            dfs.append(df)

        self.merged_df = pd.concat(dfs, ignore_index=True)
        self.merged_df = self.merged_df.sort_values(['station_name', 'date']).reset_index(drop=True)

        if 'station_name' in self.merged_df.columns:
            self.stations = self.merged_df['station_name'].dropna().unique().tolist()
        else:
            self.stations = ['전체']
        return True

    # ── 날짜 필터 ──────────────────────────────────
    def filter_by_date(self, start: str = '', end: str = '') -> pd.DataFrame:
        df = self.merged_df.copy()
        placeholder = 'YYYY-MM-DD'
        try:
            if start and start != placeholder:
                df = df[df['date'] >= pd.to_datetime(start)]
        except Exception:
            pass
        try:
            if end and end != placeholder:
                df = df[df['date'] <= pd.to_datetime(end)]
        except Exception:
            pass
        return df

    # ── 선택 컬럼 목록 ────────────────────────────
    def get_selected_columns(self, element_settings: dict) -> list:
        cols = []
        for group, enabled in element_settings.items():
            if enabled and group in ELEMENT_GROUPS:
                cols.extend(ELEMENT_GROUPS[group])
        return cols

    # ── 통계 계산 ─────────────────────────────────
    def calculate_statistics(self, df: pd.DataFrame, selected_cols: list) -> dict:
        """관측소별 전체/월별 통계 계산"""
        avail = [c for c in selected_cols if c in df.columns and c in ELEMENT_LABELS]
        stats = {}

        if 'station_name' in df.columns:
            for stn in df['station_name'].dropna().unique():
                stats[stn] = self._calc_one_station(df[df['station_name'] == stn], avail)
        else:
            stats['전체'] = self._calc_one_station(df, avail)

        return stats

    def _calc_one_station(self, df: pd.DataFrame, cols: list) -> dict:
        result = {
            'period': {
                'start': df['date'].min(),
                'end':   df['date'].max(),
                'days':  len(df),
            },
            'overall': {},
            'monthly': {},
        }

        for col in cols:
            if col not in df.columns:
                continue
            series = df[col].dropna()
            if series.empty:
                continue

            label = ELEMENT_LABELS[col]

            # 전체 통계
            if col in SUM_ELEMENTS:
                days_with = int((series > 0).sum())
                result['overall'][label] = {
                    '기간합계': round(float(series.sum()), 1),
                    '일평균':   round(float(series.mean()), 1),
                    '최대':     round(float(series.max()), 1),
                    '발생일수': days_with,
                }
            elif col in MODE_ELEMENTS:
                result['overall'][label] = {
                    '최다풍향': str(series.mode().iloc[0]) if not series.mode().empty else '-',
                }
            else:
                result['overall'][label] = {
                    '평균':   round(float(series.mean()), 1),
                    '최대':   round(float(series.max()), 1),
                    '최소':   round(float(series.min()), 1),
                }

            # 월별 통계
            tmp = df[['date', col]].copy()
            tmp['ym'] = tmp['date'].dt.to_period('M')
            monthly = {}
            for period, grp in tmp.groupby('ym'):
                s = grp[col].dropna()
                if s.empty:
                    continue
                ym_str = str(period)
                if col in SUM_ELEMENTS:
                    monthly[ym_str] = {
                        '합계': round(float(s.sum()), 1),
                        '최대': round(float(s.max()), 1),
                    }
                elif col in MODE_ELEMENTS:
                    monthly[ym_str] = {
                        '최다풍향': str(s.mode().iloc[0]) if not s.mode().empty else '-',
                    }
                else:
                    monthly[ym_str] = {
                        '평균': round(float(s.mean()), 1),
                        '최대': round(float(s.max()), 1),
                        '최소': round(float(s.min()), 1),
                    }
            result['monthly'][label] = monthly

        return result

    # ── 미리보기 텍스트 ───────────────────────────
    def get_preview_text(self) -> str:
        if not self.dataframes:
            return "불러온 파일이 없습니다."
        lines = [f"▶ 불러온 파일: {len(self.dataframes)}개\n"]
        for i, item in enumerate(self.dataframes):
            df  = item['df']
            stn = item['station']['name']
            found = [ELEMENT_LABELS[c] for c in df.columns if c in ELEMENT_LABELS]
            lines += [
                f"[{i+1}] {item['filename']}",
                f"    관측소 : {stn}",
                f"    기간   : {df['date'].min().strftime('%Y-%m-%d')} ~ {df['date'].max().strftime('%Y-%m-%d')}",
                f"    행 수  : {len(df):,}행",
                f"    요소   : {', '.join(found) if found else '없음'}",
                "",
            ]
        return "\n".join(lines)

    def remove_file(self, index: int):
        """특정 인덱스의 파일 제거"""
        if 0 <= index < len(self.dataframes):
            self.dataframes.pop(index)

    def clear(self):
        self.dataframes = []
        self.merged_df  = None
        self.stations   = []
