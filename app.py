# -*- coding: utf-8 -*-
"""
가계부 대시보드 (Streamlit)
- CSV/XLSX 업로드, 인코딩 자동 탐색, 안전 파싱
- 요약 카드 + Plotly 파이 차트
- 사이드바 필터링
- 수입/지출 추가 폼
- CSV(utf-8-sig) / XLSX 다운로드
"""

import io
import os
import signal
import time
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------------------------
# 기본 설정
# ---------------------------------------------------------------------------
st.set_page_config(page_title="가계부 대시보드", page_icon="💰", layout="wide")

# Deploy 버튼 숨김 + 모바일 최적화 스타일
st.markdown(
    """
    <style>
    /* 우측 상단 Deploy 버튼만 숨김 (⋮ 메뉴는 유지) */
    [data-testid="stAppDeployButton"] { display: none !important; }
    .stDeployButton { display: none !important; }

    /* 본문 좌우/상단 여백 축소 (좁은 화면에서 공간 확보) */
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }

    /* 요약 metric 글자 크기 — 모바일에서 넘치지 않게 */
    [data-testid="stMetricValue"] { font-size: 1.1rem; }
    [data-testid="stMetricLabel"] { font-size: 0.8rem; }

    /* 모바일 (가로폭 640px 이하) 전용 */
    @media (max-width: 640px) {
        .block-container { padding-left: 0.6rem; padding-right: 0.6rem; }

        /* 요약 metric 더 작게 */
        [data-testid="stMetricValue"] { font-size: 0.95rem; }
        [data-testid="stMetricLabel"] { font-size: 0.72rem; }

        /* 버튼/입력 터치 영역 확보, 글자 약간 축소 */
        .stButton button { padding: 0.35rem 0.3rem; font-size: 0.85rem; }

        /* 탭 라벨이 좁은 화면에서 줄바꿈되도록 */
        .stTabs [data-baseweb="tab"] { padding: 0.3rem 0.5rem; font-size: 0.85rem; }

        /* 제목 크기 축소 */
        h1 { font-size: 1.5rem; }
        h2 { font-size: 1.2rem; }
        h3 { font-size: 1.05rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _shutdown_app():
    """Streamlit 서버와 이를 띄운 콘솔(cmd) 창까지 함께 종료한다."""
    st.markdown(
        "## 👋 앱을 종료했습니다\n이 브라우저 탭은 닫으셔도 됩니다. "
        "다시 사용하려면 `가계부_실행.bat`을 실행하세요."
    )
    st.balloons()
    time.sleep(1)

    pid = os.getpid()
    if os.name == "nt":
        import subprocess

        def _parent_pid(target):
            """target 프로세스의 부모 PID와 부모 이름을 반환 (없으면 (None, ''))."""
            try:
                out = subprocess.check_output(
                    [
                        "wmic",
                        "process",
                        "where",
                        f"ProcessId={target}",
                        "get",
                        "ParentProcessId,Name",
                        "/format:csv",
                    ],
                    stderr=subprocess.DEVNULL,
                    text=True,
                )
            except Exception:  # noqa: BLE001
                return None, ""
            # CSV: Node,Name,ParentProcessId
            for line in out.splitlines():
                parts = [p.strip() for p in line.split(",")]
                if len(parts) >= 3 and parts[-1].isdigit():
                    return int(parts[-1]), parts[1]
            return None, ""

        # 이 python을 띄운 "콘솔(cmd)" 부모만 추가로 종료 대상에 포함.
        # explorer/kiro 등 상위는 절대 건드리지 않도록 cmd/conhost 계열만 허용.
        targets = [pid]
        ppid, pname = _parent_pid(pid)
        if ppid and pname.lower() in ("cmd.exe", "conhost.exe", "powershell.exe"):
            targets.append(ppid)

        # 브라우저에 응답이 먼저 가도록 잠깐 뒤 트리째 강제 종료
        kill_cmd = "ping 127.0.0.1 -n 2 >nul & " + " & ".join(
            f"taskkill /PID {p} /T /F" for p in targets
        )
        subprocess.Popen(kill_cmd, shell=True)
    else:
        os.kill(pid, signal.SIGTERM)
        os._exit(0)


# 종료 버튼이 눌렸으면 다른 UI를 그리기 전에 즉시 종료 처리
if st.session_state.get("_shutdown"):
    _shutdown_app()
    st.stop()


# ---------------------------------------------------------------------------
# 로그인 게이트
# secrets 에 [auth] 가 있으면 비밀번호를 요구하고, 없으면(로컬 테스트) 통과한다.
#   [auth]
#   죠니 = "비밀번호1"
#   묭스니 = "비밀번호2"
# ---------------------------------------------------------------------------
def _require_login():
    try:
        users = dict(st.secrets["auth"]) if "auth" in st.secrets else {}
    except Exception:  # noqa: BLE001
        users = {}

    if not users:
        return  # 로그인 미설정(로컬 테스트) → 통과

    if st.session_state.get("_authed"):
        return  # 이미 로그인됨

    st.title("🔒 가계부 로그인")
    with st.form("login_form"):
        username = st.text_input("사용자")
        password = st.text_input("비밀번호", type="password")
        ok = st.form_submit_button("로그인")
    if ok:
        if username in users and str(users[username]) == password:
            st.session_state["_authed"] = True
            st.session_state["_user"] = username
            st.rerun()
        else:
            st.error("사용자 또는 비밀번호가 올바르지 않습니다.")
    st.stop()  # 로그인 전에는 아래 UI를 그리지 않음


_require_login()

# 표준 컬럼 정의
COLUMNS = ["날짜", "입금/출금", "입력자", "주체", "카테고리", "지불 방식", "메모", "금액"]

# Plotly 한글 폰트 설정 (시스템에 있는 폰트 중 하나 사용)
PLOTLY_FONT = "Malgun Gothic, AppleGothic, NanumGothic, sans-serif"

# "지출"로 간주하는 라벨 (유연하게 매칭)
EXPENSE_LABELS = {"출금", "지출", "expense", "withdraw", "withdrawal"}
INCOME_LABELS = {"입금", "수입", "income", "deposit"}


# ---------------------------------------------------------------------------
# 유틸: 파일 읽기 (인코딩 자동 탐색 + 안전 파싱)
# ---------------------------------------------------------------------------
def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """업로드된 CSV/XLSX 파일을 안전하게 DataFrame으로 읽는다."""
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(raw))

    # CSV: 인코딩 자동 탐색
    encodings = ["utf-8-sig", "utf-8", "cp949", "euc-kr"]
    last_err = None
    for enc in encodings:
        try:
            text = raw.decode(enc)
        except (UnicodeDecodeError, LookupError) as e:
            last_err = e
            continue

        # 메모 안의 쉼표/따옴표 문제를 피하기 위해 Python 엔진 + 견고한 옵션 사용
        try:
            df = pd.read_csv(
                io.StringIO(text),
                sep=",",
                engine="python",
                quotechar='"',
                skipinitialspace=True,
                on_bad_lines="skip",
            )
            return df
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue

    raise ValueError(f"파일을 읽을 수 없습니다. 마지막 오류: {last_err}")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼 이름 공백 제거 및 표준 컬럼 보강."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # 흔한 별칭 매핑
    aliases = {
        "구분": "입금/출금",
        "입출금": "입금/출금",
        "타입": "입금/출금",
        "결제수단": "지불 방식",
        "지불방식": "지불 방식",
        "금액(원)": "금액",
        "amount": "금액",
        "date": "날짜",
    }
    df = df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

    # 누락된 표준 컬럼은 빈 값으로 추가
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    return df


def parse_dates(series: pd.Series) -> pd.Series:
    """여러 한국식 날짜 형식을 datetime으로 변환.

    지원 예: 2025-05-01, 2025.05.01, 2025/5/1, 2025년 5월 1일,
            20250501, 2025 05 01, 엑셀 시리얼 숫자 등.
    """
    s = series.copy()

    # 0) 이미 datetime 타입이면 그대로 반환 (추가 폼에서 들어온 Timestamp 보존)
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce").dt.normalize()

    # 1) 문자열 정규화: 한글/구분자 통일
    def norm(x):
        if pd.isna(x):
            return None
        # 이미 datetime/날짜 객체면 문자열로 바꾸지 않고 그대로 둔다
        if isinstance(x, (pd.Timestamp, datetime)):
            return x
        t = str(x).strip()
        if t == "" or t.lower() in ("nan", "none", "nat"):
            return None
        # 시:분:초가 붙어 있으면 날짜 부분만 사용 ('2025-05-01 00:00:00' -> '2025-05-01')
        t = t.split(" ")[0] if (":" in t) else t
        # '2025년 5월 1일' -> '2025-5-1', 구분자 통일
        t = (
            t.replace("년", "-")
            .replace("월", "-")
            .replace("일", "")
            .replace(".", "-")
            .replace("/", "-")
            .replace(" ", "-")
        )
        while "--" in t:
            t = t.replace("--", "-")
        t = t.strip("-")
        return t

    normalized = s.map(norm)

    # 2) 일반 파싱 (연-월-일 순서 우선)
    parsed = pd.to_datetime(normalized, errors="coerce", yearfirst=True)

    # 3) 아직 비어있는 값이 정확히 8자리 숫자(YYYYMMDD)면 재시도
    missing = parsed.isna() & pd.Series(normalized, index=s.index).notna()
    if missing.any():
        cand = pd.Series(normalized, index=s.index)[missing].astype(str)
        eight = cand.str.fullmatch(r"\d{8}")
        if eight.any():
            retry = pd.to_datetime(cand[eight], format="%Y%m%d", errors="coerce")
            parsed.loc[retry.index] = retry

    # 4) 그래도 비어있고 원본이 숫자(엑셀 시리얼)면 변환 시도
    missing = parsed.isna() & s.notna()
    if missing.any():
        nums = pd.to_numeric(s[missing], errors="coerce")
        serial = pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce")
        parsed.loc[missing] = serial

    # 시각 성분 제거 (날짜만 유지)
    return parsed.dt.normalize()


def coerce_types(df: pd.DataFrame) -> pd.DataFrame:
    """날짜 -> datetime, 금액 -> 숫자형으로 변환."""
    df = df.copy()

    # 날짜 변환 (여러 형식 대응)
    df["날짜"] = parse_dates(df["날짜"])

    # 금액: 통화기호/쉼표/공백 제거 후 숫자화
    def to_number(x):
        if pd.isna(x):
            return pd.NA
        s = str(x)
        for ch in ["₩", ",", " ", "원", "\t"]:
            s = s.replace(ch, "")
        s = s.strip()
        if s in ("", "-", "nan", "None"):
            return pd.NA
        try:
            v = float(s)
            return int(v) if v.is_integer() else v
        except ValueError:
            return pd.NA

    df["금액"] = df["금액"].map(to_number)
    df["금액"] = pd.to_numeric(df["금액"], errors="coerce")

    # 문자열 컬럼 정리
    for col in ["입금/출금", "입력자", "주체", "카테고리", "지불 방식", "메모"]:
        df[col] = df[col].astype("string").str.strip()

    return df[COLUMNS]


def render_calendar(df: pd.DataFrame):
    """월별 달력에 일별 금액 표시 + 날짜 선택 시 그날 내역."""
    import calendar as _cal

    dated = df[df["날짜"].notna()].copy()
    if dated.empty:
        st.info("날짜가 있는 데이터가 없습니다. 내역을 추가하거나 파일을 업로드하세요.")
        return

    dated["_flow"] = classify_flow(dated["입금/출금"])

    # 연/월 선택 (데이터에 있는 범위 기준)
    years = sorted(dated["날짜"].dt.year.unique().tolist())
    cc1, cc2 = st.columns(2)
    with cc1:
        y = st.selectbox("연도", years, index=len(years) - 1, key="cal_year")
    months = sorted(dated[dated["날짜"].dt.year == y]["날짜"].dt.month.unique().tolist())
    with cc2:
        m = st.selectbox("월", months, index=len(months) - 1, key="cal_month") if months else None
    if m is None:
        st.info("선택한 연도에 데이터가 없습니다.")
        return

    month_df = dated[(dated["날짜"].dt.year == y) & (dated["날짜"].dt.month == m)]

    # 일별 수입/지출 집계
    daily = {}
    for day, g in month_df.groupby(month_df["날짜"].dt.day):
        inc = float(g.loc[g["_flow"] == "입금", "금액"].sum())
        exp = float(g.loc[g["_flow"] == "지출", "금액"].sum())
        daily[int(day)] = (inc, exp)

    st.markdown(f"#### {y}년 {m}월")

    # 요일 헤더
    week_days = ["월", "화", "수", "목", "금", "토", "일"]
    head = st.columns(7)
    for i, wd in enumerate(week_days):
        head[i].markdown(f"**{wd}**")

    # 달력 그리드 (월요일 시작)
    _cal.setfirstweekday(_cal.MONDAY)
    weeks = _cal.monthcalendar(y, m)
    for week in weeks:
        cols = st.columns(7)
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.write("")
                    continue
                inc, exp = daily.get(day, (0.0, 0.0))
                # 날짜 버튼
                if st.button(f"{day}", key=f"cal_day_{y}_{m}_{day}", width="stretch"):
                    st.session_state["cal_selected"] = (y, m, day)
                # 금액 표시 (지출 빨강, 수입 초록) — 좁은 칸에서 넘치지 않게 축약
                def _fmt(v):
                    if v >= 10000:
                        return f"{v/10000:.0f}만" if v % 10000 == 0 else f"{v/10000:.1f}만"
                    return f"{v:,.0f}"
                if exp:
                    st.markdown(
                        f"<div style='text-align:center;color:#d33;font-size:10px;line-height:1.1;"
                        f"white-space:nowrap;overflow:hidden'>-{_fmt(exp)}</div>",
                        unsafe_allow_html=True,
                    )
                if inc:
                    st.markdown(
                        f"<div style='text-align:center;color:#2a7;font-size:10px;line-height:1.1;"
                        f"white-space:nowrap;overflow:hidden'>+{_fmt(inc)}</div>",
                        unsafe_allow_html=True,
                    )

    # 선택한 날짜 상세
    sel = st.session_state.get("cal_selected")
    if sel and sel[0] == y and sel[1] == m:
        _, _, d = sel
        st.divider()
        st.markdown(f"### 📌 {y}년 {m}월 {d}일 내역")
        day_df = month_df[month_df["날짜"].dt.day == d]
        if day_df.empty:
            st.info("이 날짜에는 내역이 없습니다.")
        else:
            inc = float(day_df.loc[day_df["_flow"] == "입금", "금액"].sum())
            exp = float(day_df.loc[day_df["_flow"] == "지출", "금액"].sum())
            s1, s2, s3 = st.columns(3)
            s1.metric("수입", f"{inc:,.0f} 원")
            s2.metric("지출", f"{exp:,.0f} 원")
            s3.metric("합계", f"{inc - exp:,.0f} 원")

            exp_rows = day_df[day_df["_flow"] == "지출"].drop(columns=["_flow"])
            inc_rows = day_df[day_df["_flow"] == "입금"].drop(columns=["_flow"])

            def _show(rows):
                if rows.empty:
                    st.caption("내역 없음")
                    return
                st.dataframe(
                    rows,
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "날짜": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
                        "금액": st.column_config.NumberColumn("금액", format="%d"),
                    },
                )

            st.markdown("**💸 지출**")
            _show(exp_rows)
            st.markdown("**💰 수입**")
            _show(inc_rows)


def render_sidebar_footer():
    """사이드바 하단: 데이터 관리 + 앱 종료 (어느 보기 모드든 항상 그려짐)."""
    with st.sidebar:
        st.divider()
        st.header("🧹 데이터 관리")
        n_total = len(st.session_state.data)
        n_nodate = int(st.session_state.data["날짜"].isna().sum()) if n_total else 0
        st.caption(f"전체 {n_total:,}건 · 날짜 없음 {n_nodate:,}건")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("날짜없는행 삭제", disabled=n_nodate == 0, width="stretch"):
                st.session_state.data = (
                    st.session_state.data[st.session_state.data["날짜"].notna()]
                    .reset_index(drop=True)
                )
                save_store(st.session_state.data)
                st.success(f"날짜 없는 {n_nodate:,}건을 삭제했습니다.")
                st.rerun()
        with c2:
            if st.button("전체 초기화", type="secondary", disabled=n_total == 0, width="stretch"):
                st.session_state.data = pd.DataFrame(columns=COLUMNS)
                st.session_state.loaded_file = None
                save_store(st.session_state.data)
                st.success("모든 데이터를 비웠습니다. 파일을 다시 업로드하세요.")
                st.rerun()

        st.divider()
        if st.button("⏹ 앱 종료하기", type="primary", width="stretch"):
            st.session_state["_shutdown"] = True
            st.rerun()
        st.caption("종료하면 서버와 콘솔(cmd) 창이 닫힙니다. 브라우저 탭은 직접 닫으세요.")


def classify_flow(series: pd.Series) -> pd.Series:
    """입금/출금 라벨을 '입금'/'지출'로 정규화된 소문자 분류 반환."""
    def label(x):
        if pd.isna(x):
            return "기타"
        s = str(x).strip().lower()
        if s in EXPENSE_LABELS:
            return "지출"
        if s in INCOME_LABELS:
            return "입금"
        return str(x).strip()

    return series.map(label)


# ---------------------------------------------------------------------------
# 다운로드 헬퍼
# ---------------------------------------------------------------------------
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    """utf-8-sig CSV 바이트 (엑셀에서 한글 안 깨짐)."""
    out = df.copy()
    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
    return out.to_csv(index=False).encode("utf-8-sig")


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """XLSX 바이트."""
    out = df.copy()
    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
    buffer = io.BytesIO()
    # openpyxl 우선, 없으면 xlsxwriter
    try:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            out.to_excel(writer, index=False, sheet_name="가계부")
    except ModuleNotFoundError:
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            out.to_excel(writer, index=False, sheet_name="가계부")
    return buffer.getvalue()


def style_pie(fig):
    fig.update_traces(
        textposition="inside",
        texttemplate="%{label}<br>%{percent}<br>%{value:,.0f}원",
        hovertemplate="%{label}<br>%{value:,.0f}원 (%{percent})<extra></extra>",
        insidetextorientation="horizontal",
    )
    fig.update_layout(
        font=dict(family=PLOTLY_FONT),
        height=430,
        # 범례를 그래프 하단 바깥에 가로로 배치하고, 아래 여백을 넉넉히 확보해 겹침 방지
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.08,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
        margin=dict(t=50, b=110, l=10, r=10),
    )
    return fig


# ---------------------------------------------------------------------------
# 영구 저장소 (로컬 CSV 파일에 누적 저장/로드)
# ---------------------------------------------------------------------------
import os

STORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "가계부_데이터.csv")

# Google Sheets 백엔드 (secrets 설정 시 자동 활성화, 없으면 로컬 CSV 사용)
try:
    import sheets_backend
    USE_SHEETS = sheets_backend.is_enabled()
except Exception:  # noqa: BLE001
    sheets_backend = None
    USE_SHEETS = False


def storage_label() -> str:
    """현재 사용 중인 저장소 설명 문구."""
    return "☁️ Google Sheets (공유)" if USE_SHEETS else "💾 로컬 CSV"


def dedupe(df: pd.DataFrame) -> pd.DataFrame:
    """완전히 동일한 내역의 중복만 제거하고 날짜순 정렬."""
    df = df.drop_duplicates(subset=COLUMNS, keep="first").reset_index(drop=True)
    return df.sort_values("날짜", na_position="last").reset_index(drop=True)


def load_store() -> pd.DataFrame:
    """저장소에서 데이터를 불러온다. Google Sheets 우선, 없으면 로컬 CSV."""
    if USE_SHEETS:
        try:
            raw = sheets_backend.load()
            return coerce_types(normalize_columns(raw))
        except Exception as e:  # noqa: BLE001
            st.warning(f"Google Sheets 로드 실패, 로컬 CSV로 대체합니다: {e}")
    if os.path.exists(STORE_PATH):
        try:
            raw = pd.read_csv(STORE_PATH, encoding="utf-8-sig")
            return coerce_types(normalize_columns(raw))
        except Exception:  # noqa: BLE001
            return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(columns=COLUMNS)


def save_store(df: pd.DataFrame) -> None:
    """현재 데이터를 저장소에 저장. Google Sheets 우선, 없으면 로컬 CSV."""
    if USE_SHEETS:
        try:
            sheets_backend.save(df)
            return
        except Exception as e:  # noqa: BLE001
            st.warning(f"Google Sheets 저장 실패, 로컬 CSV에 저장합니다: {e}")
    out = df.copy()
    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
    out.to_csv(STORE_PATH, index=False, encoding="utf-8-sig")


def merge_into_store(new_df: pd.DataFrame) -> None:
    """새 데이터를 기존 세션 데이터에 누적 병합 후 저장."""
    combined = pd.concat([st.session_state.data, new_df], ignore_index=True)
    st.session_state.data = dedupe(coerce_types(normalize_columns(combined)))
    save_store(st.session_state.data)


# ---------------------------------------------------------------------------
# 세션 상태 초기화 (시작 시 저장 파일에서 자동 로드)
# ---------------------------------------------------------------------------
if "data" not in st.session_state:
    st.session_state.data = load_store()
if "loaded_file" not in st.session_state:
    st.session_state.loaded_file = None


# ---------------------------------------------------------------------------
# 헤더 & 파일 업로드
# ---------------------------------------------------------------------------
st.title("💰 가계부 대시보드")
st.caption("CSV / XLSX 업로드 → 데이터 누적 저장 · 연/월 조회 · 시각화 · 내역 추가/삭제")

uploaded = st.file_uploader(
    "가계부 파일 업로드 (.csv 또는 .xlsx) — 업로드한 내역은 기존 데이터에 누적됩니다",
    type=["csv", "xlsx", "xls"],
)

if uploaded is not None and uploaded.name != st.session_state.loaded_file:
    try:
        before = len(st.session_state.data)
        df_raw = read_uploaded_file(uploaded)
        df_clean = coerce_types(normalize_columns(df_raw))
        merge_into_store(df_clean)  # 기존 데이터에 누적 + 파일 저장
        added = len(st.session_state.data) - before
        st.session_state.loaded_file = uploaded.name
        st.success(
            f"'{uploaded.name}' 불러오기 완료 — 신규 {added:,}건 추가 "
            f"(전체 {len(st.session_state.data):,}건, 중복 제외)"
        )
    except Exception as e:  # noqa: BLE001
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

data = st.session_state.data

if USE_SHEETS:
    st.caption(f"저장소: {storage_label()} — 추가/삭제/업로드 시 공유 시트에 자동 저장됩니다.")
else:
    st.caption(f"저장소: {storage_label()} — `{STORE_PATH}` 에 저장됩니다.")

if data.empty:
    st.info("파일을 업로드하거나, 사이드바에서 직접 내역을 추가해 시작하세요.")


# ---------------------------------------------------------------------------
# 사이드바: 내역 추가 폼
# ---------------------------------------------------------------------------
def unique_options(col):
    if data.empty:
        return []
    vals = (
        data[col]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(vals)


def select_or_input(label, col, key):
    """빈 칸으로 시작. 화살표를 누르면 기존 값 목록이 뜨고, 없는 값은 직접 타이핑해 추가."""
    options = unique_options(col)
    picked = st.selectbox(
        label,
        options,
        index=None,  # 기본값: 선택 안 함(빈 칸)
        placeholder="선택하거나 입력하세요",
        accept_new_options=True,  # 목록에 없는 값 직접 입력 허용
        key=f"{key}_sel",
    )
    return (picked or "").strip()


with st.sidebar:
    st.header("➕ 내역 추가")
    st.caption("빈 칸에 직접 입력하거나, 화살표를 눌러 기존 값에서 선택하세요.")
    with st.form("add_form", clear_on_submit=True):
        in_date = st.date_input("날짜", value=date.today())
        in_flow = st.selectbox("입금/출금", ["지출", "입금"])
        in_writer = select_or_input("입력자", "입력자", "add_writer")
        in_subject = select_or_input("주체", "주체", "add_subject")
        in_category = select_or_input("카테고리", "카테고리", "add_category")
        in_pay = select_or_input("지불 방식", "지불 방식", "add_pay")
        in_memo = st.text_input("메모")
        in_amount = st.number_input("금액", min_value=0, step=1000, value=0)
        submitted = st.form_submit_button("추가하기")

        if submitted:
            new_row = {
                "날짜": pd.to_datetime(in_date),
                "입금/출금": in_flow,
                "입력자": in_writer,
                "주체": in_subject,
                "카테고리": in_category,
                "지불 방식": in_pay,
                "메모": in_memo.strip(),
                "금액": int(in_amount),
            }
            merge_into_store(pd.DataFrame([new_row]))  # 누적 + 파일 저장
            st.success("내역이 추가되어 저장되었습니다.")
            st.rerun()

    st.divider()
    st.header("🔎 세부 필터")
    st.caption("기간은 메인 화면 상단에서 선택하세요.")


data = st.session_state.data  # 폼 추가 반영

# ---------------------------------------------------------------------------
# 사이드바: 필터
# ---------------------------------------------------------------------------
filtered = data.copy()

if not data.empty:
    valid_dates = data["날짜"].dropna()
    date_range = None
    year_sel = "전체"
    month_sel = "전체"
    years = sorted(valid_dates.dt.year.unique().tolist()) if not valid_dates.empty else []

    # --- 조회 기간: 메인 상단에 가로 배치 ---
    st.subheader("🗓️ 조회 기간")
    pc0, pc1, pc2 = st.columns([1.4, 1, 1])
    with pc0:
        period_mode = st.radio(
            "기간 방식", ["전체", "연도별", "연·월별", "날짜 범위"], horizontal=True
        )

    if period_mode == "연도별" and years:
        with pc1:
            year_sel = st.selectbox("연도", years, index=len(years) - 1)

    elif period_mode == "연·월별" and years:
        with pc1:
            year_sel = st.selectbox("연도", years, index=len(years) - 1)
        months_in_year = sorted(
            valid_dates[valid_dates.dt.year == year_sel].dt.month.unique().tolist()
        )
        with pc2:
            month_sel = st.selectbox("월", months_in_year) if months_in_year else "전체"

    elif period_mode == "날짜 범위" and not valid_dates.empty:
        min_d = valid_dates.min().date()
        max_d = valid_dates.max().date()
        with pc1:
            date_range = st.date_input(
                "날짜 범위", value=(min_d, max_d), min_value=min_d, max_value=max_d
            )

    # --- 세부 필터: 사이드바 유지 ---
    with st.sidebar:
        # 입금/출금
        flow_sel = st.selectbox("입금/출금 구분", ["전체", "입금", "지출"])

        # 입력자
        writer_opts = ["전체"] + unique_options("입력자")
        writer_sel = st.selectbox("입력자", writer_opts)

        # 주체
        subject_opts = ["전체"] + unique_options("주체")
        subject_sel = st.selectbox("주체", subject_opts)

        # 지불 방식
        pay_opts = ["전체"] + unique_options("지불 방식")
        pay_sel = st.selectbox("지불 방식", pay_opts)

        # 카테고리
        cat_opts = ["전체"] + unique_options("카테고리")
        cat_sel = st.selectbox("카테고리", cat_opts)

    # 기간 필터 적용
    if period_mode == "연도별" and year_sel != "전체":
        filtered = filtered[filtered["날짜"].dt.year == year_sel]

    elif period_mode == "연·월별" and year_sel != "전체" and month_sel != "전체":
        filtered = filtered[
            (filtered["날짜"].dt.year == year_sel) & (filtered["날짜"].dt.month == month_sel)
        ]

    elif period_mode == "날짜 범위":
        if date_range and isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            start, end = date_range
            end_dt = pd.to_datetime(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            mask = (filtered["날짜"] >= pd.to_datetime(start)) & (filtered["날짜"] <= end_dt)
            filtered = filtered[mask]
        else:
            st.sidebar.info("종료일까지 선택하면 범위가 적용됩니다.")

    flow_norm = classify_flow(filtered["입금/출금"])
    if flow_sel != "전체":
        filtered = filtered[flow_norm == flow_sel]

    if writer_sel != "전체":
        filtered = filtered[filtered["입력자"].astype(str).str.strip() == writer_sel]
    if subject_sel != "전체":
        filtered = filtered[filtered["주체"].astype(str).str.strip() == subject_sel]
    if pay_sel != "전체":
        filtered = filtered[filtered["지불 방식"].astype(str).str.strip() == pay_sel]
    if cat_sel != "전체":
        filtered = filtered[filtered["카테고리"].astype(str).str.strip() == cat_sel]


# ---------------------------------------------------------------------------
# 적용된 필터 요약 배지
# ---------------------------------------------------------------------------
if not data.empty:
    badges = []
    # 기간
    if period_mode == "연도별" and year_sel != "전체":
        badges.append(f"📅 {year_sel}년")
    elif period_mode == "연·월별" and year_sel != "전체" and month_sel != "전체":
        badges.append(f"📅 {year_sel}년 {month_sel}월")
    elif period_mode == "날짜 범위" and date_range and isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        badges.append(f"📅 {date_range[0]} ~ {date_range[1]}")
    else:
        badges.append("📅 전체 기간")
    # 세부 필터
    if flow_sel != "전체":
        badges.append(f"입출금: {flow_sel}")
    if writer_sel != "전체":
        badges.append(f"입력자: {writer_sel}")
    if subject_sel != "전체":
        badges.append(f"주체: {subject_sel}")
    if pay_sel != "전체":
        badges.append(f"지불: {pay_sel}")
    if cat_sel != "전체":
        badges.append(f"카테고리: {cat_sel}")

    st.caption("적용된 필터")
    st.markdown("  ".join(f"`{b}`" for b in badges) + f"  —  **{len(filtered):,}건**")


# ---------------------------------------------------------------------------
# 보기 모드: 대시보드 / 달력
# ---------------------------------------------------------------------------
view_mode = st.radio(
    "보기", ["📊 대시보드", "📅 달력"], horizontal=True, label_visibility="collapsed"
)

# 사이드바 하단(데이터 관리/종료)은 보기 모드와 무관하게 항상 그린다.
render_sidebar_footer()

if view_mode == "📅 달력":
    st.subheader("📅 달력")
    st.caption("날짜 칸의 금액은 그날 합계입니다. 날짜를 클릭하면 아래에 상세 내역이 나옵니다.")
    render_calendar(filtered)
    st.stop()  # 달력 모드에서는 아래 대시보드를 그리지 않음


# ---------------------------------------------------------------------------
# 요약 카드
# ---------------------------------------------------------------------------
flow_class = classify_flow(filtered["입금/출금"]) if not filtered.empty else pd.Series(dtype=str)

income_total = float(filtered.loc[flow_class == "입금", "금액"].sum()) if not filtered.empty else 0.0
expense_total = float(filtered.loc[flow_class == "지출", "금액"].sum()) if not filtered.empty else 0.0
net_total = income_total - expense_total

st.subheader("📊 요약")
c1, c2, c3 = st.columns(3)
c1.metric("총 수입", f"{income_total:,.0f} 원")
c2.metric("총 지출", f"{expense_total:,.0f} 원")
c3.metric("순수입 (수입-지출)", f"{net_total:,.0f} 원", delta=f"{net_total:,.0f} 원")


# ---------------------------------------------------------------------------
# 시각화: 파이 차트
# ---------------------------------------------------------------------------
def pie_by(df, group_col, value_col="금액", title=""):
    if df.empty:
        st.info(f"{title}: 데이터가 없습니다.")
        return
    g = (
        df.assign(**{group_col: df[group_col].fillna("미지정").replace("", "미지정")})
        .groupby(group_col, dropna=False)[value_col]
        .sum()
        .reset_index()
    )
    g = g[g[value_col] > 0]
    if g.empty:
        st.info(f"{title}: 표시할 값이 없습니다.")
        return
    fig = px.pie(g, names=group_col, values=value_col, title=title, hole=0.3)
    st.plotly_chart(style_pie(fig), width="stretch")


if not filtered.empty:
    st.subheader("📈 시각화")

    expense_df = filtered[flow_class == "지출"]
    income_df = filtered[flow_class == "입금"]

    tab1, tab2 = st.tabs(["지출 분석", "수입 분석"])

    with tab1:
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            pie_by(expense_df, "카테고리", title="카테고리별 지출")
        with col_b:
            pie_by(expense_df, "지불 방식", title="지불 방식별 지출")
        with col_c:
            pie_by(expense_df, "주체", title="주체별 지출")

    with tab2:
        col_d, col_e = st.columns(2)
        with col_d:
            pie_by(income_df, "카테고리", title="카테고리별 수입")
        with col_e:
            pie_by(income_df, "주체", title="주체별 수입")


# ---------------------------------------------------------------------------
# 데이터 테이블 + 다운로드
# ---------------------------------------------------------------------------
st.subheader("📋 데이터")


def _data_table(sub_df, editor_key):
    """체크박스 삭제가 가능한 데이터 표를 그리고, 삭제 버튼까지 처리한다."""
    if sub_df.empty:
        st.info("표시할 내역이 없습니다.")
        return
    view = sub_df.sort_values("날짜", ascending=False, na_position="last").copy()
    view.insert(0, "🗑 삭제", False)

    edited = st.data_editor(
        view,
        width="stretch",
        hide_index=True,
        column_config={
            "🗑 삭제": st.column_config.CheckboxColumn(
                "삭제", help="삭제할 행을 체크한 뒤 아래 버튼을 누르세요", default=False
            ),
            "날짜": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
            "금액": st.column_config.NumberColumn("금액", format="%d"),
        },
        disabled=[c for c in view.columns if c != "🗑 삭제"],
        key=editor_key,
    )

    del_count = int(edited["🗑 삭제"].sum())
    col_del, col_info = st.columns([1, 3])
    with col_del:
        if st.button(
            f"선택한 {del_count}건 삭제",
            type="primary",
            disabled=del_count == 0,
            key=f"{editor_key}_delbtn",
        ):
            drop_idx = edited.index[edited["🗑 삭제"]].tolist()
            st.session_state.data = st.session_state.data.drop(index=drop_idx).reset_index(drop=True)
            save_store(st.session_state.data)
            st.success(f"{len(drop_idx)}건을 삭제했습니다.")
            st.rerun()
    with col_info:
        st.caption("체크박스로 행을 선택한 뒤 '삭제' 버튼을 누르면 해당 내역이 제거됩니다.")


if filtered.empty:
    st.info("표시할 데이터가 없습니다.")
else:
    exp_tbl = filtered[flow_class == "지출"]
    inc_tbl = filtered[flow_class == "입금"]
    etc_tbl = filtered[~flow_class.isin(["지출", "입금"])]

    tabs = [
        f"📋 전체 ({len(filtered):,})",
        f"💸 지출 ({len(exp_tbl):,})",
        f"💰 수입 ({len(inc_tbl):,})",
    ]
    if not etc_tbl.empty:
        tabs.append(f"기타 ({len(etc_tbl):,})")
    data_tabs = st.tabs(tabs)

    with data_tabs[0]:
        _data_table(filtered, "editor_all")
    with data_tabs[1]:
        _data_table(exp_tbl, "editor_expense")
    with data_tabs[2]:
        _data_table(inc_tbl, "editor_income")
    if not etc_tbl.empty:
        with data_tabs[3]:
            _data_table(etc_tbl, "editor_etc")

st.caption(f"필터 적용: {len(filtered):,}건 / 전체: {len(data):,}건")

st.subheader("⬇️ 다운로드")
dl1, dl2 = st.columns(2)
export_df = st.session_state.data  # 전체 (추가 포함) 데이터 내보내기

with dl1:
    st.download_button(
        "CSV 다운로드 (utf-8-sig)",
        data=to_csv_bytes(export_df),
        file_name=f"가계부_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        width="stretch",
        disabled=export_df.empty,
    )
with dl2:
    st.download_button(
        "Excel 다운로드 (.xlsx)",
        data=to_xlsx_bytes(export_df) if not export_df.empty else b"",
        file_name=f"가계부_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=export_df.empty,
    )


# (사이드바 하단은 render_sidebar_footer() 로 분리되어 메인 흐름에서 호출됨)
