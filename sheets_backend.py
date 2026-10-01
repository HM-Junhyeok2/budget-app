# -*- coding: utf-8 -*-
"""Google Sheets 저장 백엔드.

- Streamlit secrets에 서비스 계정 키와 시트 정보가 있으면 활성화된다.
- 설정이 없으면 enabled=False 가 되어, app.py 는 로컬 CSV로 자동 폴백한다.

필요한 secrets 구조 (.streamlit/secrets.toml 또는 Streamlit Cloud):

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "...@...iam.gserviceaccount.com"
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"

    [gsheets]
    spreadsheet = "가계부"          # 스프레드시트 이름 또는 URL
    worksheet = "데이터"            # 워크시트(탭) 이름
"""
from __future__ import annotations

import pandas as pd

COLUMNS = ["날짜", "입금/출금", "입력자", "주체", "카테고리", "지불 방식", "메모", "금액"]


def _get_secrets():
    """streamlit secrets 를 읽는다. 없으면 None."""
    try:
        import streamlit as st

        if "gcp_service_account" in st.secrets and "gsheets" in st.secrets:
            return st.secrets
    except Exception:  # noqa: BLE001
        pass
    return None


def is_enabled() -> bool:
    return _get_secrets() is not None


def _open_worksheet():
    """gspread 로 워크시트 핸들을 연다. 실패 시 예외."""
    import gspread
    from google.oauth2.service_account import Credentials

    secrets = _get_secrets()
    if secrets is None:
        raise RuntimeError("Google Sheets secrets 가 설정되지 않았습니다.")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(secrets["gcp_service_account"]), scopes=scopes
    )
    client = gspread.authorize(creds)

    cfg = secrets["gsheets"]
    name_or_url = cfg["spreadsheet"]
    ws_name = cfg.get("worksheet", "데이터")

    # 이름 또는 URL 로 열기
    if str(name_or_url).startswith("http"):
        sh = client.open_by_url(name_or_url)
    else:
        sh = client.open(name_or_url)

    try:
        ws = sh.worksheet(ws_name)
    except Exception:  # noqa: BLE001 - 워크시트가 없으면 생성
        ws = sh.add_worksheet(title=ws_name, rows=1000, cols=len(COLUMNS))
        ws.update([COLUMNS])
    return ws


def load() -> pd.DataFrame:
    """시트에서 전체 데이터를 읽어 DataFrame 으로 반환."""
    ws = _open_worksheet()
    records = ws.get_all_records()  # 첫 행을 헤더로 사용
    if not records:
        return pd.DataFrame(columns=COLUMNS)
    df = pd.DataFrame(records)
    # 누락 컬럼 보강
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA
    return df[COLUMNS]


def save(df: pd.DataFrame) -> None:
    """DataFrame 전체를 시트에 덮어쓴다 (헤더 포함)."""
    ws = _open_worksheet()
    out = df.copy()
    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")
    out = out.fillna("")
    values = [COLUMNS] + out[COLUMNS].astype(str).values.tolist()
    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")
