# -*- coding: utf-8 -*-
"""
비밀번호 화면.

비밀번호는 코드에 적지 않는다. `.streamlit/secrets.toml` 의
    passwords = ["..."]        (여러 개 가능)
또는 환경변수 PROFIT_APP_PASSWORD 를 쓴다.

비밀번호를 아예 설정하지 않으면 잠그지 않는다(사무실 PC 단독 사용 상황).
단, 인터넷에 올린 경우에는 반드시 설정하도록 화면에 경고를 띄운다.
"""

import hmac
import os

import streamlit as st


def _allowed():
    """허용된 비밀번호 목록"""
    out = []
    try:
        v = st.secrets.get("passwords")
        if isinstance(v, str):
            out = [v]
        elif v:
            out = list(v)
        elif st.secrets.get("password"):
            out = [st.secrets["password"]]
    except Exception:
        pass
    env = os.environ.get("PROFIT_APP_PASSWORD")
    if env:
        out.append(env)
    return [str(p) for p in out if str(p).strip()]


def _is_cloud():
    """클라우드에서 돌고 있는지 대략 판단 (경고 문구용)"""
    return bool(os.environ.get("HOSTNAME", "").startswith("streamlit")
                or os.environ.get("STREAMLIT_SERVER_HEADLESS") == "true"
                or os.environ.get("STREAMLIT_RUNTIME_ENV"))


def require_password():
    """통과하면 True. 아니면 잠금 화면을 그리고 False (호출측은 st.stop())."""
    pws = _allowed()
    if not pws:
        if _is_cloud():
            st.error("비밀번호가 설정되지 않았습니다. "
                     "앱 설정의 Secrets 에 `passwords = [\"...\"]` 를 넣어 주세요.")
            return False
        return True                     # 내 PC 에서만 쓰는 경우 잠그지 않음

    if st.session_state.get("auth_ok"):
        return True

    st.markdown(
        "<div style='max-width:380px;margin:14vh auto 0;padding:2rem 1.9rem;"
        "background:#FFFFFF;border:1px solid #E2E8F0;border-radius:12px;"
        "box-shadow:0 1px 3px rgba(15,23,42,.05)'>"
        "<div style='font-size:1.15rem;font-weight:700;color:#0F172A'>📊 이익률 계산기</div>"
        "<div style='margin-top:.35rem;font-size:.84rem;color:#64748B'>"
        "사내 전용입니다. 비밀번호를 입력하세요.</div></div>",
        unsafe_allow_html=True)

    box = st.container()
    with box:
        c = st.columns([1, 1.6, 1])[1]
        with c:
            pw = st.text_input("비밀번호", type="password", key="pw_in",
                               label_visibility="collapsed",
                               placeholder="비밀번호")
            if st.button("들어가기", type="primary", width="stretch"):
                if any(hmac.compare_digest(pw, p) for p in pws):
                    st.session_state["auth_ok"] = True
                    st.session_state.pop("pw_in", None)
                    st.rerun()
                else:
                    st.session_state["pw_tries"] = st.session_state.get("pw_tries", 0) + 1
                    st.error("비밀번호가 맞지 않습니다.")
            if st.session_state.get("pw_tries", 0) >= 3:
                st.caption("비밀번호를 모르면 담당자에게 문의하세요.")
    return False


def logout_button(where=None):
    target = where or st.sidebar
    if st.session_state.get("auth_ok") and _allowed():
        if target.button("로그아웃", width="stretch"):
            st.session_state.pop("auth_ok", None)
            st.rerun()
