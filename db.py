# -*- coding: utf-8 -*-
"""
Supabase(PostgreSQL) 저장소.

접속 정보는 코드에 적지 않는다. `.streamlit/secrets.toml` 의 [db] url 을 쓴다.
  (Streamlit Cloud 에서는 앱 설정 화면의 Secrets 에 같은 내용을 넣는다)

DB 가 없으면 모든 함수가 조용히 비활성화되고, 앱은 기존처럼 엑셀 파일로 동작한다.
따라서 사무실 PC(파일 모드)와 클라우드(DB 모드) 양쪽에서 같은 코드가 돈다.

표 구성
  cost_master     상품코드 → 원가        (마진율표에서 동기화)
  fee_master      상품코드 → 셀러수수료율 (마진율표에서 동기화)
  cost_override   수동 보정 원가
  fee_override    수동 보정 요율
  card_fee        월별 카드수수료 (결제수단별)
  sample_cost     월별 샘플비용
  channel_config  채널 설정
  app_setting     기타 설정(기준월 등)
"""

import os

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:                                   # psycopg 미설치 = 파일 모드
    psycopg = None
    dict_row = None


SCHEMA = """
create table if not exists cost_master (
    code       text primary key,
    cost       numeric not null,
    source     text,
    updated_at timestamptz not null default now()
);
create table if not exists fee_master (
    code       text primary key,
    fee        numeric not null,
    source     text,
    updated_at timestamptz not null default now()
);
create table if not exists cost_override (
    code       text primary key,
    cost       numeric not null,
    force_apply boolean not null default false,
    item_name  text,
    note       text,
    updated_at timestamptz not null default now()
);
create table if not exists fee_override (
    code       text primary key,
    fee        numeric not null,
    item_name  text,
    note       text,
    updated_at timestamptz not null default now()
);
create table if not exists card_fee (
    period text not null,
    method text not null,
    amount numeric not null default 0,
    rate   numeric not null default 0,
    primary key (period, method)
);
create table if not exists sample_cost (
    period text primary key,
    item   text not null default '자체샘플+이벤트지원',
    amount numeric not null default 0
);
create table if not exists channel_config (
    name       text primary key,
    sort_order int  not null default 0,
    file_name  text,
    seller     boolean not null default false,
    card       boolean not null default false,
    sample     boolean not null default false,
    fee_base   numeric not null default 0
);
create table if not exists app_setting (
    key   text primary key,
    value text
);
create table if not exists sync_log (
    id         bigserial primary key,
    kind       text,
    detail     text,
    synced_at  timestamptz not null default now()
);
"""


# ---------------------------------------------------------------- 접속
# 아직 값을 안 채운 예시 주소를 걸러낸다 (이걸 그냥 두면 앱이 DB 모드로 잘못 켜진다)
PLACEHOLDERS = ("[YOUR-PASSWORD]", "YOUR-PASSWORD", "비밀번호@", "postgresql://...",
                "xxxx", "여기에")


def looks_filled(url):
    """접속 문자열이 실제 값으로 채워졌는지"""
    if not url or not str(url).strip():
        return False
    u = str(url)
    return not any(p in u for p in PLACEHOLDERS)


def get_url():
    """secrets.toml → 환경변수 순으로 접속 문자열을 찾는다. 없거나 미완성이면 None."""
    cand = None
    try:
        import streamlit as st
        if "db" in st.secrets and st.secrets["db"].get("url"):
            cand = st.secrets["db"]["url"]
    except Exception:
        pass
    cand = cand or os.environ.get("PROFIT_DB_URL")
    return cand if looks_filled(cand) else None


def url_problem():
    """접속 문자열에 문제가 있으면 사람이 읽을 안내문을 돌려준다. 없으면 None."""
    raw = None
    try:
        import streamlit as st
        if "db" in st.secrets:
            raw = st.secrets["db"].get("url")
    except Exception:
        pass
    raw = raw or os.environ.get("PROFIT_DB_URL")
    if not raw:
        return None
    if not looks_filled(raw):
        return ("접속 문자열에 예시 값이 그대로 남아 있습니다. "
                "`[YOUR-PASSWORD]` 를 실제 DB 비밀번호로 바꿔주세요.")
    if "db." in raw and ".supabase.co" in raw and ":5432" in raw:
        return ("직접 연결(5432) 주소입니다. 이 주소는 IPv6 전용이라 "
                "Streamlit Cloud 에서 접속되지 않습니다. "
                "Supabase 의 **Connection pooling** 주소(6543)를 쓰세요.")
    return None


def enabled():
    return psycopg is not None and bool(get_url())


def connect():
    """열린 연결 반환. 사용 후 close() 하거나 with 로 감쌀 것."""
    url = get_url()
    if not url:
        raise RuntimeError("DB 접속 정보가 없습니다. .streamlit/secrets.toml 의 [db] url 을 채우세요.")
    if psycopg is None:
        raise RuntimeError("psycopg 가 설치되어 있지 않습니다.  pip install \"psycopg[binary]\"")
    return psycopg.connect(url, row_factory=dict_row, connect_timeout=10)


def init_schema():
    with connect() as cn, cn.cursor() as cur:
        cur.execute(SCHEMA)
        cn.commit()


def ping():
    """(성공여부, 메시지)"""
    try:
        with connect() as cn, cn.cursor() as cur:
            cur.execute("select version()")
            v = cur.fetchone()["version"]
        return True, v.split(",")[0]
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------- 읽기
def _rows(sql, args=None):
    with connect() as cn, cn.cursor() as cur:
        cur.execute(sql, args or ())
        return cur.fetchall()


def load_lookups():
    """DB 에 저장된 기준 원가·요율 → (cost, fee, origin)"""
    cost, fee, origin = {}, {}, {}
    for r in _rows("select code, cost, source from cost_master"):
        cost[r["code"]] = float(r["cost"])
        origin[r["code"]] = r["source"] or "DB"
    for r in _rows("select code, fee from fee_master"):
        fee[r["code"]] = float(r["fee"])
    return cost, fee, origin


def load_overrides():
    """(원가보정 {code:(값, 우선적용)}, 요율보정 {code:율})"""
    ov = {r["code"]: (float(r["cost"]), bool(r["force_apply"]))
          for r in _rows("select code, cost, force_apply from cost_override")}
    fv = {r["code"]: float(r["fee"]) for r in _rows("select code, fee from fee_override")}
    return ov, fv


def load_cards():
    """{기준월: [(결제수단, 금액, 요율), ...]}"""
    out = {}
    for r in _rows("select period, method, amount, rate from card_fee "
                   "order by period, method"):
        out.setdefault(r["period"], []).append(
            (r["method"], float(r["amount"]), float(r["rate"])))
    return out


def load_samples():
    """{기준월: 금액}"""
    return {r["period"]: float(r["amount"])
            for r in _rows("select period, amount from sample_cost")}


def load_channels():
    out = []
    for r in _rows("select name, file_name, seller, card, sample, fee_base "
                   "from channel_config order by sort_order, name"):
        out.append({"name": r["name"], "file": r["file_name"] or "",
                    "seller": r["seller"], "card": r["card"],
                    "sample": r["sample"], "fee_base": float(r["fee_base"])})
    return out


def get_setting(key, default=None):
    r = _rows("select value from app_setting where key = %s", (key,))
    return r[0]["value"] if r else default


def last_sync():
    r = _rows("select kind, detail, synced_at from sync_log "
              "order by synced_at desc limit 1")
    return r[0] if r else None


# ---------------------------------------------------------------- 쓰기
def save_setting(key, value):
    with connect() as cn, cn.cursor() as cur:
        cur.execute("insert into app_setting(key, value) values (%s, %s) "
                    "on conflict (key) do update set value = excluded.value",
                    (key, str(value)))
        cn.commit()


def replace_lookups(cost, fee, origin=None):
    """마진율표 전체를 통째로 갈아끼운다 (동기화용)."""
    origin = origin or {}
    with connect() as cn, cn.cursor() as cur:
        cur.execute("truncate cost_master")
        cur.execute("truncate fee_master")
        if cost:
            cur.executemany(
                "insert into cost_master(code, cost, source) values (%s, %s, %s)",
                [(k, v, origin.get(k)) for k, v in cost.items()])
        if fee:
            cur.executemany(
                "insert into fee_master(code, fee, source) values (%s, %s, %s)",
                [(k, v, origin.get(k)) for k, v in fee.items()])
        cn.commit()


def upsert_overrides(cost_ov, names=None, note=None):
    """{code: (값, 우선적용)} 또는 {code: 값} 저장"""
    names = names or {}
    rows = []
    for k, v in (cost_ov or {}).items():
        val, force = v if isinstance(v, (tuple, list)) else (v, False)
        nm = names.get(k)
        rows.append((k, float(val), bool(force),
                     nm[1] if isinstance(nm, (tuple, list)) else nm, note))
    if not rows:
        return 0
    with connect() as cn, cn.cursor() as cur:
        cur.executemany(
            "insert into cost_override(code, cost, force_apply, item_name, note) "
            "values (%s, %s, %s, %s, %s) on conflict (code) do update set "
            "cost = excluded.cost, force_apply = excluded.force_apply, "
            "item_name = coalesce(excluded.item_name, cost_override.item_name), "
            "note = excluded.note, updated_at = now()", rows)
        cn.commit()
    return len(rows)


def upsert_fee_overrides(fee_ov, note=None):
    rows = [(k, float(v), note) for k, v in (fee_ov or {}).items()]
    if not rows:
        return 0
    with connect() as cn, cn.cursor() as cur:
        cur.executemany(
            "insert into fee_override(code, fee, note) values (%s, %s, %s) "
            "on conflict (code) do update set fee = excluded.fee, "
            "note = excluded.note, updated_at = now()", rows)
        cn.commit()
    return len(rows)


def delete_overrides(codes):
    if not codes:
        return 0
    with connect() as cn, cn.cursor() as cur:
        cur.execute("delete from cost_override where code = any(%s)", (list(codes),))
        n = cur.rowcount
        cn.commit()
    return n


def save_cards(period, cards):
    """그 달의 카드수수료만 갈아끼운다. cards: [(결제수단, 금액, 요율), ...]"""
    with connect() as cn, cn.cursor() as cur:
        cur.execute("delete from card_fee where period = %s", (period,))
        if cards:
            cur.executemany(
                "insert into card_fee(period, method, amount, rate) values (%s,%s,%s,%s)",
                [(period, m, float(a), float(r)) for m, a, r in cards])
        cn.commit()


def save_samples(by_month):
    """{기준월: 금액} 저장 (해당 달만 갱신)"""
    rows = [(m, float(a)) for m, a in (by_month or {}).items()]
    if not rows:
        return
    with connect() as cn, cn.cursor() as cur:
        cur.executemany(
            "insert into sample_cost(period, amount) values (%s, %s) "
            "on conflict (period) do update set amount = excluded.amount", rows)
        cn.commit()


def save_channels(channels):
    with connect() as cn, cn.cursor() as cur:
        cur.execute("truncate channel_config")
        cur.executemany(
            "insert into channel_config(name, sort_order, file_name, seller, card, sample, fee_base) "
            "values (%s,%s,%s,%s,%s,%s,%s)",
            [(c["name"], i, c.get("file", ""), bool(c.get("seller")), bool(c.get("card")),
              bool(c.get("sample")), float(c.get("fee_base") or 0))
             for i, c in enumerate(channels)])
        cn.commit()


def log_sync(kind, detail):
    with connect() as cn, cn.cursor() as cur:
        cur.execute("insert into sync_log(kind, detail) values (%s, %s)", (kind, detail))
        cn.commit()
