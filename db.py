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
create table if not exists monthly_result (
    period     text not null,
    channel    text not null,
    sales      numeric not null default 0,
    cost       numeric not null default 0,
    fee        numeric not null default 0,
    card       numeric not null default 0,
    sample     numeric not null default 0,
    profit     numeric not null default 0,
    delivery   numeric not null default 0,
    saved_at   timestamptz not null default now(),
    primary key (period, channel)
);
create table if not exists cost_candidate (
    code       text not null,
    cost       numeric not null,
    source     text,
    eff_date   date,
    priority   int not null default 0,
    row_no     int not null default 0,
    primary key (code, priority, row_no)
);
create index if not exists cost_candidate_code_idx on cost_candidate (code);
create table if not exists monthly_report (
    period    text primary key,
    filename  text not null,
    content   bytea not null,
    summary   text,
    saved_at  timestamptz not null default now()
);
create table if not exists monthly_input (
    period    text not null,
    channel   text not null,
    filename  text not null,
    content   bytea not null,
    saved_at  timestamptz not null default now(),
    primary key (period, channel)
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
    # Supabase 풀러(6543, 트랜잭션 모드)는 prepared statement 를 지원하지 않는다.
    # prepare_threshold=None 으로 꺼야 executemany 가 동작한다.
    return psycopg.connect(url, row_factory=dict_row, connect_timeout=10,
                           prepare_threshold=None)


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
    order = {m: i for i, (m, _r) in enumerate(
        [("신용카드", 0), ("무통장입금", 0), ("페이코", 0), ("토스", 0)])}
    for r in sorted(_rows("select period, method, amount, rate from card_fee"),
                    key=lambda x: (x["period"], order.get(x["method"], 99), x["method"])):
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


def save_month_result(period, results, total):
    """그 달의 계산 결과를 통째로 갈아끼운다 (월별 추이 그래프용)"""
    rows = [(period, r["ch"]["name"], r["sales"], r["cost"], r["fee"],
             r.get("card", 0), r.get("sample", 0), r["profit"], r["delivery_sales"])
            for r in results]
    rows.append((period, "전체", total["sales"], total["cost"], total["fee"],
                 total["card"], total["sample"], total["profit"], total["delivery"]))
    with connect() as cn, cn.cursor() as cur:
        cur.execute("delete from monthly_result where period = %s", (period,))
        cur.executemany(
            "insert into monthly_result(period, channel, sales, cost, fee, card, "
            "sample, profit, delivery) values (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            [(p, c, float(a), float(b), float(d), float(e), float(f), float(g), float(h))
             for p, c, a, b, d, e, f, g, h in rows])
        cn.commit()
    return len(rows)


def load_month_results():
    """[{period, channel, sales, cost, profit, ...}, ...] 기간 오름차순"""
    return [{k: (float(v) if k not in ("period", "channel", "saved_at") else v)
             for k, v in r.items() if k != "saved_at"}
            for r in _rows("select period, channel, sales, cost, fee, card, sample, "
                           "profit, delivery from monthly_result order by period, channel")]


def save_candidates(rows):
    """마진율표의 모든 후보 단가를 보관.
       rows: [(code, cost, source, eff_date(str|None), priority, row_no), ...]"""
    with connect() as cn, cn.cursor() as cur:
        cur.execute("truncate cost_candidate")
        if rows:
            cur.executemany(
                "insert into cost_candidate(code, cost, source, eff_date, priority, row_no) "
                "values (%s,%s,%s,%s,%s,%s)", rows)
        cn.commit()
    return len(rows)


def load_lookups_asof(period=None):
    """기준월까지 기재된 단가만으로 조회표를 만든다.
       규칙: 앞선 파일(priority 작은 값) 우선 → 같은 파일 안에서는 기재날짜 최신.
       후보표가 비어 있으면 기존 cost_master 로 물러난다."""
    if not _rows("select 1 from cost_candidate limit 1"):
        return load_lookups()
    cutoff = None
    if period and len(str(period)) >= 7:
        y, m = int(str(period)[:4]), int(str(period)[5:7])
        y2, m2 = (y + 1, 1) if m == 12 else (y, m + 1)
        cutoff = "{:04d}-{:02d}-01".format(y2, m2)      # 다음 달 1일 미만
    sql = ("select distinct on (code) code, cost, source from cost_candidate "
           "{} order by code, priority asc, eff_date desc nulls last, row_no asc")
    where = "where eff_date is null or eff_date < %s" if cutoff else ""
    rs = _rows(sql.format(where), (cutoff,) if cutoff else None)
    cost = {r["code"]: float(r["cost"]) for r in rs}
    origin = {r["code"]: (r["source"] or "DB") for r in rs}
    fee = {r["code"]: float(r["fee"]) for r in _rows("select code, fee from fee_master")}
    return cost, fee, origin


def save_month_report(period, filename, content, summary=""):
    """그 달의 결과 엑셀을 통째로 보관 — 나중에 다시 계산하지 않고 내려받을 수 있게"""
    with connect() as cn, cn.cursor() as cur:
        cur.execute(
            "insert into monthly_report(period, filename, content, summary) "
            "values (%s,%s,%s,%s) on conflict (period) do update set "
            "filename = excluded.filename, content = excluded.content, "
            "summary = excluded.summary, saved_at = now()",
            (period, filename, content, summary))
        cn.commit()


def load_month_report(period):
    """(파일명, bytes, 요약, 저장시각) 또는 None"""
    r = _rows("select filename, content, summary, saved_at from monthly_report "
              "where period = %s", (period,))
    if not r:
        return None
    r = r[0]
    return (r["filename"], bytes(r["content"]), r["summary"] or "", r["saved_at"])


def list_month_reports():
    """[(기준월, 파일명, 요약, 저장시각), ...] 최신월 순"""
    return [(r["period"], r["filename"], r["summary"] or "", r["saved_at"])
            for r in _rows("select period, filename, summary, saved_at "
                           "from monthly_report order by period desc")]


def save_month_files(period, files):
    """그 달에 쓴 이카운트 매출 파일을 통째로 보관 (페이지를 나갔다 와도 결과 복원용).
       files: {채널명: (파일명, bytes)}"""
    if not files:
        return 0
    with connect() as cn, cn.cursor() as cur:
        cur.execute("delete from monthly_input where period = %s", (period,))
        cur.executemany(
            "insert into monthly_input(period, channel, filename, content) "
            "values (%s, %s, %s, %s)",
            [(period, ch, fn, data) for ch, (fn, data) in files.items()])
        cn.commit()
    return len(files)


def load_month_files(period):
    """{채널명: (파일명, bytes)}  없으면 {}"""
    out = {}
    for r in _rows("select channel, filename, content from monthly_input "
                   "where period = %s", (period,)):
        out[r["channel"]] = (r["filename"], bytes(r["content"]))
    return out


def months_with_input():
    return [r["period"] for r in
            _rows("select distinct period from monthly_input order by period")]


def log_sync(kind, detail):
    with connect() as cn, cn.cursor() as cur:
        cur.execute("insert into sync_log(kind, detail) values (%s, %s)", (kind, detail))
        cn.commit()
