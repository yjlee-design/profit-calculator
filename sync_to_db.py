# -*- coding: utf-8 -*-
"""
사무실 PC → Supabase 동기화.

클라우드 서버에는 G: 구글드라이브가 없다. 그래서 **구글드라이브가 보이는 이 PC에서**
마진율표를 읽어 DB 에 올려두고, 인터넷에 올린 앱은 DB 만 읽는다.

하는 일
  1) 입력/설정.xlsx 의 [기준파일] 경로에서 연도별 마진율표를 읽어 원가·요율을 DB 에 저장
  2) 원가보정.xlsx → DB
  3) 설정.xlsx 의 채널설정 / 월별 카드수수료 / 월별 샘플비용 → DB

실행:  동기화.bat  더블클릭   (또는  python sync_to_db.py)
"""

import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import engine as E          # noqa: E402
import db                   # noqa: E402
import calc                 # noqa: E402  (설정.xlsx 읽기 재사용)

OVERRIDE = BASE / "원가보정.xlsx"


def log(m):
    print(m, flush=True)


def main():
    log("=" * 64)
    log("  Supabase 동기화")
    log("=" * 64)

    if not db.enabled():
        log("\n[오류] DB 접속 정보가 없습니다.")
        log("       .streamlit/secrets.toml 을 만들고 [db] url 을 채우세요.")
        log("       (.streamlit/secrets.toml.example 을 복사해 쓰면 됩니다)")
        return 1

    ok, msg = db.ping()
    if not ok:
        log("\n[오류] DB 에 접속하지 못했습니다:\n       " + msg)
        return 1
    log("\n접속 성공 — " + msg)

    log("\n[1/5] 표 만드는 중...")
    db.init_schema()
    log("  완료")

    log("\n[2/5] 설정 읽는 중...")
    sources, years, channels, cards, card_total, sample_items, sample, period = calc.load_config()
    resolved, notes = E.expand_sources(sources, BASE, years)
    for n in notes:
        log("  " + n)
    if not resolved:
        log("\n[오류] 읽을 기준 파일이 없습니다.")
        return 1

    log("\n[3/5] 마진율표 읽는 중... (파일이 커서 잠시 걸립니다)")
    cost, fee, conflicts, report, origin = E.load_lookups(resolved)
    for r in report:
        log("  {:<12} 원가 {:>6,}건 (신규 {:>5,}) / 요율 {:>5,}건".format(
            r["label"], r["cost"], r["cost_new"], r["fee"]))
    log("  -> 합계 원가 {:,}건 / 셀러수수료 {:,}건".format(len(cost), len(fee)))

    log("\n[4/5] DB 에 올리는 중...")
    db.replace_lookups(cost, fee, origin)
    log("  기준 원가·요율 저장 완료")

    ov, fov = E.load_overrides(OVERRIDE if OVERRIDE.exists() else None)
    n1 = db.upsert_overrides(ov, note="원가보정.xlsx 동기화")
    n2 = db.upsert_fee_overrides(fov, note="원가보정.xlsx 동기화")
    log("  보정 저장 — 원가 {}건 / 요율 {}건".format(n1, n2))

    db.save_channels(channels)
    log("  채널설정 {}개 저장".format(len(channels)))

    # 월별 비용 — 설정.xlsx 전체를 훑어 모든 달을 올린다
    wb = E.open_wb(BASE / "입력" / "설정.xlsx", data_only=True)
    try:
        months_c, months_s = _read_all_months(wb)
    finally:
        wb.close()
    for p, rows in months_c.items():
        db.save_cards(p, rows)
    db.save_samples(months_s)
    log("  카드수수료 {}개월 / 샘플비용 {}개월 저장".format(len(months_c), len(months_s)))

    if period:
        db.save_setting("period", period)

    log("\n[5/5] 기록 남기는 중...")
    detail = "원가 {:,} / 요율 {:,} / {}".format(
        len(cost), len(fee), " + ".join(lb for lb, _f in resolved))
    db.log_sync("master", detail)
    log("  " + detail)

    log("\n" + "=" * 64)
    log("  동기화 완료  ({})".format(datetime.now().strftime("%Y-%m-%d %H:%M")))
    log("  이제 인터넷에 올린 앱에서도 같은 원가로 계산됩니다.")
    log("=" * 64 + "\n")
    return 0


def _read_all_months(wb):
    """설정.xlsx 의 카드수수료·샘플비용을 월 구분 없이 전부 읽는다"""
    cards, samples = {}, {}
    if "카드수수료" in wb.sheetnames:
        ws = wb["카드수수료"]
        hr, col = E.find_header(ws, {"method": ["결제수단"],
                                     "amount": ["정상금액", "결제금액", "금액"],
                                     "rate": ["수수료율", "수수료"]})
        if hr is not None:
            _, o = E.find_header(ws, {"month": ["기준월"]})
            if o:
                col.update(o)
            for row in ws.iter_rows(min_row=hr + 1, values_only=True):
                m = E._pick(row, col, "method")
                if not m or not str(m).strip() or E.norm(m) == "합계":
                    continue
                p = str(E._pick(row, col, "month") or "").strip()
                if not p:
                    continue
                cards.setdefault(p, []).append(
                    (str(m).strip(), E.num(E._pick(row, col, "amount")),
                     E.rate(E._pick(row, col, "rate")) or 0.0))
    if "샘플비용" in wb.sheetnames:
        ws = wb["샘플비용"]
        hr, col = E.find_header(ws, {"item": ["항목"], "amount": ["금액"]})
        if hr is not None:
            _, o = E.find_header(ws, {"month": ["기준월"]})
            if o:
                col.update(o)
            for row in ws.iter_rows(min_row=hr + 1, values_only=True):
                it = E._pick(row, col, "item")
                if not it or not str(it).strip() or E.norm(it) == "합계":
                    continue
                p = str(E._pick(row, col, "month") or "").strip()
                if not p:
                    continue
                samples[p] = samples.get(p, 0.0) + E.num(E._pick(row, col, "amount"))
    return cards, samples


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        import traceback
        traceback.print_exc()
        code = 1
    try:
        input("\n엔터를 누르면 닫힙니다...")
    except EOFError:
        pass
    sys.exit(code)
