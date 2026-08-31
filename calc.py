# -*- coding: utf-8 -*-
"""
이익률 자동 계산 — 명령줄(배치) 버전.  계산 로직은 엔진.py 에 있다.

사용법
  1) 이카운트에서 채널별 매출을 다운로드해 `입력/` 폴더에 넣는다.
     (파일명은 `입력/설정.xlsx`의 [채널설정] 시트에 적힌 이름과 같아야 함)
  2) `입력/설정.xlsx`에 이번 달 기준월·카드수수료·샘플비용을 입력한다.
  3) 실행.bat 을 더블클릭하거나  python calc.py  실행
  4) `결과/YYYY-MM 이익률.xlsx` 생성

웹 화면으로 쓰려면 `웹앱실행.bat` 을 실행하세요.
"""

import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

try:
    import engine as E
except ImportError:
    sys.exit("openpyxl 이 필요합니다.  실행:  pip install openpyxl")

MASTER = BASE / "이익률 마스터.xlsx"
INPUT_DIR = BASE / "입력"
CONFIG = INPUT_DIR / "설정.xlsx"
OVERRIDE = BASE / "원가보정.xlsx"
OUT_DIR = BASE / "결과"


def log(msg):
    print(msg, flush=True)


def load_config():
    """입력/설정.xlsx → (채널목록, 카드내역, 카드합계, 샘플내역, 샘플합계, 기준월)"""
    if not CONFIG.exists():
        sys.exit("[오류] 설정 파일이 없습니다: 입력/설정.xlsx\n"
                 "       초기설정.bat 을 먼저 실행하세요.")
    wb = E.open_wb(CONFIG, data_only=True)

    sources = []
    years = E.DEFAULT_YEARS
    if "기준파일" in wb.sheetnames:
        ws = wb["기준파일"]
        hr, col = E.find_header(ws, {"name": ["이름"], "path": ["파일경로"]})
        if hr is not None:
            for _f, _al in (("use", ["사용"]), ("years", ["최근연도수"])):
                _, opt = E.find_header(ws, {_f: _al})
                if opt:
                    col.update(opt)
            for row in ws.iter_rows(min_row=hr + 1, values_only=True):
                p = str(E._pick(row, col, "path") or "").strip()
                if not p:
                    continue
                if "use" in col and not E.is_yes(E._pick(row, col, "use")):
                    continue
                nm = str(E._pick(row, col, "name") or "").strip() or Path(p).stem
                sources.append((nm, p))
                if "years" in col:
                    _n = E.num(E._pick(row, col, "years"))
                    if _n >= 1:
                        years = int(_n)

    channels = []
    ws = wb["채널설정"]
    hr, col = E.find_header(ws, {
        "name": ["채널명"], "file": ["파일명"],
        "seller": ["셀러수수료적용"], "card": ["카드수수료적용"], "sample": ["샘플비용적용"],
    })
    if hr is None:
        sys.exit("[오류] 설정.xlsx [채널설정] 시트 열 이름을 확인하세요.")
    _, opt = E.find_header(ws, {"base": ["셀러수수료기본율"]})
    if opt:
        col.update(opt)
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        nm = E._pick(row, col, "name")
        fn = str(E._pick(row, col, "file") or "").strip()
        if not nm or not str(nm).strip():
            continue
        nm = str(nm).strip()
        if nm.startswith("*") or not fn:          # 안내문구 줄 건너뛰기
            continue
        channels.append({
            "name": nm,
            "file": fn,
            "seller": E.is_yes(E._pick(row, col, "seller")),
            "card": E.is_yes(E._pick(row, col, "card")),
            "sample": E.is_yes(E._pick(row, col, "sample")),
            "fee_base": (E.rate(E._pick(row, col, "base")) or 0.0) if "base" in col else 0.0,
        })

    period = ""
    if "기간" in wb.sheetnames:
        for row in wb["기간"].iter_rows(values_only=True):
            if row and E.norm(row[0]) == "기준월":
                for v in row[1:]:
                    if v is not None and str(v).strip():
                        period = str(v).strip()
                        break

    def same_month(v):
        """기준월 열이 없으면 전부, 있으면 그 달만"""
        return period == "" or E.norm(v) == E.norm(period)

    raw_cards = []
    ws = wb["카드수수료"]
    hr, col = E.find_header(ws, {"method": ["결제수단"],
                                 "amount": ["정상금액", "결제금액", "금액"],
                                 "rate": ["수수료율", "수수료"]})
    if hr is not None:
        _, opt = E.find_header(ws, {"month": ["기준월"]})
        if opt:
            col.update(opt)
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            m = E._pick(row, col, "method")
            if not m or not str(m).strip() or E.norm(m) == "합계":
                continue
            if "month" in col and not same_month(E._pick(row, col, "month")):
                continue
            raw_cards.append((str(m).strip(), E._pick(row, col, "amount"), E._pick(row, col, "rate")))
    cards, card_total = E.card_rows_total(raw_cards)

    sample, sample_items = 0.0, []
    ws = wb["샘플비용"]
    hr, col = E.find_header(ws, {"item": ["항목"], "amount": ["금액"]})
    if hr is not None:
        _, opt = E.find_header(ws, {"month": ["기준월"]})
        if opt:
            col.update(opt)
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            it = E._pick(row, col, "item")
            if not it or not str(it).strip() or E.norm(it) == "합계":
                continue
            if "month" in col and not same_month(E._pick(row, col, "month")):
                continue
            a = E.num(E._pick(row, col, "amount"))
            sample_items.append((str(it).strip(), a))
            sample += a

    wb.close()
    return sources, years, channels, cards, card_total, sample_items, sample, period


def resolve(stem):
    """입력/ 에서 확장자가 달라도 찾아준다"""
    p = INPUT_DIR / stem
    if p.exists():
        return p
    for c in sorted(INPUT_DIR.glob(Path(stem).stem + ".*")):
        if c.suffix.lower() in (".xlsx", ".xlsm", ".csv") and not c.name.startswith("~$"):
            return c
    return None


def main():
    log("=" * 64)
    log("  이익률 계산")
    log("=" * 64)

    log("\n[1/4] 설정 읽는 중...")
    sources, years, channels, cards, card_total, sample_items, sample, period = load_config()
    if not sources:
        sources = [("마스터", str(MASTER))]

    log("\n[2/4] 기준 파일 읽는 중...")
    resolved, notes = E.expand_sources(sources, BASE, years)
    for n in notes:
        log("  " + n)
    if not resolved:
        sys.exit("[오류] 읽을 기준 파일이 없습니다. 입력/설정.xlsx [기준파일] 시트를 확인하세요.")
    try:
        cost, fee, conflicts, report, origin = E.load_lookups(
            resolved, E.month_cutoff(period))
    except ValueError as e:
        sys.exit("[오류] " + str(e))
    for r in report:
        log("  {:<12} [{}] 원가 {:>6,}건 (신규 {:>5,}) / 요율 {:>5,}건".format(
            r["label"], r["format"], r["cost"], r["cost_new"], r["fee"]))
    log("  -> 합계 원가 {:,}건 / 셀러수수료 {:,}건".format(len(cost), len(fee)))
    override, fee_override = E.load_overrides(OVERRIDE if OVERRIDE.exists() else None)
    log("  보정 — 원가 {:,}건 / 셀러수수료 {:,}건".format(len(override), len(fee_override)))

    if not period:
        period = datetime.now().strftime("%Y-%m")
    log("  기준월 {} / 채널 {}개".format(period, len(channels)))
    log("  카드수수료 {:,.0f}원 / 샘플비용 {:,.0f}원".format(card_total, sample))

    log("\n[3/4] 매출 파일 계산 중...")
    results, errs = [], []
    for ch in channels:
        path = resolve(ch["file"])
        if path is None:
            errs.append("  [건너뜀] {}: 입력/{} 없음".format(ch["name"], ch["file"]))
            continue
        try:
            rows = E.read_sales(path, path.name)
        except Exception as e:
            errs.append("  [오류] {} ({}): {}".format(ch["name"], path.name, e))
            continue
        res = E.calc_channel(ch, rows, cost, fee, override, fee_override, origin)
        results.append(res)
        nfee = sum(1 for g in res["goods"] if g["fee_missing"])
        log("  {:<12} 매출 {:>15,.0f}   원가 {:>15,.0f}   원가미매칭 {:>3}건   요율기본값 {:>3}건".format(
            ch["name"], res["sales"], res["cost"], len(res["missing"]), nfee))

    for e in errs:
        log(e)
    if not results:
        sys.exit("\n계산할 매출 파일이 없습니다. 입력/ 폴더를 확인하세요.")

    log("\n[4/4] 결과 저장 중...")
    total = E.apply_totals(results, card_total, sample)
    wb = E.build_report(results, total, cards, card_total, sample_items, sample, conflicts, period)
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / "{} 이익률.xlsx".format(period)
    try:
        wb.save(out)
    except PermissionError:
        sys.exit("[오류] '{}' 를 저장할 수 없습니다. 엑셀에서 열려 있으면 닫고 다시 실행하세요.".format(out.name))

    log("\n" + "=" * 64)
    log("  총매출(배송포함) : {:>18,.0f} 원".format(total["gross"]))
    log("  총이익액         : {:>18,.0f} 원".format(total["profit"]))
    log("  이익률           : {:>17.2f} %".format(total["margin"] * 100))
    log("=" * 64)
    nmiss = sum(len(r["missing"]) for r in results)
    if nmiss:
        miss_amt = sum(x[4] for x in E.missing_rows(results))
        log("\n  [주의] 원가 미매칭 {}건 / 판매액 {:,.0f}원 — [미매칭] 시트를 확인하세요.".format(nmiss, miss_amt))
    log("\n  저장 완료: 결과\\{}\n".format(out.name))


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code:
            log(str(e.code))
    except Exception as e:
        import traceback
        traceback.print_exc()
        log("\n[오류] {}".format(e))
    try:
        if sys.stdin is not None and sys.stdin.isatty():
            input("엔터를 누르면 종료합니다...")
    except Exception:
        pass
