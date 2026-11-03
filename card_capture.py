# -*- coding: utf-8 -*-
"""
'매출내역 기간 통계' 화면 캡쳐 → 결제수단별 결제합계 읽기 (무료 글자인식 Tesseract)

열 제목 글자는 잘못 읽힐 수 있어서 위치로 열을 찾지 않는다.
대신 각 줄의 숫자를 왼쪽부터 늘어놓고
    결제합계 = 상품구매금액 + 배송비 − (사이에 있는 할인들)
이 맞아떨어지는 칸을 결제합계로 본다. 계산이 안 맞는 줄은 버린다.

서버에 tesseract 프로그램이 있어야 한다.
  - Streamlit Cloud: packages.txt 의 tesseract-ocr, tesseract-ocr-kor
  - Windows: https://github.com/UB-Mannheim/tesseract/wiki 에서 설치
"""

import difflib
import os
import re
import shutil

try:
    import pytesseract
    from PIL import Image, ImageOps
except ImportError:                     # 설치 안 됨 → available() 가 False
    pytesseract = None

_WIN_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]

NUM_RE = re.compile(r"^\d{1,3}([,.]\d{3})+$|^\d+$")


def _setup():
    if pytesseract is None:
        return False
    if shutil.which("tesseract"):
        return True
    for p in _WIN_PATHS:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return True
    return False


def available():
    """글자인식을 쓸 수 있는지"""
    return _setup()


def _langs():
    try:
        have = set(pytesseract.get_languages(config=""))
    except Exception:
        have = set()
    return "kor+eng" if "kor" in have else "eng"


def _words(img):
    """[(글자, x, y중심, 높이, 오른쪽x)]"""
    g = ImageOps.grayscale(img)
    if g.width < 2000:                  # 작은 캡쳐는 키워야 숫자를 잘 읽는다
        k = 2000 / g.width
        g = g.resize((int(g.width * k), int(g.height * k)), Image.LANCZOS)
    d = pytesseract.image_to_data(g, lang=_langs(), config="--psm 11",
                                  output_type=pytesseract.Output.DICT)
    out = []
    for i, t in enumerate(d["text"]):
        t = (t or "").strip()
        if not t:
            continue
        x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
        out.append((t, x, y + h / 2.0, h, x + w))
    return out


def _num(t):
    t = t.strip().strip("|").strip()
    if not NUM_RE.match(t):
        return None
    return int(re.sub(r"[,.]", "", t))


def _rows(words):
    """y 위치가 비슷한 글자끼리 한 줄로 묶는다"""
    if not words:
        return []
    hs = sorted(w[3] for w in words)
    tol = max(6.0, hs[len(hs) // 2] * 0.7)
    rows = []
    for w in sorted(words, key=lambda w: w[2]):
        if rows and abs(w[2] - rows[-1]["y"]) <= tol:
            r = rows[-1]
            r["w"].append(w)
            r["y"] = sum(x[2] for x in r["w"]) / len(r["w"])
        else:
            rows.append({"y": w[2], "w": [w]})
    return rows


def _pay_total(nums):
    """[건수, 상품구매금액, 배송비, 할인..., 결제합계, ...] 에서 결제합계 위치"""
    if len(nums) < 4:
        return None
    goods, ship = nums[1], nums[2]
    disc = 0
    for j in range(3, len(nums)):
        if nums[j] > 0 and nums[j] == goods + ship - disc:
            return j
        disc += nums[j]
    return None


def read_capture(img):
    """캡쳐 이미지 → {"rows": [{"label", "total", "count"}], "subtotal": 금액 또는 None}
       label 은 글자인식으로 읽은 결제수단 이름(틀릴 수 있음)."""
    if not _setup():
        raise RuntimeError("글자인식 프로그램(tesseract)이 설치되어 있지 않습니다.")
    found = []
    for r in _rows(_words(img)):
        ws = sorted(r["w"], key=lambda w: w[1])
        nums, first_x = [], None
        for t, x, _y, _h, _r in ws:
            v = _num(t)
            if v is not None:
                nums.append(v)
                if first_x is None:
                    first_x = x
        j = _pay_total(nums)
        if j is None:
            continue
        label = " ".join(t for t, x, *_ in ws if _num(t) is None and x < first_x)
        found.append({"label": re.sub(r"[^0-9A-Za-z가-힣]", "", label),
                      "total": float(nums[j]), "count": nums[0]})

    # 다른 줄들의 합과 같은 줄은 소계로 본다 (글자를 못 읽어도 알 수 있게)
    subtotal = None
    for i, f in enumerate(found):
        rest = sum(g["total"] for k, g in enumerate(found) if k != i)
        if len(found) > 2 and abs(f["total"] - rest) < 1 or "소계" in f["label"] or "합계" in f["label"]:
            subtotal = f["total"]
            found.pop(i)
            break
    return {"rows": found, "subtotal": subtotal}


ALIASES = {"카드": "신용카드", "신용": "신용카드", "무통장": "무통장입금", "입금": "무통장입금",
           "payco": "페이코", "toss": "토스", "토스페이": "토스"}


def match_method(label, methods):
    """읽은 이름 → 화면의 결제수단 이름 (못 맞추면 None)"""
    if not label:
        return None
    low = label.lower()
    for m in methods:
        if m == label or m in label or (len(label) >= 2 and label in m):
            return m
    for k, v in ALIASES.items():
        if k in low and v in methods:
            return v
    got = difflib.get_close_matches(label, methods, n=1, cutoff=0.4)
    return got[0] if got else None
