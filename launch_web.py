# -*- coding: utf-8 -*-
"""
웹앱 실행 도우미.

웹앱실행.bat 이 이 파일을 호출한다.
(배치 파일에는 한글을 넣지 않는다 — cmd.exe 가 한글이 포함된 .bat 을
 잘못 읽어 명령어가 깨지는 문제가 있어서, 안내 문구는 전부 여기서 출력한다.)

하는 일
  1) streamlit / openpyxl / pandas 설치 확인
  2) 이미 실행 중이면 브라우저만 열고 끝냄 (중복 실행 방지)
  3) Streamlit 최초 실행 시 나오는 이메일 질문 끄기
  4) 서버를 띄우고 **접속 가능해진 뒤에** 브라우저를 연다
     (브라우저가 먼저 열려 '사이트에 연결할 수 없음' 이 뜨는 것을 막는다)
"""

import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

BASE = Path(__file__).resolve().parent
APP = BASE / "app.py"
PORT = 8501
URL = "http://localhost:{}".format(PORT)
WAIT_SECONDS = 90

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def say(*args):
    print(*args, flush=True)


def port_open(port, host="127.0.0.1", timeout=0.4):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        return s.connect_ex((host, port)) == 0


def ensure_credentials():
    """이메일 입력 프롬프트에서 멈추지 않도록 빈 값으로 만들어 둔다.
       BOM 이 있으면 Streamlit 이 손상된 파일로 보고 지워버리므로 BOM 없이 쓴다."""
    d = Path(os.path.expanduser("~")) / ".streamlit"
    f = d / "credentials.toml"
    try:
        d.mkdir(parents=True, exist_ok=True)
        need = True
        if f.exists():
            raw = f.read_bytes()
            need = raw.startswith(b"\xef\xbb\xbf") or b"email" not in raw
        if need:
            with open(f, "w", encoding="utf-8", newline="") as fp:
                fp.write('[general]\nemail = ""\n')
    except OSError:
        pass          # 못 만들어도 치명적이지 않음 (프롬프트만 뜸)


def banner(extra=""):
    say()
    say("  " + "=" * 58)
    say("   이익률 계산기")
    say()
    say("   주소 : " + URL)
    if extra:
        say("   " + extra)
    say()
    say("   ** 이 창을 닫으면 앱이 종료됩니다 **")
    say("  " + "=" * 58)
    say()


def main():
    missing = []
    for mod in ("streamlit", "openpyxl", "pandas"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    if missing:
        say()
        say("  [설치 필요] 아래 프로그램이 없습니다: " + ", ".join(missing))
        say()
        say("  명령 프롬프트에서 아래를 실행한 뒤 다시 시도하세요:")
        say()
        say("      pip install " + " ".join(missing))
        say()
        return 1

    if not APP.exists():
        say()
        say("  [오류] app.py 를 찾을 수 없습니다.")
        say("         폴더: {}".format(BASE))
        say()
        return 1

    if port_open(PORT):
        banner("이미 실행 중이라 브라우저만 엽니다.")
        webbrowser.open(URL)
        return 0

    ensure_credentials()

    say()
    say("  이익률 계산기를 시작합니다. 잠시만 기다려 주세요...")

    # headless 로 띄우고 우리가 직접 브라우저를 연다 (타이밍을 맞추기 위해)
    cmd = [sys.executable, "-m", "streamlit", "run", str(APP),
           "--server.port", str(PORT),
           "--server.headless", "true",
           "--browser.gatherUsageStats", "false"]
    proc = subprocess.Popen(cmd, cwd=str(BASE))

    opened = False
    deadline = time.time() + WAIT_SECONDS
    while time.time() < deadline:
        if proc.poll() is not None:
            say()
            say("  [오류] 서버가 시작되지 못했습니다. 위의 메시지를 확인하세요.")
            say()
            return proc.returncode or 1
        if port_open(PORT):
            banner()
            webbrowser.open(URL)
            opened = True
            break
        time.sleep(0.5)

    if not opened:
        banner("자동으로 열리지 않았습니다. 위 주소를 직접 입력하세요.")

    try:
        return proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        import traceback
        traceback.print_exc()
        code = 1
    if code:
        try:
            input("엔터를 누르면 종료합니다...")
        except EOFError:
            pass
    sys.exit(code or 0)
