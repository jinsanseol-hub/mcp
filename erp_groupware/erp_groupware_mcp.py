"""
ERP 메일/일정 관리 MCP 서버
============================
사내 ERP(그룹웨어) 시스템에 접속하여 메일 송수신 및 일정 관리를
자동화하는 MCP(Model Context Protocol) 서버.

Playwright(웹 자동화) + FastMCP 기반.

제공 기능:
  [메일] 수신함 조회, 메일 상세 읽기, 메일 발송, 메일 검색
  [일정] 일정 조회, 일정 등록, 일정 수정, 일정 삭제

사용법:
  1. .env 파일에 자격증명 설정
  2. python erp_groupware_mcp.py          (stdio 모드 - Claude Desktop)
  3. python erp_groupware_mcp.py --http   (HTTP 모드 - 웹서비스, 포트 8000)

Claude Desktop 설정 (claude_desktop_config.json):
  {
    "mcpServers": {
      "erp_groupware": {
        "command": "python",
        "args": ["/path/to/erp_groupware_mcp.py"],
        "env": {
          "ERP_BASE_URL": "https://groupware.company.com",
          "ERP_USERNAME": "your_id",
          "ERP_PASSWORD": "your_pw"
        }
      }
    }
  }
"""

import os
import re
import sys
import json
import time
import shutil
import asyncio
import logging
import subprocess
from datetime import datetime, date as dt_date
from typing import Optional, List
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP, Context

# ============================================================
# 환경 설정
# ============================================================

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

ERP_BASE_URL = os.getenv("ERP_BASE_URL", "https://groupware.company.com")
ERP_USERNAME = os.getenv("ERP_USERNAME", "")
ERP_PASSWORD = os.getenv("ERP_PASSWORD", "")

# ERP_BASE_URL 에서 호스트 부분만 추출 (예: https://erp.boninfo.co.kr/#/login → https://erp.boninfo.co.kr)
ERP_HOST = ERP_BASE_URL.split("#")[0].rstrip("/")

# 모듈별 URL (더존 아마란스 해시 라우팅)
SCHEDULE_PAGE_URL = f"{ERP_HOST}/#/UE/UEA/UEA0000?specialLnb=Y&moduleCode=UE&menuCode=UEA&pageCode=UEA0000"

logger = logging.getLogger("erp_groupware_mcp")
logging.basicConfig(level=logging.INFO, stream=sys.stderr)

# DEBUG_SCREENSHOTS=1 환경변수 설정 시 디버그 스크린샷 저장 (기본 비활성)
DEBUG_SCREENSHOTS = os.getenv("DEBUG_SCREENSHOTS", "1") == "1"
SCREENSHOT_DIR = os.path.join(os.path.dirname(__file__))


# ============================================================
# MCP 설정 로딩 (mcp_config.json)
# ============================================================

def _load_mcp_config() -> dict:
    """mcp_config.json 을 로드하여 반환. 없으면 빈 dict."""
    config_path = os.path.join(os.path.dirname(__file__), "mcp_config.json")
    try:
        with open(config_path, encoding="utf-8") as f:
            raw = json.load(f)
        # _comment* 키 제거 후 반환
        return {k: v for k, v in raw.items() if not k.startswith("_")}
    except FileNotFoundError:
        return {}
    except Exception as e:
        logger.warning(f"mcp_config.json 로드 실패: {e}")
        return {}

MCP_CONFIG = _load_mcp_config()

# 파일 탐색 시 건너뛸 디렉터리 (캐시/로그/바이너리 등)
def _time_to_minutes(t: str) -> int:
    """'HH:MM' 문자열을 분(int)으로 변환. 예: '09:30' → 570"""
    h, m = map(int, t.split(":"))
    return h * 60 + m

# 파일 탐색 시 건너뛸 디렉터리 (캐시/로그/바이너리 등)
_FILE_SCAN_SKIP_DIRS = frozenset({
    "Cache", "Code Cache", "GPUCache", "DawnGraphiteCache",
    "DawnWebGPUCache", "CrashPad", "Crashpad", "logs",
    "INetCache", "INetHistory", "__pycache__",
})
_FILE_SCAN_SKIP_EXTS = frozenset({
    ".log", ".json", ".db", ".sqlite", ".tmp", ".js", ".css",
    ".html", ".htm", ".map", ".pak", ".bin", ".dll", ".exe",
    ".png", ".jpg", ".jpeg", ".ico", ".svg",
})

def _expand_env(value: str) -> str:
    """문자열 내 %VAR% 환경변수를 실제 값으로 치환."""
    return re.sub(r'%(\w+)%', lambda m: os.environ.get(m.group(1), m.group(0)), value)


# ============================================================
# Playwright 기반 그룹웨어 클라이언트
# ============================================================

class GroupwareClient:
    """
    Playwright를 사용한 그룹웨어 웹 자동화 클라이언트.

    ⚠️ 중요: 아래 셀렉터와 URL 패턴은 실제 그룹웨어(더존 아마란스,
    한글과컴퓨터, 다우오피스 등)에 맞게 반드시 수정해야 합니다.
    """

    def __init__(self):
        self.browser = None
        self.context = None
        self.page = None
        self.logged_in = False
        self._login_lock = asyncio.Lock()
        # ── 일정 모듈 상태 캐시 (불필요한 재설정 방지) ──
        self._sched_calendar = None   # 현재 선택된 사이드바 캘린더
        self._sched_view = None       # 'list' | 'grid'
        self._sched_month = None      # 현재 표시 월 'YYYY-MM'

    # ---------- 초기화 / 로그인 ----------

    async def initialize(self):
        """브라우저 초기화"""
        from playwright.async_api import async_playwright
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        self.context = await self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="ko-KR"
        )
        self.page = await self.context.new_page()
        logger.info("브라우저 초기화 완료")

    async def login(self) -> dict:
        """그룹웨어 로그인 (더존 BIZON 2단계 로그인)"""
        if not ERP_USERNAME or not ERP_PASSWORD:
            return {"success": False, "error": "자격증명이 설정되지 않았습니다. .env를 확인하세요."}
        async with self._login_lock:
            return await self._do_login()

    async def _do_login(self) -> dict:
        """실제 로그인 수행 (락 내부에서 호출)"""
        try:
            await self.page.goto(ERP_BASE_URL, wait_until="domcontentloaded", timeout=30000)
            await self.page.wait_for_timeout(3000)  # SPA 렌더링 대기

            # 이미 로그인된 상태: /login이 사라지면 성공으로 처리
            if "/login" not in self.page.url:
                self.logged_in = True
                logger.info("그룹웨어 이미 로그인된 상태")
                return {"success": True, "message": "로그인 성공 (세션 유지 중)"}

            await self.page.wait_for_selector("#reqLoginId", timeout=15000)

            # 1단계: 회사코드(이미 입력됨) + 아이디 입력 → "다음"
            await self.page.fill("#reqLoginId", ERP_USERNAME)
            await self.page.click("button.loginBtnFlex")

            # 2단계: 비밀번호 입력 → "로그인"
            await self.page.wait_for_selector("#reqLoginPw", state="visible", timeout=10000)
            await self.page.fill("#reqLoginPw", ERP_PASSWORD)
            await self.page.click("button.loginBtnFlex")

            # 로그인 완료 대기 (URL에서 /login이 사라지면 성공)
            await self.page.wait_for_function(
                "() => !window.location.hash.includes('/login')",
                timeout=20000
            )

            self.logged_in = True
            logger.info("그룹웨어 로그인 성공")
            return {"success": True, "message": "로그인 성공"}
        except Exception as e:
            logger.error(f"로그인 실패: {e}")
            return {"success": False, "error": f"로그인 실패: {str(e)}"}

    async def ensure_logged_in(self) -> bool:
        """세션 유효성 확인 → 필요 시 재로그인"""
        if not self.logged_in:
            return (await self.login())["success"]
        try:
            if "/login" in self.page.url:
                self.logged_in = False
                return (await self.login())["success"]
        except Exception:
            self.logged_in = False
            return (await self.login())["success"]
        return True

    # ============================================================
    #  메일 기능
    # ============================================================

    def _find_file_by_name(self, fname: str) -> Optional[str]:
        """파일명으로 Windows 내 알려진 위치를 검색해 실제 경로 반환.

        검색 순서:
          1. 로컬 PC 다운로드(Downloads) 폴더  ← 최우선
          2. 바탕화면(Desktop), 문서(Documents)
          3. C:/mcp/temp  (수동 배치 공유 폴더)
          4. Claude Desktop 패키지 전체 (pending-uploads / uploads / outputs)
        """
        userprofile = os.environ.get("USERPROFILE", "")

        # ── 1. Downloads 최우선 ──
        downloads = os.path.join(userprofile, "Downloads", fname)
        if os.path.isfile(downloads):
            logger.info(f"[find] Downloads에서 발견: {downloads}")
            return downloads

        # ── 2. Desktop / Documents ──
        for subdir in ("Desktop", "Documents"):
            candidate = os.path.join(userprofile, subdir, fname)
            if os.path.isfile(candidate):
                logger.info(f"[find] {subdir}에서 발견: {candidate}")
                return candidate

        # ── 3. MCP 공유 temp 폴더 ──
        tmp_dir = MCP_CONFIG.get("temp_dir", r"C:\mcp\temp")
        tmp_candidate = os.path.join(tmp_dir, fname)
        if os.path.isfile(tmp_candidate):
            logger.info(f"[find] temp_dir에서 발견: {tmp_candidate}")
            return tmp_candidate

        # ── 4. Claude Desktop 패키지 전체 검색 ──
        local_app = os.environ.get("LOCALAPPDATA", "")
        claude_pkg = os.path.join(local_app, "Packages", "Claude_pzs8sxrjxfjjc")
        if os.path.isdir(claude_pkg):
            candidates: list[tuple[float, str]] = []
            for dirpath, dirs, files in os.walk(claude_pkg):
                dirs[:] = [d for d in dirs if d not in _FILE_SCAN_SKIP_DIRS]
                if fname in files:
                    full = os.path.join(dirpath, fname)
                    candidates.append((os.path.getmtime(full), full))
            if candidates:
                candidates.sort(reverse=True)
                best = candidates[0][1]
                logger.info(f"[find] Claude Desktop 패키지에서 발견 (최신): {best}")
                return best

        return None

    def list_recent_output_files(self, hours: int = 24, limit: int = 10) -> dict:
        """Claude Desktop outputs 및 주요 경로에서 최근 생성/수정된 파일 목록 반환.

        Claude Desktop이 생성한 결과 파일을 gw_send_mail 첨부 전에 탐색할 때 사용.
        검색 위치: Claude Desktop outputs, pending-uploads, Downloads, C:/mcp/temp
        """
        cutoff = time.time() - (hours * 3600)
        candidates: list = []  # (mtime, path, source)

        def scan_dir(base: str, source: str):
            if not os.path.isdir(base):
                return
            for dirpath, dirs, files in os.walk(base):
                dirs[:] = [d for d in dirs if d not in _FILE_SCAN_SKIP_DIRS]
                for fname in files:
                    ext = os.path.splitext(fname)[1].lower()
                    if ext in _FILE_SCAN_SKIP_EXTS:
                        continue
                    full = os.path.join(dirpath, fname)
                    try:
                        mtime = os.path.getmtime(full)
                        if mtime >= cutoff:
                            candidates.append((mtime, full, source))
                    except Exception:
                        pass

        # 1. Claude Desktop local-agent-mode-sessions (outputs 포함)
        local_app = os.environ.get("LOCALAPPDATA", "")
        roaming_claude = os.path.join(
            local_app, "Packages", "Claude_pzs8sxrjxfjjc",
            "LocalCache", "Roaming", "Claude"
        )
        scan_dir(os.path.join(roaming_claude, "local-agent-mode-sessions"), "claude_outputs")
        scan_dir(os.path.join(roaming_claude, "pending-uploads"), "claude_pending")

        # 2. Downloads
        userprofile = os.environ.get("USERPROFILE", "")
        scan_dir(os.path.join(userprofile, "Downloads"), "downloads")

        # 3. C:\mcp\temp
        tmp_dir = MCP_CONFIG.get("temp_dir", r"C:\mcp\temp")
        scan_dir(tmp_dir, "mcp_temp")

        # 최신순 정렬, 파일명 중복 제거 후 limit 적용
        candidates.sort(key=lambda x: x[0], reverse=True)
        result = []
        seen_names: set = set()
        for mtime, path, source in candidates:
            fname = os.path.basename(path)
            if fname in seen_names:
                continue
            seen_names.add(fname)
            try:
                size_kb = round(os.path.getsize(path) / 1024, 1)
            except Exception:
                size_kb = 0
            result.append({
                "filename": fname,
                "path": path.replace("\\", "/"),
                "source": source,
                "modified": datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
                "size_kb": size_kb,
            })
            if len(result) >= limit:
                break

        return {
            "success": True,
            "count": len(result),
            "hours_range": hours,
            "files": result,
            "note": (
                "파일명 또는 path를 gw_send_mail의 attachments에 전달하면 "
                "자동으로 탐색하여 첨부합니다. 예: attachments=['파일명.xlsx']"
            ),
        }

    async def _resolve_path(self, file_path: str) -> str:
        """첨부파일 경로를 Windows MCP 서버에서 접근 가능한 경로로 변환.

        Claude Desktop은 자체 가상 파일시스템을 사용하므로 컨테이너 내부의
        /mnt/c/ 가 실제 Windows C: 드라이브와 공유되지 않습니다.
        따라서 경로 접두사 매핑 대신 파일명 기반 검색을 우선 사용합니다.

        변환 순서:
          1. 이미 Windows 경로(드라이브 문자 포함) → 즉시 반환
          2. Windows에서 직접 접근 가능(예: WSL /mnt/c/... 실 마운트) → 즉시 반환
          3. 파일명 기반 검색: C:/mcp/temp → Claude Desktop uploads → 사용자 폴더
          4. WSL cp 시도 (WSL 설치된 환경 한정)
        """
        file_path = file_path.strip()

        # ── 1. 이미 Windows 경로 ──
        if not file_path.startswith("/"):
            # 전체 경로로 접근 가능하면 즉시 반환
            if os.path.isfile(file_path):
                return file_path
            # 파일명만 지정한 경우(경로 구분자 없음) → Downloads부터 검색
            if not any(c in file_path for c in ("/", "\\")):
                found = self._find_file_by_name(file_path)
                if found:
                    return found
            return file_path

        # ── 2. 직접 접근 가능 (WSL 실 마운트 등) ──
        if os.path.isfile(file_path):
            return file_path

        fname = os.path.basename(file_path)

        # ── 3. 파일명 기반 검색 (컨테이너 경로 형식 무관하게 동작) ──
        found = self._find_file_by_name(fname)
        if found:
            return found

        # ── 4. WSL cp → C:/mcp/temp (WSL 환경 한정 최후 수단) ──
        tmp_dir = MCP_CONFIG.get("temp_dir", r"C:\mcp\temp")
        os.makedirs(tmp_dir, exist_ok=True)
        tmp_win = os.path.join(tmp_dir, fname)
        # C:\mcp\temp → /mnt/c/mcp/temp
        drive = tmp_dir[0].lower()
        tmp_linux = f"/mnt/{drive}/" + tmp_dir[3:].replace("\\", "/")
        try:
            result = subprocess.run(
                ["wsl", "cp", file_path, f"{tmp_linux}/{fname}"],
                capture_output=True, text=True, timeout=15
            )
            if result.returncode == 0 and os.path.isfile(tmp_win):
                logger.info(f"[wsl cp] 복사 완료: {file_path} → {tmp_win}")
                return tmp_win
            logger.warning(f"wsl cp 실패: {result.stderr.strip()}")
        except Exception as e:
            logger.warning(f"wsl cp 예외: {e}")

        logger.warning(f"경로 변환 실패: {file_path} (파일명: {fname})")
        return file_path

    async def _close_all_popups(self):
        """아마란스 접속 후 남아있는 모든 팝업/패널/모달/오버레이를 완전 제거 (일정 등록 전 사전 정리)"""
        try:
            # 1단계: 업데이트 배너 · 로딩 오버레이 DOM 직접 숨김
            await self.page.evaluate("""() => {
                const banner = document.querySelector('.systemAlertUpdate');
                if (banner) banner.style.display = 'none';
                const loading = document.getElementById('a10-first-loading');
                if (loading) loading.style.display = 'none';
                // 전체화면 고정 오버레이 제거
                for (const el of document.querySelectorAll(
                    '[class*="loading"], [class*="Loading"], [class*="overlay"], [class*="Overlay"], ' +
                    '[class*="spinner"], [class*="Spinner"], [class*="mask"], [class*="Mask"]'
                )) {
                    const s = window.getComputedStyle(el);
                    if (s.position === 'fixed' || s.position === 'absolute') {
                        const r = el.getBoundingClientRect();
                        if (r.width > 200 && r.height > 200) el.style.display = 'none';
                    }
                }
            }""")
            await self.page.wait_for_timeout(300)

            # 2단계: OBTDialog / 모달 확인·닫기 버튼 클릭
            await self.page.evaluate("""() => {
                const btns = document.querySelectorAll(
                    '[class*="OBTDialog_dialogRoot"] button, .obtdialog button, ' +
                    '[data-orbit-component="OBTDialog"] button, ' +
                    '[class*="modal"] button, [class*="Modal"] button, ' +
                    '[class*="popup"] button, [class*="Popup"] button'
                );
                for (const btn of btns) {
                    const txt = (btn.innerText || '').trim();
                    if (['확인', '닫기', '취소', '확인하기', 'OK', 'Close'].includes(txt)) {
                        btn.click(); break;
                    }
                }
                const dim = document.querySelector('._dimClicker');
                if (dim) dim.click();
            }""")
            await self.page.wait_for_timeout(300)

            # 3단계: 활성 슬라이드 패널(.pubScLayer, .pubLayerSlide) 닫기 버튼 클릭 반복
            for _ in range(4):
                active_count = await self.page.evaluate("""() =>
                    document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active').length
                """)
                if active_count == 0:
                    break
                closed = await self.page.evaluate("""() => {
                    for (const panel of document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active')) {
                        const closeBtn = panel.querySelector(
                            'button[class*="close"], button[aria-label*="닫기"], button[title*="닫기"], ' +
                            '.scLayerClose, .btn_close, .ic_close, button.close'
                        );
                        if (closeBtn) { closeBtn.click(); return true; }
                    }
                    return false;
                }""")
                if not closed:
                    # 닫기 버튼 없으면 ESC (슬라이드 패널은 React 상태 오염 없음)
                    await self.page.keyboard.press("Escape")
                await self.page.wait_for_timeout(200)

            # 4단계: 열린 드롭다운 외부 클릭으로 닫기
            await self.page.evaluate("""() => {
                const hasOpenDropdown = document.querySelector(
                    '[class*="OBTDropDownList_show__"], [class*="scDropDown_list"]'
                );
                if (hasOpenDropdown) {
                    const fc = document.querySelector('.fc-view, .fc-daygrid, body');
                    if (fc) fc.click();
                }
            }""")
            await self.page.wait_for_timeout(200)
            logger.info("_close_all_popups: 팝업 정리 완료")
        except Exception as e:
            logger.warning(f"_close_all_popups 오류(무시): {e}")

    async def _close_dialogs(self):
        """열린 OBTDialog/팝업/드롭다운/로딩오버레이를 모두 닫음 (클릭 방해 방지)"""
        try:
            await self.page.evaluate("""() => {
                // 업데이트 배너 / 로딩 오버레이 숨기기
                const banner = document.querySelector(".systemAlertUpdate");
                if (banner) banner.style.display = "none";
                const loading = document.getElementById("a10-first-loading");
                if (loading) loading.style.display = "none";
                // pointer-events를 막는 전체 오버레이 숨기기 (클래스명 기반)
                for (const el of document.querySelectorAll(
                    '[class*="loading"], [class*="Loading"], [class*="overlay"], [class*="Overlay"], ' +
                    '[class*="spinner"], [class*="Spinner"], [class*="mask"], [class*="Mask"]'
                )) {
                    const s = window.getComputedStyle(el);
                    if (s.position === 'fixed' || s.position === 'absolute') {
                        const r = el.getBoundingClientRect();
                        if (r.width > 200 && r.height > 200) el.style.display = "none";
                    }
                }
                // OBTDialog 닫기
                const btns = document.querySelectorAll(
                    '[class*="OBTDialog_dialogRoot"] button, .obtdialog button, ' +
                    '[data-orbit-component="OBTDialog"] button'
                );
                for (const btn of btns) {
                    const txt = (btn.innerText || '').trim();
                    if (['확인','닫기','취소','OK','Close'].includes(txt)) { btn.click(); break; }
                }
                // 딤 클릭커 닫기
                const dim = document.querySelector('._dimClicker');
                if (dim) dim.click();
                // 열린 드롭다운 닫기: 폼 패널 내 제목 입력란 클릭 (fc-view 클릭 시 폼 패널이 닫히는 버그 방지)
                const hasOpenDropdown = document.querySelector(
                    '[class*="OBTDropDownList_show__"], [class*="scDropDown_list"]'
                );
                if (hasOpenDropdown) {
                    const titleInp = document.querySelector('#scTitleInput input');
                    if (titleInp) { titleInp.click(); }
                    else {
                        const hdr = document.querySelector('.pubScLayerIn .scTitleWrap, .pubScLayerIn');
                        if (hdr) hdr.click();
                    }
                }
            }""")
            await self.page.wait_for_timeout(500)
        except Exception:
            pass

    async def _navigate_to_mail(self):
        """메일함 페이지로 이동 (사이드바 '메일' 클릭)"""
        await self._close_dialogs()
        self._sched_calendar = None
        self._sched_view = None
        self._sched_month = None

        current_url = self.page.url
        already_in_mail = "moduleCode=UD" in current_url and "menuCode=UDA" in current_url

        await self.page.get_by_text('메일', exact=True).first.click()

        if already_in_mail:
            # 이미 메일 모듈 — 짧은 대기 후 통과
            try:
                await self.page.wait_for_selector(".listItem", timeout=5000)
            except Exception:
                pass
        else:
            # 메일 모듈 URL 전환 대기 (listItem 유무와 무관)
            try:
                await self.page.wait_for_function(
                    "() => window.location.hash.includes('moduleCode=UD')",
                    timeout=15000
                )
            except Exception:
                await self.page.wait_for_timeout(2000)
            # 메일 컨테이너 또는 메일 목록 중 먼저 나타나는 것 대기
            try:
                await self.page.wait_for_selector(
                    ".listItem, .UDA0020, .mailWrap, .mailList, .inboxList",
                    timeout=5000
                )
            except Exception:
                await self.page.wait_for_timeout(1000)

    async def _navigate_to_schedule(self):
        """일정 페이지로 이동 (FullCalendar 완전 로딩 대기)"""
        current_url = self.page.url
        needs_nav = "moduleCode=UE" not in current_url or "menuCode=UEA" not in current_url
        if needs_nav:
            # 다른 모듈로 이동하면 캐시 초기화
            self._sched_calendar = None
            self._sched_view = None
            self._sched_month = None
            await self.page.goto(SCHEDULE_PAGE_URL, wait_until="domcontentloaded")
        else:
            # 이미 일정 페이지: .fc-view가 DOM에 있으면 즉시 반환
            already_loaded = await self.page.evaluate(
                "() => !!document.querySelector('.fc-view, .fc-daygrid')"
            )
            if already_loaded:
                return

        # FullCalendar 완전 로딩 대기 (최대 10초)
        try:
            await self.page.wait_for_selector(".fc-view, .fc-daygrid", timeout=10000)
        except Exception:
            # FullCalendar 미로드 시 페이지 새로고침 후 재시도
            logger.warning("캘린더 로딩 실패, 페이지 새로고침...")
            await self.page.reload(wait_until="domcontentloaded")
            await self.page.wait_for_selector(".fc-view, .fc-daygrid, .btn_sideRegi", timeout=8000)

    async def _select_time_from_dropdown(self, comp_id: str, target_time: str) -> bool:
        """OBTComplete2 시간 드롭다운에서 특정 시간 선택.

        방식: page.mouse.click(좌표) → Ctrl+A → type → ArrowDown → Enter
        - Vue.js 컴포넌트는 dispatchEvent 합성 이벤트 미인식 → 좌표 기반 실제 마우스 이벤트 필수
        - visible input만 사용: fallback(hidden form) 타이핑은 부작용 유발
        """
        # visible input 좌표 탐색
        coords = await self.page.evaluate(f"""() => {{
            for (const inp of document.querySelectorAll('#{comp_id} input')) {{
                const r = inp.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) {{
                    return {{
                        ok: true,
                        x: Math.round(r.left + r.width/2),
                        y: Math.round(r.top + r.height/2),
                        w: Math.round(r.width), h: Math.round(r.height)
                    }};
                }}
            }}
            return {{ok: false}};
        }}""")
        if not coords or not coords.get('ok'):
            logger.warning(f"[{comp_id}] 가시 input 없음 — 섹션 미전개 상태")
            return False

        # 실제 마우스 클릭으로 Vue click 핸들러 발화
        await self.page.mouse.click(coords['x'], coords['y'])
        await self.page.wait_for_timeout(300)

        # Ctrl+A → type → ArrowDown(드롭다운 열고 첫 항목 선택) → Enter
        await self.page.keyboard.press("Control+a")
        await self.page.keyboard.type(target_time, delay=50)
        await self.page.wait_for_timeout(300)   # 드롭다운 필터링 대기
        await self.page.keyboard.press("ArrowDown")  # 첫 번째 매칭 항목으로 이동
        await self.page.wait_for_timeout(100)
        await self.page.keyboard.press("Enter")      # 선택 확정
        await self.page.wait_for_timeout(500)

        # 결과 확인 (visible input 한정)
        val = await self.page.evaluate(f"""() => {{
            for (const inp of document.querySelectorAll('#{comp_id} input')) {{
                if (inp.getBoundingClientRect().width > 0) return inp.value;
            }}
            return '';
        }}""")
        logger.info(f"시간 선택 [{comp_id}]: 목표={target_time}, 실제={val}")
        return val == target_time

    async def _diagnose_form(self) -> dict:
        """일정 등록 폼의 실제 DOM 구조를 수집 (캘린더·메모 셀렉터 디버깅용)."""
        return await self.page.evaluate("""() => {
            const res = {textareas: [], contentEditables: [], calRow: [], buttonsInPanel: []};

            // 모든 textarea
            for (const ta of document.querySelectorAll('textarea')) {
                const r = ta.getBoundingClientRect();
                res.textareas.push({ph: ta.placeholder, id: ta.id,
                    cls: ta.className.substring(0, 80),
                    vis: r.width > 0, x: Math.round(r.x), y: Math.round(r.y),
                    w: Math.round(r.width), h: Math.round(r.height)});
            }

            // contenteditable 요소
            for (const el of document.querySelectorAll('[contenteditable]')) {
                const r = el.getBoundingClientRect();
                if (r.width > 0) res.contentEditables.push({
                    tag: el.tagName, cls: el.className.substring(0, 80),
                    x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)});
            }

            // '캘린더' 텍스트 포함 요소 및 주변 컨트롤
            for (const el of document.querySelectorAll('*')) {
                const t = (el.innerText || '').trim();
                if (!t.startsWith('캘린더')) continue;
                const r = el.getBoundingClientRect();
                if (r.width <= 0 || r.width > 400) continue;
                res.calRow.push({tag: el.tagName, cls: el.className.substring(0,60),
                    text: t.substring(0,40), x: Math.round(r.x), y: Math.round(r.y)});
                const par = el.closest('tr,div,li,section') || el.parentElement;
                if (par) {
                    for (const c of par.querySelectorAll('button,[role="button"],select,input')) {
                        const cr = c.getBoundingClientRect();
                        if (cr.width > 0) res.calRow.push({tag:'→'+c.tagName, cls:c.className.substring(0,60),
                            text:(c.innerText||c.value||'').trim().substring(0,30), x:Math.round(cr.x), y:Math.round(cr.y)});
                    }
                }
            }

            // x > 1200인 우측 패널의 버튼/select
            for (const el of document.querySelectorAll('button,select,[role="combobox"],[role="button"]')) {
                const r = el.getBoundingClientRect();
                if (r.x > 1200 && r.width > 0) res.buttonsInPanel.push({
                    tag: el.tagName, cls: el.className.substring(0, 60),
                    text: (el.innerText||el.value||'').trim().substring(0,30),
                    x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)});
            }
            return res;
        }""")

    async def _select_calendar(self, calendar_name: str) -> bool:
        """일정 등록 폼에서 캘린더를 이름으로 선택.

        실제 DOM 구조 (더존 아마란스):
          - '캘린더' 라벨: div.txt (td/span이 아닌 div 사용)
          - 드롭다운 영역: div.scDropDown 또는 div.OBTDropDownList_default
            (OBTComplete2의 dropDownButton 클래스 없음)

        처리 순서:
          1. '캘린더' 라벨(div/td/span 등) → 같은 행 scDropDown/OBTDropDownList 탐색
          2. 없으면 elementFromPoint로 라벨 우측 영역 스캔
          3. 드롭다운 클릭 → 리스트 팝업 대기
          4. LI / div.txt 항목에서 calendar_name 탐색 → mouse.click()
        """
        try:
            # ── 1. '캘린더' 라벨 위치 파악 + 드롭다운 클릭 좌표 획득 ──
            btn_rect = await self.page.evaluate("""() => {
                // div 포함 모든 텍스트 라벨 탐색 (더존 아마란스는 div.txt 사용)
                let labelY = null;
                for (const el of document.querySelectorAll('td, th, label, span, p, div')) {
                    // 자식 요소가 있으면 순수 텍스트 노드가 아님 → 정확한 라벨만 선별
                    if (el.children.length > 0) continue;
                    const t = (el.innerText || '').trim();
                    if (t !== '캘린더') continue;
                    const r = el.getBoundingClientRect();
                    // 우측 패널(x > 1000), 너비 150px 이하
                    if (r.width <= 0 || r.width > 150 || r.x < 1000) continue;
                    labelY = r.top + r.height / 2;
                    break;
                }
                if (!labelY) return null;

                // 방법 A: scDropDown 또는 OBTDropDownList (더존 아마란스 실제 구조)
                for (const sel of [
                    '[class*="scDropDown"]',
                    '[class*="OBTDropDownList"]',
                    '[class*="dropDownButton"]',
                    '[class*="DropDownButton"]',
                ]) {
                    for (const el of document.querySelectorAll(sel)) {
                        const r = el.getBoundingClientRect();
                        if (r.width <= 0 || r.x < 1000) continue;
                        if (Math.abs((r.top + r.height / 2) - labelY) <= 30) {
                            return {l: Math.round(r.left), t: Math.round(r.top),
                                    w: Math.round(r.width), h: Math.round(r.height), method: sel};
                        }
                    }
                }

                // 방법 B: elementFromPoint로 라벨 우측 영역 스캔
                for (let x = 1880; x >= 1050; x -= 6) {
                    const el = document.elementFromPoint(x, labelY);
                    if (!el) continue;
                    const r = el.getBoundingClientRect();
                    if (r.width > 5 && r.width < 280 && r.height > 5 && r.x >= 1000) {
                        return {l: Math.round(r.left), t: Math.round(r.top),
                                w: Math.round(r.width), h: Math.round(r.height),
                                method: 'elementFromPoint', tag: el.tagName,
                                cls: el.className.substring(0, 80)};
                    }
                }
                return null;
            }""")

            if not btn_rect:
                logger.warning("캘린더 드롭다운 버튼 미발견")
                return False

            # ── 1b. 현재 선택된 캘린더 값 로깅 (조기반환 없음 — 항상 드롭다운에서 명시적 선택) ──
            current_val = await self.page.evaluate("""([l, t, w, h]) => {
                var cx = l + w / 2, cy = t + h / 2;
                var el = document.elementFromPoint(cx, cy);
                return el ? (el.innerText || '').trim() : '';
            }""", [btn_rect["l"], btn_rect["t"], btn_rect["w"], btn_rect["h"]])
            logger.info(f"캘린더 드롭다운 현재값: '{current_val}' (목표: '{calendar_name}')")

            # ── 2. 드롭다운 클릭 ──
            click_x = btn_rect["l"] + btn_rect["w"] // 2
            click_y = btn_rect["t"] + btn_rect["h"] // 2
            logger.info(f"캘린더 드롭다운 클릭: ({click_x}, {click_y})  method={btn_rect.get('method')}")
            await self.page.mouse.click(click_x, click_y)
            await self.page.wait_for_timeout(800)

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_cal_dropdown.png"))

            # ── 3a. 드롭다운 전체 항목 수집 후 로깅 (x 제한 1000으로 완화) ──
            dropdown_top = btn_rect["t"]
            item_coord = await self.page.evaluate("""([calName, minY]) => {
                var allItems = [];
                // 방법 A: scDropDown 형제 항목
                for (var el of document.querySelectorAll('[class*="scDropDown"]')) {
                    var t = (el.innerText || '').trim();
                    var r = el.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0 && r.height < 60 && r.x > 1000 && r.y > minY)
                        allItems.push({x: Math.round(r.left+r.width/2), y: Math.round(r.top+r.height/2),
                                       text: t, method: 'scDropDown'});
                }
                // 방법 B: LI 또는 a 태그
                for (var el2 of document.querySelectorAll('li, a')) {
                    var t2 = (el2.innerText || '').trim();
                    var r2 = el2.getBoundingClientRect();
                    if (r2.width > 0 && r2.height > 0 && r2.height < 60 && r2.x > 1000 && r2.y > minY)
                        allItems.push({x: Math.round(r2.left+r2.width/2), y: Math.round(r2.top+r2.height/2),
                                       text: t2, method: 'li/a'});
                }
                // 정확히 일치하는 항목 우선, 없으면 포함하는 항목
                var exact = allItems.filter(i => i.text === calName);
                var contains = allItems.filter(i => i.text !== calName && i.text.includes(calName));
                var all = exact.concat(contains);
                if (all.length > 0) return {match: all[0], allTexts: allItems.map(i=>i.text).slice(0,10)};
                return {match: null, allTexts: allItems.map(i=>i.text).slice(0,10)};
            }""", [calendar_name, dropdown_top + 10])

            logger.info(f"드롭다운 항목 목록: {item_coord.get('allTexts', [])}")
            found_item = item_coord.get("match")

            if found_item:
                await self.page.mouse.click(found_item["x"], found_item["y"])
                await self.page.wait_for_timeout(400)
                # 선택 후 실제 값 확인
                new_val = await self.page.evaluate("""([l, t, w, h]) => {
                    var cx = l + w / 2, cy = t + h / 2;
                    var el = document.elementFromPoint(cx, cy);
                    return el ? (el.innerText || '').trim() : '';
                }""", [btn_rect["l"], btn_rect["t"], btn_rect["w"], btn_rect["h"]])
                logger.info(f"캘린더 선택 후 값: '{new_val}' (목표: '{calendar_name}', 클릭: {found_item})")
                return True

            # ── 3b. elementFromPoint Y축 스캔 (fallback) ──
            scan_x = click_x
            target_y = await self.page.evaluate("""([x, calName, minY]) => {
                var matches = [];
                for (var y = minY; y <= 900; y += 2) {
                    var el = document.elementFromPoint(x, y);
                    if (!el) continue;
                    var text = (el.innerText || '').trim();
                    if (text === calName || text.includes(calName)) matches.push(y);
                }
                // 첫 번째 정확 일치 우선
                return matches.length > 0 ? matches[0] : null;
            }""", [scan_x, calendar_name, dropdown_top + 10])

            if target_y is not None:
                await self.page.mouse.click(scan_x, target_y)
                await self.page.wait_for_timeout(400)
                logger.info(f"캘린더 선택 완료 (LI scan): {calendar_name}")
                return True

            logger.warning(f"캘린더 '{calendar_name}' 항목 미발견 (드롭다운 열렸으나 항목 없음)")
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_cal_select_fail.png"))
            # 드롭다운이 열린 채로 반환되면 이후 폼 입력을 방해하므로 닫기
            # ⚠️ Escape 금지 — 폼 패널 전체가 닫힘. 제목 입력란 클릭으로 드롭다운만 닫기
            await self.page.evaluate("""() => {
                const inp = document.querySelector('#scTitleInput input');
                if (inp) { inp.click(); return; }
                const hdr = document.querySelector('.pubScLayerIn .scTitleWrap, .pubScLayerIn');
                if (hdr) hdr.click();
            }""")
            await self.page.wait_for_timeout(300)
            return False

        except Exception as e:
            logger.warning(f"캘린더 선택 오류: {e}")
            return False

    async def _hide_update_banner(self):
        """업데이트 진행중 배너 숨김 (클릭 방해 방지)"""
        await self.page.evaluate("""() => {
            const banner = document.querySelector(".systemAlertUpdate");
            if (banner) banner.style.display = "none";
            const loading = document.getElementById("a10-first-loading");
            if (loading) loading.style.display = "none";
        }""")

    async def _parse_mail_rows(self, rows) -> list:
        """메일 목록 행(.listItem) 파싱 공통 헬퍼 (list_inbox / search_mail 공용)"""
        mails = []
        for idx, row in enumerate(rows):
            row_class = await row.get_attribute("class") or ""
            is_read = "unRead" not in row_class
            sender_el = await row.query_selector(".item-sender .addr")
            subject_el = await row.query_selector(".item-subject .title")
            date_el = await row.query_selector(".item-date")
            size_el = await row.query_selector(".item-size")
            file_el = await row.query_selector(".item-file")
            mails.append({
                "index": idx,
                "mail_id": str(idx),
                "from": (await sender_el.text_content()).strip() if sender_el else "",
                "subject": (await subject_el.text_content()).strip() if subject_el else "",
                "date": (await date_el.text_content()).strip() if date_el else "",
                "size": (await size_el.text_content()).strip() if size_el else "",
                "has_attachment": bool(await file_el.query_selector("span, img, svg")) if file_el else False,
                "is_read": is_read,
            })
        return mails

    async def list_inbox(self, page_num: int = 1, per_page: int = 20) -> dict:
        """수신함 메일 목록 조회"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            await self._navigate_to_mail()

            rows = await self.page.query_selector_all(".listItem")
            mails = await self._parse_mail_rows(rows)
            return {"success": True, "page": page_num, "count": len(mails), "mails": mails}
        except Exception as e:
            return {"success": False, "error": str(e)}

    async def read_mail(self, mail_id: str) -> dict:
        """메일 상세 읽기 (목록에서 클릭하여 내용 읽기)"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            await self._navigate_to_mail()
            rows = await self.page.query_selector_all(".listItem")

            try:
                idx = int(mail_id)
            except ValueError:
                return {"success": False, "error": f"mail_id는 정수(index) 형식이어야 합니다: {mail_id}"}

            if idx >= len(rows):
                return {"success": False, "error": f"메일 인덱스 {idx} 없음 (목록 총 {len(rows)}개)"}

            # 로딩 오버레이 제거 후 JS 클릭 (listItem은 DOM에 있지만 Playwright visibility 기준 미충족)
            await self.page.evaluate("""() => {
                document.querySelectorAll('.OBTLoading_wrapper__1NIDH, [class*="Loading"]')
                    .forEach(el => { el.style.display = 'none'; el.style.pointerEvents = 'none'; });
            }""")
            await self.page.evaluate(
                "(el) => { el.scrollIntoView({block:'center'}); el.click(); }", rows[idx]
            )

            # 메일 상세 패널 로딩 대기
            await self.page.wait_for_timeout(2000)

            # ── 본문 iframe 탐색 (읽기 뷰 iframe: 더존 아마란스는 읽기용 iframe 별도 사용) ──
            body = ""
            body_frame = None
            for frame in self.page.frames:
                name = frame.name or ""
                url = frame.url or ""
                if any(k in name for k in ("dzeditor", "mail", "view", "body", "read")):
                    body_frame = frame
                    break
                if any(k in url for k in ("mailBody", "viewMail", "readMail")):
                    body_frame = frame
                    break

            if body_frame:
                try:
                    body_el = await body_frame.query_selector("body")
                    body = (await body_el.text_content()).strip() if body_el else ""
                except Exception:
                    body = ""

            # ── 헤더(발신/수신/제목/날짜) + 첨부파일: JS innerText 파싱 ──
            detail_js = await self.page.evaluate("""() => {
                // 메일 상세 패널: 우측(x>700) 또는 페이지 전체에서 메일 뷰 찾기
                const candidates = [
                    ...document.querySelectorAll(
                        '.viewHead, .mailView, [class*="viewHead"], [class*="mailView"], ' +
                        '[class*="mailRead"], [class*="readMail"], [class*="mailDetail"], ' +
                        '.UDA0050, .UDA0060, [class*="UDA00"]'
                    )
                ].filter(el => {
                    const r = el.getBoundingClientRect();
                    return r.width > 100 && r.height > 50;
                });

                // 가장 크고 오른쪽에 있는 패널 선택
                candidates.sort((a, b) => {
                    const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
                    return (rb.width * rb.height) - (ra.width * ra.height);
                });

                const panel = candidates[0] || document.body;
                const panelText = (panel.innerText || '').substring(0, 2000);

                // 간단 라벨 파싱 (보낸사람/받는사람/제목/날짜)
                let sender = '', to_addr = '', subject = '', date = '';

                // 방법 A: 특정 셀렉터 시도
                const selMap = {
                    sender: ['.item-from .addr', '.from .addr', '.from', '[class*="sender"]', '[class*="from"]'],
                    to:     ['.item-to .addr', '.to .addr', '.to', '[class*="receiver"]'],
                    subject:['.item-subject .title', '.subject', '[class*="subject"]', '[class*="title"]'],
                    date:   ['.item-date', '.date', '[class*="date"]', '[class*="time"]'],
                };
                for (const [field, sels] of Object.entries(selMap)) {
                    for (const sel of sels) {
                        const el = panel.querySelector(sel);
                        if (el) {
                            const t = (el.innerText || el.textContent || '').trim();
                            if (t) {
                                if (field === 'sender') sender = t;
                                else if (field === 'to') to_addr = t;
                                else if (field === 'subject') subject = t;
                                else if (field === 'date') date = t;
                                break;
                            }
                        }
                    }
                }

                // 첨부파일
                const attachments = [];
                for (const el of panel.querySelectorAll(
                    '.attachList .attachItem, .fileList .fileItem, [class*="attach"] a, [class*="file"] a'
                )) {
                    const t = (el.innerText || el.textContent || '').trim();
                    if (t && !attachments.find(a => a.filename === t)) {
                        attachments.push({filename: t});
                    }
                }

                // frame 이름 목록 (디버그용)
                const frames = [...document.querySelectorAll('iframe')].map(f =>
                    (f.name || f.id || f.src || '').substring(0, 60)
                );

                return {sender, to_addr, subject, date, attachments, panelText, frames};
            }""")

            # iframe body 미확보 시 패널 내 텍스트에서 body 추출
            if not body and detail_js.get("panelText"):
                # 헤더 다음 텍스트를 body로 간주 (간이 파싱)
                panel_lines = detail_js["panelText"].split("\n")
                body_lines = [l.strip() for l in panel_lines if l.strip()
                              and l.strip() not in (detail_js.get("subject",""), detail_js.get("date",""),
                                                     detail_js.get("sender",""), detail_js.get("to_addr",""))]
                body = "\n".join(body_lines[3:]) if len(body_lines) > 3 else ""

            # iframe에서 본문을 읽지 못했으면 실제 읽기 iframe 재시도
            if not body:
                for frame in self.page.frames:
                    if frame.name and frame != self.page.main_frame:
                        try:
                            body_el = await frame.query_selector("body")
                            candidate = (await body_el.text_content()).strip() if body_el else ""
                            if len(candidate) > 20:
                                body = candidate
                                break
                        except Exception:
                            continue

            detail = {
                "mail_id": mail_id,
                "from": detail_js.get("sender", ""),
                "to": detail_js.get("to_addr", ""),
                "subject": detail_js.get("subject", ""),
                "date": detail_js.get("date", ""),
                "body": body,
                "attachments": detail_js.get("attachments", []),
                "_debug_frames": detail_js.get("frames", []),
                "_debug_panel_preview": detail_js.get("panelText", "")[:300],
            }
            return {"success": True, "mail": detail}
        except Exception as e:
            return {"success": False, "error": str(e)}

    async def send_mail(self, mail_data: dict) -> dict:
        """메일 발송 (더존 아마란스 메일쓰기 UI)"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        staged_temps: list = []  # 발송 후 삭제할 temp 복사본 (try 밖에 선언하여 예외 핸들러에서도 접근 가능)
        try:
            # ── 메일 모듈로 이동 + 다이얼로그 닫기 (_navigate_to_mail이 _close_dialogs 포함) ──
            await self._navigate_to_mail()

            # ── 로딩 오버레이 제거 ──
            await self.page.evaluate("""() => {
                document.querySelectorAll('.OBTLoading_wrapper__1NIDH, [class*="Loading"]')
                    .forEach(el => { el.style.display = 'none'; el.style.pointerEvents = 'none'; });
            }""")
            await self.page.wait_for_timeout(500)

            # ── 메일쓰기 버튼 클릭 (JS - 오버레이 우회, 오타 주의: btn_sideWirte) ──
            await self.page.evaluate(
                "() => { const b = document.querySelector('button.btn_sideWirte'); if (b) b.click(); }"
            )
            try:
                await self.page.wait_for_selector('input[placeholder="받는사람을 입력해주세요."]', timeout=5000)
            except Exception:
                await self.page.wait_for_timeout(1500)

            # ── 받는사람 입력 (OBTAutoCompleteChips) ──
            to_input = self.page.locator('input[placeholder="받는사람을 입력해주세요."]')
            await to_input.wait_for(state="visible", timeout=10000)
            for addr in mail_data["to"].replace(";", ",").split(","):
                addr = addr.strip()
                if addr:
                    await to_input.fill(addr)
                    await self.page.wait_for_timeout(500)
                    await self.page.keyboard.press("Tab")
                    await self.page.wait_for_timeout(500)

            # ── 참조(CC) 입력 ──
            if mail_data.get("cc"):
                cc_input = self.page.locator('input[placeholder="참조를 입력해주세요."]')
                for addr in mail_data["cc"].replace(";", ",").split(","):
                    addr = addr.strip()
                    if addr:
                        await cc_input.fill(addr)
                        await self.page.wait_for_timeout(500)
                        await self.page.keyboard.press("Tab")
                        await self.page.wait_for_timeout(500)

            # ── 제목 입력 ──
            subject_input = self.page.locator('tr:has-text("제목") input')
            await subject_input.click()
            await subject_input.fill(mail_data["subject"])
            await self.page.wait_for_timeout(300)

            # ── 본문 입력 (dzeditor_0 iframe) ──
            editor_frame = self.page.frame(name="dzeditor_0")
            if editor_frame:
                await editor_frame.evaluate("""(bodyText) => {
                    const doc = document.body;
                    const sig = doc.querySelector('.dze_signature');
                    // 서명 앞의 빈 p 정리
                    if (sig) {
                        let prev = sig.previousElementSibling;
                        while (prev && prev.tagName === 'P' && !prev.textContent.trim()) {
                            const rm = prev;
                            prev = prev.previousElementSibling;
                            rm.remove();
                        }
                    }
                    // 본문 p 태그 삽입
                    const frag = document.createDocumentFragment();
                    for (const line of bodyText.split('\\n')) {
                        const p = document.createElement('p');
                        p.textContent = line;
                        frag.appendChild(p);
                    }
                    if (sig) {
                        doc.insertBefore(frag, sig);
                    } else {
                        doc.appendChild(frag);
                    }
                }""", mail_data["body"])
            else:
                logger.warning("dzeditor_0 iframe을 찾지 못했습니다. 본문 입력 생략.")

            await self.page.wait_for_timeout(500)

            # ── 첨부파일 업로드 ──
            attachments = mail_data.get("attachments") or []
            attached_files = []
            failed_files = []

            tmp_dir = MCP_CONFIG.get("temp_dir", r"C:\mcp\temp")
            os.makedirs(tmp_dir, exist_ok=True)

            if attachments:
                for file_path in attachments:
                    original_path = file_path.strip()

                    # ── STEP 1: 파일 탐색 (Linux 컨테이너 경로 → Windows 실제 경로) ──
                    resolved = await self._resolve_path(original_path)
                    if not os.path.isfile(resolved):
                        searched_names = [os.path.basename(fp.strip()) for fp in attachments]
                        userprofile = os.environ.get("USERPROFILE", "")
                        downloads_dir = os.path.join(userprofile, "Downloads")
                        return {
                            "success": False,
                            "needs_file_location": True,
                            "searched_files": searched_names,
                            "message": (
                                f"첨부파일을 찾을 수 없어 메일 발송을 중단했습니다.\n\n"
                                f"검색한 파일명: {', '.join(searched_names)}\n"
                                f"검색한 위치: Downloads, Desktop, Documents, {tmp_dir}\n\n"
                                f"파일이 어디에 있는지 알려주세요.\n"
                                f"예시:\n"
                                f"  - 다운로드 폴더에 있음: '{downloads_dir}'에 파일을 저장 후 다시 요청\n"
                                f"  - 다른 위치: 전체 경로(예: C:/Users/.../파일명.docx)를 알려주세요"
                            ),
                        }

                    # ── STEP 2: C:\mcp\temp 로 임시 복사 ──
                    fname = os.path.basename(resolved)
                    staged = os.path.join(tmp_dir, fname)
                    if os.path.abspath(resolved) != os.path.abspath(staged):
                        shutil.copy2(resolved, staged)
                        staged_temps.append(staged)
                        logger.info(f"임시복사: {resolved} → {staged}")
                    else:
                        staged_temps.append(staged)   # temp에 이미 있어도 발송 후 삭제
                    file_path = staged

                    # ── STEP 3: ERP 메일 UI에 첨부 ──
                    try:
                        uploaded = False

                        # 방법 1: "내 PC" 버튼 Playwright 클릭 → file chooser
                        # JS el.click()은 브라우저 보안상 파일 다이얼로그를 열 수 없으므로
                        # Playwright locator.click()을 사용해야 함
                        try:
                            btn_내pc = self.page.locator('button').filter(has_text="내 PC").first
                            async with self.page.expect_file_chooser(timeout=8000) as fc_info:
                                await btn_내pc.click()
                            fc = await fc_info.value
                            await fc.set_files(file_path)
                            await self.page.wait_for_timeout(2000)
                            logger.info(f"첨부완료(방법1-내PC버튼): {fname}")
                            uploaded = True
                        except Exception as e1:
                            logger.info(f"방법1 실패: {e1}")

                        # 방법 2: input.btn_fileAdd → Playwright set_input_files
                        # "내 PC" 버튼이 내부적으로 트리거하는 hidden input에 직접 파일 세팅
                        if not uploaded:
                            try:
                                await self.page.evaluate("""() => {
                                    const el = document.querySelector('input.btn_fileAdd, input.btn.btn_fileAdd');
                                    if (el) {
                                        el.style.cssText = 'display:block!important;position:fixed;top:0;left:0;width:1px;height:1px;opacity:0.01;z-index:9999';
                                    }
                                }""")
                                fi_loc2 = self.page.locator('input.btn_fileAdd').first
                                await fi_loc2.set_input_files(file_path, timeout=5000)
                                await self.page.evaluate("""() => {
                                    const el = document.querySelector('input.btn_fileAdd');
                                    if (el) {
                                        el.dispatchEvent(new Event('change', {bubbles: true}));
                                        el.dispatchEvent(new Event('input',  {bubbles: true}));
                                    }
                                }""")
                                await self.page.wait_for_timeout(2000)
                                logger.info(f"첨부완료(방법2-btn_fileAdd): {fname}")
                                uploaded = True
                            except Exception as e2:
                                logger.info(f"방법2 실패: {e2}")

                        # 방법 3: input#uploadFile → Playwright set_input_files
                        if not uploaded:
                            try:
                                await self.page.evaluate("""() => {
                                    const el = document.querySelector('input#uploadFile, input[name="uploadFile"]');
                                    if (el) {
                                        el.style.cssText = 'display:block!important;position:fixed;top:0;left:0;width:1px;height:1px;opacity:0.01;z-index:9999';
                                    }
                                }""")
                                fi_loc = self.page.locator('input#uploadFile, input[name="uploadFile"]').first
                                await fi_loc.set_input_files(file_path, timeout=5000)
                                await self.page.evaluate("""() => {
                                    const el = document.querySelector('input#uploadFile, input[name="uploadFile"]');
                                    if (el) {
                                        el.dispatchEvent(new Event('change', {bubbles: true}));
                                        el.dispatchEvent(new Event('input',  {bubbles: true}));
                                    }
                                }""")
                                await self.page.wait_for_timeout(2000)
                                logger.info(f"첨부완료(방법3-input#uploadFile): {fname}")
                                uploaded = True
                            except Exception as e3:
                                logger.info(f"방법3 실패: {e3}")

                        # ── 첨부 성공 여부: 첨부 영역에 항목이 추가됐는지 확인 ──
                        if uploaded:
                            await self.page.wait_for_timeout(500)
                            # ERP는 파일명을 축약 표시할 수 있으므로 첨부 행(체크박스+아이콘) 존재로 확인
                            confirmed = await self.page.evaluate("""() => {
                                // 첨부파일 목록 행: 체크박스가 있는 li/tr, 또는 파일 아이콘 요소
                                const area = document.querySelector(
                                    '.fileList, .attachFileList, [class*="fileList"], [class*="attachList"]'
                                );
                                if (area) return true;
                                // 체크박스+파일 아이콘 조합으로 확인
                                const rows = document.querySelectorAll(
                                    'input[type="checkbox"] ~ span, .attach_file_name, [class*="fileName"]'
                                );
                                return rows.length > 0;
                            }""")
                            if confirmed:
                                attached_files.append(fname)
                                logger.info(f"첨부 UI 확인 완료: {fname}")
                            else:
                                # 확인 불가해도 file chooser 성공이면 첨부된 것으로 처리
                                attached_files.append(fname)
                                logger.info(f"첨부 UI 확인 생략(file chooser 성공): {fname}")
                        else:
                            logger.warning(f"모든 첨부 방법 실패: {file_path}")
                            failed_files.append({"file": fname, "reason": "UI 첨부 실패"})

                    except Exception as att_e:
                        logger.warning(f"첨부파일 업로드 실패: {att_e}")
                        failed_files.append({"file": fname, "reason": str(att_e)})

                if attached_files:
                    await self.page.wait_for_timeout(1000)

            # ── 발송 버튼 클릭 ──
            send_btn = self.page.locator(".UDA0140 button[class*='themeblue']").filter(has_text="보내기").first
            await send_btn.click()
            try:
                await self.page.wait_for_selector(".listItem", timeout=8000)
            except Exception:
                await self.page.wait_for_timeout(2000)

            logger.info(f"메일 발송 완료: {mail_data['subject']}")

            # ── STEP 4: temp 임시 파일 삭제 ──
            deleted_temps = []
            for tmp_file in staged_temps:
                try:
                    if os.path.isfile(tmp_file):
                        os.remove(tmp_file)
                        deleted_temps.append(os.path.basename(tmp_file))
                        logger.info(f"임시파일 삭제: {tmp_file}")
                except Exception as del_e:
                    logger.warning(f"임시파일 삭제 실패: {tmp_file} - {del_e}")

            result = {"success": True, "message": f"메일이 발송되었습니다: {mail_data['subject']}"}
            if attached_files:
                result["attached_files"] = attached_files
            if deleted_temps:
                result["temp_cleaned"] = deleted_temps
            if failed_files:
                result["failed_attachments"] = failed_files
            return result
        except Exception as e:
            # 예외 발생 시에도 temp 파일 정리 시도
            for tmp_file in staged_temps:
                try:
                    if os.path.isfile(tmp_file):
                        os.remove(tmp_file)
                        logger.info(f"예외 후 임시파일 삭제: {tmp_file}")
                except Exception:
                    pass
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_mail_send_error.png"))
            return {"success": False, "error": str(e)}

    async def search_mail(self, keyword: str, folder: str = "inbox", page_num: int = 1) -> dict:
        """메일 검색 (더존 아마란스 검색 UI 사용)"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            await self._navigate_to_mail()

            # 검색창 찾기 (여러 후보 셀렉터 시도)
            search_input = None
            for sel in [
                'input[placeholder*="검색"]',
                'input[placeholder*="Search"]',
                '.searchBox input',
                '.OBTSearcher input',
                '.search-input input',
                'input[type="search"]',
            ]:
                el = await self.page.query_selector(sel)
                if el:
                    search_input = el
                    break

            if not search_input:
                return {"success": False, "error": "검색 입력창을 찾을 수 없습니다. (셀렉터 미발견)"}

            # JS로 직접 입력 + Enter 키 이벤트 발송 (visibility 기준 미충족 우회)
            await self.page.evaluate("""([el, keyword]) => {
                el.scrollIntoView({block:'center'});
                el.focus();
                const setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                setter.call(el, keyword);
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', keyCode:13, bubbles:true}));
                el.dispatchEvent(new KeyboardEvent('keypress', {key:'Enter', keyCode:13, bubbles:true}));
                el.dispatchEvent(new KeyboardEvent('keyup',   {key:'Enter', keyCode:13, bubbles:true}));
            }""", [search_input, keyword])
            # 검색 결과 로딩 대기 (셀렉터 기반, 최대 5초)
            try:
                await self.page.wait_for_selector(".listItem, .emptyList, .noData", timeout=5000)
            except Exception:
                await self.page.wait_for_timeout(1500)

            rows = await self.page.query_selector_all(".listItem")
            mails = await self._parse_mail_rows(rows)
            return {"success": True, "keyword": keyword, "folder": folder, "count": len(mails), "mails": mails}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ============================================================
    #  일정(캘린더) 기능
    # ============================================================

    async def _select_sidebar_calendar(self, calendar_name: str) -> bool:
        """좌측 사이드바에서 특정 캘린더만 체크/선택 (나머지는 해제)"""
        # 이미 같은 캘린더가 선택된 경우 즉시 반환
        if self._sched_calendar == calendar_name:
            logger.info(f"사이드바 캘린더 캐시 사용: {calendar_name}")
            return True
        try:
            # 방법 1: Playwright get_by_text 사용 (가장 안정적)
            try:
                cal_label = self.page.get_by_text(calendar_name, exact=True).first
                if await cal_label.count() > 0:
                    await cal_label.click()
                    logger.info(f"사이드바 캘린더 선택 완료 (get_by_text): {calendar_name}")
                    self._sched_calendar = calendar_name
                    return True
            except Exception:
                pass

            # 방법 2: JS로 텍스트 매칭 후 클릭
            clicked = await self.page.evaluate("""(calName) => {
                const selectors = [
                    '.calendarList label', '.lnbCalList label',
                    '[class*="calList"] label', '[class*="calGroup"] label',
                    '[class*="calendarGroup"] span', '[class*="calItem"] span',
                    '[class*="CalendarListItem"] span', '[class*="calendarItem"] span',
                    '[class*="lnb"] span', 'aside span', 'aside label',
                    '[class*="sideBar"] span', '[class*="sidebar"] span',
                    '[class*="leftPanel"] span', '[class*="LeftPanel"] span'
                ];
                for (const sel of selectors) {
                    const els = document.querySelectorAll(sel);
                    for (const el of els) {
                        const text = (el.innerText || el.textContent || '').trim();
                        if (text === calName || text.includes(calName)) {
                            el.click();
                            return {clicked: true, selector: sel, text};
                        }
                    }
                }
                // 마지막 수단: 모든 클릭 가능한 요소에서 텍스트 탐색
                const allEls = document.querySelectorAll('span, label, div, li');
                for (const el of allEls) {
                    const text = (el.childNodes.length <= 3 ? (el.innerText || el.textContent || '') : '').trim();
                    if (text === calName) {
                        el.click();
                        return {clicked: true, selector: 'fallback-text-match', text};
                    }
                }
                return {clicked: false};
            }""", calendar_name)

            if clicked.get("clicked"):
                logger.info(f"사이드바 캘린더 선택 완료 (JS): {calendar_name} / selector: {clicked.get('selector')}")
                self._sched_calendar = calendar_name
                return True
            else:
                logger.warning(f"사이드바에서 캘린더 '{calendar_name}'를 찾지 못했습니다.")
                return False
        except Exception as e:
            logger.warning(f"사이드바 캘린더 선택 중 오류 (무시하고 계속): {e}")
            return False

    async def _switch_to_list_view(self) -> bool:
        """FullCalendar 목록 뷰 전환 (통합 헬퍼)"""
        for btn_text in ["목록", "List", "list"]:
            try:
                btn = self.page.get_by_role("button", name=btn_text, exact=True).first
                if await btn.count() > 0:
                    await btn.click()
                    try:
                        await self.page.wait_for_selector(
                            ".fc-calendarListMonth-view tr, .fc-view tr", timeout=3000
                        )
                    except Exception:
                        pass
                    logger.info(f"목록 뷰 전환 (role=button, text={btn_text})")
                    return True
            except Exception:
                pass
        for btn_sel in [
            "button:has-text('목록')",
            ".fc-calendarListMonth-button",
            ".fc-listMonth-button",
            ".fc-list-button",
            "[class*='listMonth']",
            "[class*='listWeek']",
        ]:
            list_btn = await self.page.query_selector(btn_sel)
            if list_btn:
                await list_btn.click()
                try:
                    await self.page.wait_for_selector(
                        ".fc-calendarListMonth-view tr, .fc-view tr", timeout=3000
                    )
                except Exception:
                    pass
                logger.info(f"목록 뷰 전환 버튼 클릭: {btn_sel}")
                return True
        return False

    async def _get_current_calendar_month(self) -> str:
        """툴바/목록뷰 헤더에서 현재 표시 중인 연-월을 'YYYY-MM' 형식으로 반환.
        아마란스 목록뷰는 .fc-toolbar-title 이 없거나 다른 커스텀 요소를 사용하므로
        여러 셀렉터를 순서대로 시도한다.
        파싱 실패 시 빈 문자열 반환.
        """
        title = await self.page.evaluate("""() => {
            // 1) 표준 FullCalendar 그리드뷰 toolbar
            let t = document.querySelector('.fc-toolbar-title')?.innerText?.trim();
            if (t) return t;
            // 2) 아마란스 목록뷰 커스텀 헤더 — fc-header-toolbar 내 h2/span
            t = document.querySelector('.fc-header-toolbar h2, .fc-header-toolbar [class*="title"]')?.innerText?.trim();
            if (t) return t;
            // 3) .fc-toolbar 내 임의 제목 요소
            t = document.querySelector('.fc-toolbar h2, .fc-toolbar [class*="title"]')?.innerText?.trim();
            if (t) return t;
            // 4) 목록뷰 상단 날짜 레이블 (YYYY.MM 또는 YYYY년 M월 패턴)
            const allText = [...document.querySelectorAll('*')];
            for (const el of allText) {
                if (el.children.length > 0) continue;   // 리프 노드만
                const txt = (el.innerText || el.textContent || '').trim();
                if (/^\\d{4}[년\\.\\s]\\s*\\d{1,2}[월\\.\\s]?$/.test(txt)) return txt;
            }
            return '';
        }""")
        if not title:
            return ""
        # "2026년 3월" 또는 "2026. 3." 등 다양한 형식 대응
        m = re.search(r'(\d{4})[년\s\.]+(\d{1,2})', title)
        if m:
            year, month = int(m.group(1)), int(m.group(2))
            return f"{year}-{month:02d}"
        return ""

    async def _goto_schedule_date(self, date: str):
        """FullCalendar 날짜 이동 — prev/next 버튼 클릭 방식.

        gotoDate() JS 호출은 아마란스 ERP FullCalendar에서 동작하지 않으므로
        툴바의 '.fc-prev-button' / '.fc-next-button' 을 반복 클릭하여 이동한다.
        같은 월이면 즉시 반환(캐시 활용).
        """
        target_month = date[:7]   # "YYYY-MM"
        if self._sched_month == target_month:
            return

        # ── 현재 표시 월 파악 ──────────────────────────────────
        current_ym = await self._get_current_calendar_month()

        if current_ym:
            try:
                cur_year,  cur_mon  = int(current_ym[:4]),  int(current_ym[5:])
                tgt_year,  tgt_mon  = int(target_month[:4]), int(target_month[5:])
                months_diff = (tgt_year - cur_year) * 12 + (tgt_mon - cur_mon)
            except ValueError:
                months_diff = 0
        else:
            months_diff = 0

        logger.info(f"캘린더 월 이동: {current_ym} → {target_month} (diff={months_diff})")

        if months_diff != 0:
            btn_sel   = ".fc-next-button" if months_diff > 0 else ".fc-prev-button"
            btn_count = abs(months_diff)

            for i in range(btn_count):
                is_last_click = (i == btn_count - 1)
                try:
                    btn = self.page.locator(btn_sel).first
                    if await btn.count() > 0:
                        await btn.click()
                        if is_last_click:
                            # 마지막 클릭: 목표 월 데이터 로딩 완료까지 대기
                            # YY.MM 패턴 (예: 26.01) 또는 "데이터 없음" 요소 출현 대기
                            tgt_yymm = f"{target_month[2:4]}.{target_month[5:7]}"  # "26.01"
                            try:
                                await self.page.wait_for_function(
                                    f"""() => {{
                                        // 1) 타겟 월 날짜 패턴이 td에 나타나면 로딩 완료
                                        for (const td of document.querySelectorAll('tr td')) {{
                                            const t = (td.innerText || td.textContent || '').trim();
                                            if (t.startsWith('{tgt_yymm}')) return true;
                                        }}
                                        // 2) FullCalendar "no events" 메시지가 나타나도 완료로 간주
                                        if (document.querySelector(
                                            '.fc-list-empty, .fc-no-events, ' +
                                            '[class*="noEvent"], [class*="emptyMsg"], ' +
                                            '[class*="no-event"]'
                                        )) return true;
                                        return false;
                                    }}""",
                                    timeout=5000
                                )
                                logger.info(f"  [{i+1}/{btn_count}] 목표 월 데이터 로딩 감지")
                            except Exception:
                                # 타임아웃: 이벤트가 없거나 셀렉터 미매칭 → 추가 대기
                                logger.info(f"  [{i+1}/{btn_count}] 데이터 로딩 타임아웃 → 1.5s 추가 대기")
                                await self.page.wait_for_timeout(1500)
                        else:
                            # 중간 클릭: 짧은 대기 후 다음 클릭
                            await self.page.wait_for_timeout(700)
                    else:
                        logger.warning(f"  [{i+1}/{btn_count}] '{btn_sel}' 버튼 미발견")
                        break
                except Exception as e:
                    logger.warning(f"  [{i+1}/{btn_count}] 버튼 클릭 오류: {e}")
                    break

            # 실제 이동 결과 확인 (로그용)
            actual_ym = await self._get_current_calendar_month()
            logger.info(f"  이동 결과 확인: {actual_ym} (목표: {target_month})")

        self._sched_month = target_month

    async def list_schedule(self, date_from: str, date_to: str, calendar_name: str = "기술2팀") -> dict:
        """기간별 일정 조회 (FullCalendar 목록 뷰 활용) — 복수 월 범위 지원"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            await self._navigate_to_schedule()

            # 사이드바에서 캘린더 선택 (기본: 기술2팀)
            sidebar_selected = False
            if calendar_name:
                sidebar_selected = await self._select_sidebar_calendar(calendar_name)

            # 목록 뷰로 전환 시도 (이미 목록 뷰면 스킵)
            switched_to_list = (self._sched_view == 'list')
            if not switched_to_list:
                switched_to_list = await self._switch_to_list_view()
            if switched_to_list:
                self._sched_view = 'list'
            logger.info(f"목록 뷰 전환 여부: {switched_to_list}")

            # ── 월별 순회: date_from ~ date_to 사이 모든 월을 순서대로 방문 ──
            # FullCalendar listMonth 뷰는 1개월씩만 렌더링하므로
            # 복수 월에 걸친 범위는 월마다 gotoDate → DOM 스크래핑 후 병합한다.
            def _months_in_range(from_str: str, to_str: str) -> list:
                """YYYY-MM-DD 두 날짜 사이의 월 첫날 목록 반환 (YYYY-MM-DD 형식)"""
                cur = dt_date.fromisoformat(from_str).replace(day=1)
                end = dt_date.fromisoformat(to_str).replace(day=1)
                months = []
                while cur <= end:
                    months.append(cur.isoformat())
                    # 다음 달 1일로 이동 (dateutil 없이 순수 stdlib)
                    if cur.month == 12:
                        cur = cur.replace(year=cur.year + 1, month=1)
                    else:
                        cur = cur.replace(month=cur.month + 1)
                return months

            month_list = _months_in_range(date_from, date_to)
            logger.info(f"조회 대상 월 목록: {month_list}")

            # JS DOM 스크래핑 코드 (월별로 재사용)
            # Playwright page.evaluate(expr, arg) 는 arg가 1개 — 배열로 묶어서 전달
            _scrape_js = rf"""([dateFrom, dateTo]) => {{
                const results = [];
                let currentDate = '';
                const dateCounters = {{}};

                for (const tr of document.querySelectorAll('tr')) {{
                    const cells = [...tr.querySelectorAll('td')];
                    if (cells.length < 4) continue;
                    const t = cells.map(c => (c.innerText || c.textContent || '').trim());

                    // 헤더 행 스킵
                    if (t[0] === '일자' || t[0] === '날짜' || t[1] === '시간') continue;

                    let date, time, calendar, title, person, location;

                    if (cells.length >= 7) {{
                        // 7셀: c0=날짜(YY.MM.DD(요일)), c1=시간, c2=캘린더, c3=제목, c4=담당자, c5=장소
                        const m = t[0].match(/(\d{{2}})\.(\d{{2}})\.(\d{{2}})/);
                        if (m) currentDate = `20${{m[1]}}-${{m[2]}}-${{m[3]}}`;
                        date = currentDate; time = t[1]; calendar = t[2];
                        title = t[3]; person = t[4] || ''; location = t[5] || '';
                    }} else {{
                        // 6셀: c0=시간, c1=캘린더, c2=제목, c3=담당자, c4=장소
                        date = currentDate; time = t[0]; calendar = t[1];
                        title = t[2]; person = t[3] || ''; location = t[4] || '';
                    }}

                    if (!title) continue;
                    if (dateCounters[date] === undefined) dateCounters[date] = 0;
                    const idx = dateCounters[date]++;

                    if (dateFrom && date < dateFrom) continue;
                    if (dateTo   && date > dateTo)   continue;

                    results.push({{
                        event_id: `${{date}}:${{idx}}`,
                        title, start: date, time, calendar,
                        person, location, all_day: time === '종일'
                    }});
                }}
                return results;
            }}"""

            all_events: list = []
            seen_ids: set = set()  # 중복 제거 (월 경계 이벤트 대비)

            for month_start in month_list:
                # 해당 월로 이동
                self._sched_month = None  # 강제 이동 보장 (월 변경 시)
                await self._goto_schedule_date(month_start)
                logger.info(f"월 이동 완료: {month_start}")

                # arg를 배열 하나로 전달 (Playwright 규칙 준수)
                month_events = await self.page.evaluate(_scrape_js, [date_from, date_to])
                logger.info(f"{month_start} 스크래핑 결과: {len(month_events)}건")

                for ev in month_events:
                    eid = ev.get("event_id", "")
                    if eid not in seen_ids:
                        seen_ids.add(eid)
                        all_events.append(ev)

            # 날짜 오름차순 정렬
            all_events.sort(key=lambda e: (e.get("start", ""), e.get("time", "")))

            return {
                "success": True,
                "calendar": calendar_name or "전체",
                "calendar_selected": sidebar_selected,
                "list_view_switched": switched_to_list,
                "period": f"{date_from} ~ {date_to}",
                "months_traversed": month_list,
                "count": len(all_events),
                "events": all_events,
                "note": "더존 아마란스 캘린더에서 조회된 일정입니다."
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    async def read_schedule(self, date: str, event_index: int = 0, calendar_name: str = "기술2팀") -> dict:
        """특정 날짜의 이벤트를 클릭해 메모 포함 상세 조회"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            await self._navigate_to_schedule()
            if calendar_name:
                await self._select_sidebar_calendar(calendar_name)

            # 목록 뷰 전환 (이미 목록 뷰면 스킵)
            if self._sched_view != 'list':
                if await self._switch_to_list_view():
                    self._sched_view = 'list'

            # 해당 날짜로 이동 (같은 월이면 스킵)
            await self._goto_schedule_date(date)

            # 목록 뷰에서 target date에 해당하는 이벤트 행을 JS로 한 번에 추출
            # domIdx: querySelectorAll 결과 내 위치 (Python에서 바로 사용)
            row_info = await self.page.evaluate(rf"""() => {{
                const rows = document.querySelectorAll(
                    '.fc-calendarListMonth-view tr, .fc-view tr'
                );
                const datePattern = /(\d{{2}})\.(\d{{2}})\.(\d{{2}})/;
                let currentDate = '';
                const result = [];
                for (let domIdx = 0; domIdx < rows.length; domIdx++) {{
                    const row = rows[domIdx];
                    const cells = row.querySelectorAll('td');
                    if (cells.length < 4) continue;
                    const c0 = (cells[0].innerText || '').trim();
                    if (c0 === '일자' || c0 === '날짜' || c0 === '시간') continue;
                    if (cells.length >= 7) {{
                        const m = c0.match(datePattern);
                        if (m) currentDate = '20' + m[1] + '-' + m[2] + '-' + m[3];
                    }}
                    result.push({{domIdx, date: currentDate}});
                }}
                return result;
            }}""")

            target_entries = [r for r in row_info if r["date"] == date]
            total = len(row_info)
            if not target_entries:
                return {
                    "success": False,
                    "error": f"날짜 {date}에 해당하는 이벤트 없음 (전체 {total}개 중)",
                    "tip": "gw_list_schedule로 먼저 날짜 범위를 확인하세요."
                }
            if event_index >= len(target_entries):
                return {"success": False, "error": f"이벤트 인덱스 {event_index} 없음 ({date} 총 {len(target_entries)}개)"}

            # domIdx로 직접 행 선택 (Python 루프 없이)
            all_rows = await self.page.query_selector_all(
                ".fc-calendarListMonth-view tr, .fc-view tr"
            )
            target_row = all_rows[target_entries[event_index]["domIdx"]]
            await target_row.scroll_into_view_if_needed()
            try:
                await target_row.click(force=True)
            except Exception:
                await self.page.evaluate("(el) => el.click()", target_row)

            # 패널 출현 대기 (고정 2.5초 → selector 감지)
            try:
                await self.page.wait_for_selector(
                    ".pubScDetails .layer_div_in, .layer_wrap.pubScDetails",
                    timeout=5000
                )
            except Exception:
                # 패널 미출현 시 짧게 대기 후 계속
                await self.page.wait_for_timeout(500)

            # ── 우측 상세 뷰 패널에서 필드 추출 ──
            # 더존 아마란스: 뷰 패널은 x>900인 .layer_div_in (위치 기반으로 정확히 선택)
            detail = await self.page.evaluate("""() => {
                // x > 900 위치의 .layer_div_in 중 '일정 조회' 텍스트 포함 요소
                let panel = null;
                for (const el of document.querySelectorAll('.layer_div_in, .pubScDetails')) {
                    try {
                        const r = el.getBoundingClientRect();
                        if (r.left > 900 && r.width > 50) {
                            const t = (el.innerText || '').trim();
                            if (t.includes('일정 조회') || t.includes('제목')) {
                                panel = el;
                                break;
                            }
                        }
                    } catch(e) {}
                }
                // 폴백: x > 900인 첫 번째 .layer_div_in
                if (!panel) {
                    for (const el of document.querySelectorAll('.layer_div_in')) {
                        try {
                            const r = el.getBoundingClientRect();
                            if (r.left > 900) { panel = el; break; }
                        } catch(e) {}
                    }
                }

                const panelText = panel ? (panel.innerText || '').trim().substring(0, 1500) : '';

                let title = '', dateText = '', location = '', memo = '', attendees = [];
                // 섹션 구분 라벨 (메모 종료 기준)
                const SECTION_LABELS = new Set(['제목','일시','장소','메모','캘린더','공유','참여자','공개범위','최초등록','최종수정','삭제','수정','댓글','파일']);
                if (panel) {
                    const lines = (panel.innerText || '').split('\\n')
                        .map(l => l.trim()).filter(Boolean);
                    for (let i = 0; i < lines.length; i++) {
                        const l = lines[i];
                        const next = lines[i + 1] || '';
                        if (l === '제목' && next && !SECTION_LABELS.has(next)) {
                            const tLines = []; let j = i + 1;
                            while (j < lines.length && !SECTION_LABELS.has(lines[j])) { tLines.push(lines[j]); j++; }
                            title = tLines.join(' ');
                        }
                        if (l === '일시' && next && !SECTION_LABELS.has(next))   dateText = next;
                        if (l === '장소' && next && !SECTION_LABELS.has(next))   location = next;
                        if (l === '참여자' && next && !SECTION_LABELS.has(next)) attendees = [next];
                        // 메모: "메모" 라벨 다음부터 다음 섹션 라벨 전까지 수집
                        if ((l === '메모' || l === '내용' || l === '설명내용') && next) {
                            const memoLines = [];
                            let j = i + 1;
                            while (j < lines.length && !SECTION_LABELS.has(lines[j])) {
                                memoLines.push(lines[j]);
                                j++;
                            }
                            memo = memoLines.join('\\n');
                        }
                    }
                }

                return { title, dateText, location, memo, attendees, panelText };
            }""")

            # 메모가 비어 있으면 "댓글" 탭 클릭 후 재추출
            if not detail.get("memo"):
                clicked_tab = await self.page.evaluate("""() => {
                    const panel = document.querySelector('.pubScDetails, .layer_wrap.pubScDetails');
                    if (!panel) return false;
                    const tabs = panel.querySelectorAll('button, li, [role="tab"], a');
                    for (const tab of tabs) {
                        const t = (tab.innerText || '').trim();
                        if (t === '댓글' || t.startsWith('댓글') || t === '설명' || t.startsWith('설명')) {
                            tab.click();
                            return t;
                        }
                    }
                    return false;
                }""")
                if clicked_tab:
                    try:
                        await self.page.wait_for_selector(".pubScDetails textarea, .pubScDetails [role='tabpanel']", timeout=1500)
                    except Exception:
                        pass
                    memo_after_tab = await self.page.evaluate("""() => {
                        const panel = document.querySelector('.pubScDetails .layer_div_in, .pubScDetails');
                        if (!panel) return '';
                        const ta = panel.querySelector('textarea');
                        if (ta) return ta.value.trim() || (ta.innerText || '').trim();
                        // 활성 탭 패널 텍스트
                        const lines = (panel.innerText || '').split('\\n')
                            .map(l => l.trim()).filter(Boolean);
                        return lines.slice(0, 10).join(' | ');
                    }""")
                    if memo_after_tab:
                        detail["memo"] = memo_after_tab

            return {
                "success": True,
                "event_id": f"{date}:{event_index}",
                "event_index": event_index,
                "date": date,
                "calendar": calendar_name,
                "title": detail.get("title", ""),
                "date_text": detail.get("dateText", ""),
                "location": detail.get("location", ""),
                "memo": detail.get("memo", ""),
                "attendees": detail.get("attendees", []),
                "panel_text": detail.get("panelText", ""),
                "note": "메모가 비어 있으면 해당 일정에 메모가 없는 것입니다."
            }
        except Exception as e:
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_schedule_detail_error.png"))
            return {"success": False, "error": str(e)}

    async def _fill_memo_textarea(self, text: str) -> bool:
        """일정 폼의 메모 영역에 텍스트 입력.

        mouse.click 좌표 방식 미사용 — 잘못된 좌표 클릭으로 폼 패널이 닫히는 버그 방지.
        1순위: Playwright locator.fill() (React onChange 트리거, 저장 시 확실히 반영됨)
        2순위: JS native setter + event dispatch (locator 접근 불가 시 폴백)
        """
        # ── 1. DOM에서 x>1000 위치의 textarea/contenteditable 인덱스 조회 ──
        ta_info = await self.page.evaluate("""() => {
            const allTas = [...document.querySelectorAll('textarea')];
            for (let i = 0; i < allTas.length; i++) {
                const r = allTas[i].getBoundingClientRect();
                if (r.x > 1000 && r.width > 50) return {type: 'textarea', idx: i};
            }
            const allEds = [...document.querySelectorAll('[contenteditable]')];
            for (let i = 0; i < allEds.length; i++) {
                const r = allEds[i].getBoundingClientRect();
                if (r.x > 1000 && r.width > 50) return {type: 'contenteditable', idx: i};
            }
            return null;
        }""")

        logger.info(f"메모 영역 탐색: {ta_info}")

        # ── 2. Playwright locator.fill() — primary (mouse.click 사용 안 함) ──
        if ta_info is not None:
            try:
                sel = 'textarea' if ta_info['type'] == 'textarea' else '[contenteditable]'
                loc = self.page.locator(sel).nth(ta_info['idx'])
                await loc.click(timeout=2000)
                await loc.fill(text, timeout=5000)
                logger.info(f"메모 입력 완료 (locator fill, {ta_info['type']} idx={ta_info['idx']})")
                return True
            except Exception as e:
                logger.warning(f"locator fill 실패: {e}")

        # ── 3. JS native setter (React controlled component 폴백) ──
        filled = await self.page.evaluate("""(val) => {
            const targets = [...document.querySelectorAll('textarea,[contenteditable]')]
                .filter(el => { const r = el.getBoundingClientRect(); return r.x > 1000 && r.width > 50; });
            for (const el of targets) {
                try {
                    el.scrollIntoView({block: 'center'});
                    el.focus();
                    if (el.tagName === 'TEXTAREA') {
                        const setter = Object.getOwnPropertyDescriptor(
                            window.HTMLTextAreaElement.prototype, 'value').set;
                        setter.call(el, val);
                        el.dispatchEvent(new Event('focus', {bubbles: true}));
                        el.dispatchEvent(new Event('input', {bubbles: true}));
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                        el.dispatchEvent(new Event('blur', {bubbles: true}));
                    } else {
                        document.execCommand('selectAll');
                        document.execCommand('insertText', false, val);
                    }
                    return {ok: true, tag: el.tagName, cls: el.className.substring(0, 40)};
                } catch(e) {}
            }
            return {ok: false};
        }""", text)

        if filled.get("ok"):
            await self.page.wait_for_timeout(200)
            logger.info(f"메모 입력 완료 (JS setter: {filled})")
            return True

        logger.warning("메모 입력 실패: textarea/contenteditable 미발견 (x>1000)")
        if DEBUG_SCREENSHOTS:
            await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_memo_fail.png"))
        return False

    async def create_schedule(self, event_data: dict) -> dict:
        """일정 등록 (더존 아마란스 그룹웨어 일정 모듈)"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        # ── 사전 정리: 아마란스에 남아있는 모든 팝업/패널/모달 닫기 ──
        await self._close_all_popups()
        # 필수값 검증
        if not event_data.get("title", "").strip():
            return {"success": False, "error": "title(제목)은 필수 항목입니다."}
        # 날짜 형식 검증
        start_date_raw = event_data.get("start_date", "")
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', start_date_raw):
            return {"success": False, "error": f"start_date 형식 오류: '{start_date_raw}' (YYYY-MM-DD 필요)"}
        end_date_raw = event_data.get("end_date") or start_date_raw
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', end_date_raw):
            return {"success": False, "error": f"end_date 형식 오류: '{end_date_raw}' (YYYY-MM-DD 필요)"}
        if end_date_raw < start_date_raw:
            return {"success": False, "error": f"end_date({end_date_raw})가 start_date({start_date_raw})보다 앞입니다."}
        # 시간 형식 검증 (종일 일정 제외)
        if not event_data.get("all_day"):
            for tkey in ("start_time", "end_time"):
                tval = event_data.get(tkey, "")
                if tval and not re.match(r'^\d{2}:(00|30)$', tval):
                    return {"success": False, "error": f"{tkey} 형식 오류: '{tval}' (HH:00 또는 HH:30 형식, 30분 단위만 지원)"}
        try:
            # ── 캘린더 페이지로 이동 (_navigate_to_schedule: 이미 로드된 경우 즉시 반환) ──
            self._sched_calendar = None
            self._sched_view = None
            self._sched_month = None
            await self._navigate_to_schedule()

            # ── 배너/오버레이 제거 ──
            await self._hide_update_banner()
            await self._close_dialogs()

            # ── 혹시 남아있는 활성 패널 전부 닫기 (Escape 반복) ──
            for _ in range(3):
                active_count = await self.page.evaluate("""() =>
                    document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active').length
                """)
                if active_count == 0:
                    break
                # 패널 헤더의 닫기 버튼(X) 우선 클릭, 없으면 Escape
                closed = await self.page.evaluate("""() => {
                    for (const panel of document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active')) {
                        const closeBtn = panel.querySelector(
                            'button[class*="close"], button[aria-label*="닫기"], button[title*="닫기"], ' +
                            '.scLayerClose, .btn_close, .ic_close, button.close'
                        );
                        if (closeBtn) { closeBtn.click(); return true; }
                    }
                    return false;
                }""")
                if not closed:
                    await self.page.keyboard.press("Escape")
                await self.page.wait_for_timeout(200)
            await self.page.wait_for_timeout(200)

            # "일정 등록" 버튼 클릭 (button 요소 우선 — div 부모 클릭 방지)
            clicked = await self.page.evaluate("""() => {
                // 1순위: button 요소 (interactive 요소만)
                for (const el of document.querySelectorAll('button')) {
                    if ((el.innerText || '').trim() === '일정 등록') { el.click(); return 'button'; }
                }
                // 2순위: 리프 span/a (자식 없는 요소만)
                for (const el of document.querySelectorAll('span, a')) {
                    if (el.children.length === 0 && (el.innerText || '').trim() === '일정 등록') {
                        el.click(); return 'span/a';
                    }
                }
                return false;
            }""")
            if not clicked:
                await self.page.get_by_text("일정 등록", exact=True).first.click(timeout=3000)
            logger.info(f"일정 등록 버튼 클릭: method={clicked}")
            # 패널 열림 대기 — active 클래스 포함한 .pubScLayer 확인
            try:
                await self.page.wait_for_selector(".pubScLayer.active, .pubLayerSlide.active", timeout=4000)
            except Exception:
                # 폼이 안 열렸으면 Playwright get_by_text로 재시도
                logger.warning("패널 미열림 — get_by_text로 재클릭")
                try:
                    await self.page.get_by_text("일정 등록", exact=True).first.click(timeout=2000)
                    await self.page.wait_for_selector(".pubScLayer.active, .pubLayerSlide.active", timeout=3000)
                except Exception:
                    await self.page.wait_for_timeout(500)

            # ── 패널 "펼치기" 버튼 클릭 (폼이 접혀 있으면 확장) ──
            await self.page.evaluate("""() => {
                for (const el of document.querySelectorAll('button, span, a')) {
                    const t = (el.innerText || '').trim();
                    if (t === '펼치기') { el.click(); return; }
                }
            }""")

            # 폼 핵심 요소(제목 입력 or 캘린더 라벨)가 나타날 때까지 대기 (최대 8초)
            try:
                await self.page.wait_for_selector(
                    "#scTitleInput input, .pubScLayerIn td:has-text('캘린더')",
                    timeout=4000
                )
            except Exception:
                # 패널이 열리지 않은 경우 — JS 클릭 재시도
                logger.warning("일정 등록 패널 미열림, 버튼 재클릭 시도...")
                await self.page.evaluate("""() => {
                    for (const el of document.querySelectorAll('button, span, a, div')) {
                        if ((el.innerText || '').trim() === '일정 등록') { el.click(); return; }
                    }
                }""")
                try:
                    await self.page.wait_for_selector(
                        "#scTitleInput input, .pubScLayerIn td:has-text('캘린더'), .pubScLayer.active",
                        timeout=3000
                    )
                except Exception:
                    logger.warning("일정 등록 패널 재클릭 후에도 미열림")
                    await self.page.wait_for_timeout(500)

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_sched_form.png"))

            # ── 캘린더 선택 (기본: 기술2팀) ──
            cal_name = event_data.get("calendar_name", "기술2팀") or "기술2팀"
            cal_ok = await self._select_calendar(cal_name)
            if not cal_ok:
                logger.warning(f"캘린더 '{cal_name}' 선택 실패 — 기본 캘린더로 진행")
            # 캘린더 드롭다운이 열린 채 남아있으면 제목 영역 클릭으로 반드시 닫기
            await self.page.wait_for_timeout(400)
            await self._close_dialogs()
            await self.page.wait_for_timeout(300)

            # ── 제목 입력 (force=True로 hidden 상태에서도 입력) ──
            title_loc = self.page.locator("#scTitleInput input")
            try:
                await title_loc.wait_for(state="visible", timeout=2000)
                await title_loc.click(timeout=2000)
                await title_loc.fill(event_data["title"], timeout=5000)
            except Exception:
                # hidden 상태면 JS로 직접 값 설정
                await self.page.evaluate(f"""() => {{
                    const inp = document.querySelector('#scTitleInput input');
                    if (inp) {{
                        inp.removeAttribute('hidden');
                        inp.style.display = '';
                        inp.focus();
                        const nativeInputSetter = Object.getOwnPropertyDescriptor(
                            window.HTMLInputElement.prototype, 'value').set;
                        nativeInputSetter.call(inp, {repr(event_data['title'])});
                        inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                        inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                }}""")
            await self.page.wait_for_timeout(300)

            # ── 일시 섹션 펼치기 ──
            # '일시' 텍스트를 포함한 scUnitTop을 찾아 클릭 (여러 scUnitChild 중 정확한 것 선택)
            def _js_expand_date_section():
                return """() => {
                    // '일시' 라벨이 있는 scUnitTop 우선
                    for (const el of document.querySelectorAll('.scUnitTop')) {
                        if ((el.innerText || '').includes('일시')) { el.click(); return 'datetime'; }
                    }
                    // 없으면 날짜 input이 속한 scUnitBox의 scUnitTop 클릭
                    const dateInput = document.querySelector(
                        '.scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy, ' +
                        '.scUnitChild input[placeholder*="날짜"]'
                    );
                    if (dateInput) {
                        const box = dateInput.closest('.scUnitBox');
                        const top = box && box.querySelector('.scUnitTop');
                        if (top) { top.click(); return 'by_input'; }
                    }
                    // 최후: 2번째 scUnitBox
                    const el2 = document.querySelector('.scUnitBox:nth-child(2) .scUnitTop');
                    if (el2) { el2.click(); return 'nth2'; }
                    return false;
                }"""

            # 날짜 필드 visible 여부로 섹션 상태 판단 (count()는 hidden 포함이라 신뢰 불가)
            date_loc_primary = self.page.locator(".scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy")
            date_visible = await self.page.evaluate("""() => {
                const inputs = document.querySelectorAll(
                    '.scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy'
                );
                for (const inp of inputs) {
                    const r = inp.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0) return true;
                }
                return false;
            }""")
            if not date_visible:
                await self.page.evaluate(_js_expand_date_section())
                await self.page.wait_for_timeout(800)
                # visible 될 때까지 최대 3초 추가 대기
                try:
                    await date_loc_primary.first.wait_for(state="visible", timeout=2000)
                except Exception:
                    # 재시도
                    logger.warning("일시 섹션 펼치기 재시도...")
                    await self.page.evaluate(_js_expand_date_section())
                    await self.page.wait_for_timeout(600)

            # ── 종일 여부 ──
            if event_data.get("all_day"):
                try:
                    await self.page.click("label:has-text('종일')", timeout=3000)
                    await self.page.wait_for_timeout(500)
                except Exception:
                    pass

            # ── 날짜 입력 (OBTDatePickerRebuild - YYYYMMDD 형식으로 fill) ──
            start_date = event_data["start_date"]  # YYYY-MM-DD
            end_date = event_data.get("end_date") or start_date
            start_val = start_date.replace("-", "")
            end_val = end_date.replace("-", "")

            # 폼 패널 스코프로 한정 (페이지 상단 필터의 DatePicker 제외) — visible 여부는 form 내부 4개만 관심
            date_count = await self.page.evaluate("""() => {
                const panel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                if (!panel) return 0;
                const inputs = [...panel.querySelectorAll('.OBTDatePickerRebuild_inputYMD__PtxMy')];
                return inputs.filter(inp => {
                    const r = inp.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                }).length;
            }""")
            date_locators = self.page.locator(
                ".pubScLayer.active .OBTDatePickerRebuild_inputYMD__PtxMy, "
                ".pubLayerSlide.active .OBTDatePickerRebuild_inputYMD__PtxMy"
            )
            # DEBUG: 실제 매치된 요소 수 + 가시성
            date_dbg = await self.page.evaluate("""() => {
                const panel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                if (!panel) return {panel: false};
                const inputs = [...panel.querySelectorAll('.OBTDatePickerRebuild_inputYMD__PtxMy')];
                return {
                    panel: true,
                    count: inputs.length,
                    rects: inputs.map(i => {
                        const r = i.getBoundingClientRect();
                        return {w: Math.round(r.width), h: Math.round(r.height),
                                x: Math.round(r.x), y: Math.round(r.y), val: i.value};
                    })
                };
            }""")
            logger.info(f"[create] 폼 내부 DatePicker 상태: {date_dbg}")
            if date_count < 2:
                # 클래스명 변경 가능 — placeholder 기반 탐색
                date_count = await self.page.evaluate("""() => {
                    const inputs = [...document.querySelectorAll(
                        'input[placeholder*="날짜"], input[placeholder*="시작"], .scUnitChild input[type="text"]'
                    )];
                    return inputs.filter(inp => {
                        const r = inp.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    }).length;
                }""")
                date_locators = self.page.locator(
                    'input[placeholder*="날짜"], input[placeholder*="시작"], .scUnitChild input[type="text"]'
                )

            async def _fill_date(loc, val: str):
                """OBT 날짜 입력 — loc.focus() 직접 사용 (mouse.click 버블 회피).
                scUnitTop이 click을 가로채 섹션 collapse + 포커스 이전 문제를 피하기 위해
                Playwright의 locator.focus()로 포커스만 확실히 잡고 타이핑."""
                try:
                    await loc.focus(timeout=5000)
                except Exception as e:
                    logger.warning(f"_fill_date focus 실패: {e}")
                    # 폴백: JS focus()
                    await loc.evaluate("el => el.focus()")
                await self.page.wait_for_timeout(300)
                # 포커스 확인 디버그
                focused = await self.page.evaluate("""() => {
                    const a = document.activeElement;
                    return a ? {tag: a.tagName, cls: (a.className || '').substring(0, 50), val: a.value || ''} : null;
                }""")
                logger.info(f"[_fill_date] focus 확인: {focused}")
                await self.page.keyboard.press("Control+a")
                await self.page.keyboard.press("Delete")
                await self.page.keyboard.type(val, delay=50)
                await self.page.wait_for_timeout(300)
                await self.page.keyboard.press("ArrowDown")
                await self.page.wait_for_timeout(100)
                await self.page.keyboard.press("Enter")
                await self.page.wait_for_timeout(400)

            if date_count >= 2:
                await _fill_date(date_locators.nth(0), start_val)
                await _fill_date(date_locators.nth(1), end_val)
                # 실제 DOM 값 검증: OBTDatePickerRebuild는 "YYYY-MM-DD" 포맷 보관 (대시 포함)
                actual_dates = await self.page.evaluate("""() => {
                    const panel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                    if (!panel) return [];
                    const inputs = [...panel.querySelectorAll('.OBTDatePickerRebuild_inputYMD__PtxMy')];
                    return inputs.map(i => i.value);
                }""")
                logger.info(f"[create] 날짜 타이핑 후 값: {actual_dates}")
                if len(actual_dates) >= 2 and (actual_dates[0] != start_date or actual_dates[1] != end_date):
                    logger.warning(
                        f"날짜 입력 불일치(기대={start_date}/{end_date}, 실제={actual_dates[:2]}) "
                        f"— 재타이핑 재시도 (OBTDatePicker는 OBTComplete2 버퍼 → Enter 커밋 필요)"
                    )
                    # 재시도: 섹션 재전개 + 다시 mouse.click + type + Enter
                    await self.page.evaluate(_js_expand_date_section())
                    await self.page.wait_for_timeout(500)
                    try:
                        await _fill_date(date_locators.nth(0), start_val)
                        await _fill_date(date_locators.nth(1), end_val)
                    except Exception as e:
                        logger.warning(f"재시도 _fill_date 실패: {e}")
            else:
                logger.warning(f"날짜 visible 필드 미발견 (count={date_count}) — JS native setter 사용")
                await self.page.evaluate(f"""() => {{
                    const inputs = [...document.querySelectorAll(
                        '.scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy, ' +
                        '.scUnitChild input[type="text"], .scUnitChild input:not([type])'
                    )];
                    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    if (inputs[0]) {{
                        setter.call(inputs[0], '{start_val}');
                        inputs[0].dispatchEvent(new Event('input', {{bubbles: true}}));
                        inputs[0].dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    if (inputs[1]) {{
                        setter.call(inputs[1], '{end_val}');
                        inputs[1].dispatchEvent(new Event('input', {{bubbles: true}}));
                        inputs[1].dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                }}""")
                await self.page.wait_for_timeout(500)

            # ── 시간 입력 (OBTComplete2 드롭다운 클릭) ──
            time_warnings = []
            if not event_data.get("all_day"):
                start_time = event_data.get("start_time", "09:00")
                end_time = event_data.get("end_time", "10:00")
                ok_s = await self._select_time_from_dropdown("startTimeComplete", start_time)
                if not ok_s:
                    await self.page.wait_for_timeout(500)
                    ok_s = await self._select_time_from_dropdown("startTimeComplete", start_time)
                if not ok_s:
                    time_warnings.append(f"start_time={start_time} 선택 실패")
                ok_e = await self._select_time_from_dropdown("endTimeComlpete", end_time)
                if not ok_e:
                    await self.page.wait_for_timeout(500)
                    ok_e = await self._select_time_from_dropdown("endTimeComlpete", end_time)
                if not ok_e:
                    time_warnings.append(f"end_time={end_time} 선택 실패")
                # 시간 드롭다운이 열린 채 남아있으면 제목 입력 영역 클릭으로 닫기
                # Escape는 OBTComplete2 선택을 취소(기본값 복원)하므로 사용 금지
                await self.page.evaluate("""() => {
                    const inp = document.querySelector('#scTitleInput input');
                    if (inp) { inp.click(); return; }
                    const hdr = document.querySelector('.pubScLayerIn .scTitleWrap, .pubScLayerIn');
                    if (hdr) hdr.click();
                }""")
                await self.page.wait_for_timeout(600)   # 드롭다운 완전히 닫힐 시간 확보

            # ── 장소 입력 (scrollIntoView + force fill — hidden 상태에서도 동작) ──
            if event_data.get("location"):
                loc_val = event_data["location"]
                filled = await self.page.evaluate(f"""() => {{
                    const inp = document.querySelector('input[placeholder="장소/주소를 입력하세요."]');
                    if (!inp) return false;
                    inp.scrollIntoView({{block: 'center'}});
                    inp.removeAttribute('hidden');
                    inp.style.display = '';
                    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    setter.call(inp, {repr(loc_val)});
                    inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                    inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                    return true;
                }}""")
                if not filled:
                    try:
                        await self.page.locator('input[placeholder="장소/주소를 입력하세요."]').fill(
                            loc_val, timeout=2000, force=True
                        )
                    except Exception:
                        pass
                await self.page.wait_for_timeout(200)

            # ── 메모/설명 입력 ──
            if event_data.get("description"):
                await self._fill_memo_textarea(event_data["description"])

            # ── 배너 숨기고 남은 팝업/드롭다운 정리 후 "등록" 버튼 클릭 ──
            await self._hide_update_banner()
            await self._close_dialogs()
            await self.page.wait_for_timeout(400)

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_before_save.png"))

            # 등록 전 폼 상태 진단: 제목·날짜·시간·활성 패널 + 폼 패널 안의 모든 input 덤프
            # visible 필터 제거: 섹션이 collapsed되어도 OBTDatePicker value는 저장 시 사용됨
            form_state = await self.page.evaluate("""() => {
                const titleInp = document.querySelector('#scTitleInput input');
                const activePanel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                const panel = activePanel || document;
                const dateInputs = [...panel.querySelectorAll('.OBTDatePickerRebuild_inputYMD__PtxMy')];
                const startTimeInp = [...panel.querySelectorAll('#startTimeComplete input')][0];
                const endTimeInp = [...panel.querySelectorAll('#endTimeComlpete input')][0];
                // DEBUG: 폼 패널 안의 모든 input + 날짜/시간 패턴 값 덤프
                const panelInputs = [];
                if (activePanel) {
                    for (const inp of activePanel.querySelectorAll('input')) {
                        const r = inp.getBoundingClientRect();
                        const val = inp.value || '';
                        if (val && (val.match(/\\d{8}/) || val.match(/\\d{1,2}:\\d{2}/) || val.match(/\\d{4}[-.]\\d{2}/))) {
                            panelInputs.push({
                                id: inp.id || '',
                                name: inp.name || '',
                                ph: inp.placeholder || '',
                                cls: (inp.className || '').substring(0, 60),
                                val: val,
                                visible: r.width > 0 && r.height > 0,
                                x: Math.round(r.x), y: Math.round(r.y),
                                parentCls: (inp.parentElement?.className || '').substring(0, 60),
                            });
                        }
                    }
                }
                // DEBUG: 날짜 관련 모든 OBTDatePicker 요소 덤프 (visible 여부 불문)
                const allDatePickers = [...document.querySelectorAll(
                    '[class*="OBTDatePicker"] input, [class*="datePicker"] input'
                )].slice(0, 10).map(inp => {
                    const r = inp.getBoundingClientRect();
                    return {
                        id: inp.id || '',
                        cls: (inp.className || '').substring(0, 50),
                        val: inp.value || '',
                        visible: r.width > 0 && r.height > 0,
                        x: Math.round(r.x), y: Math.round(r.y),
                    };
                });
                return {
                    title: titleInp ? titleInp.value : null,
                    panelActive: !!activePanel,
                    dateCount: dateInputs.length,
                    date0: dateInputs[0] ? dateInputs[0].value : null,
                    date1: dateInputs[1] ? dateInputs[1].value : null,
                    startTime: startTimeInp ? startTimeInp.value : null,
                    endTime: endTimeInp ? endTimeInp.value : null,
                    _panelInputs: panelInputs,
                    _allDatePickers: allDatePickers,
                };
            }""")
            logger.info(f"등록 전 폼 상태: {form_state}")
            if not form_state.get("panelActive"):
                logger.warning("등록 전 폼 패널이 이미 닫혀 있음 — 등록 불가")
                return {"success": False, "error": "일정 등록 폼 패널이 닫혀 있음 (이전 작업 오류 가능성)"}

            # 날짜 검증: OBTDatePicker는 "YYYY-MM-DD" 포맷 보관 (대시 포함) — start_date와 직접 비교
            if form_state.get("dateCount", 0) == 0:
                await self._close_all_popups()
                return {
                    "success": False,
                    "error": (
                        f"날짜 입력 검증 불가: 폼 패널 내 OBTDatePicker input 없음. "
                        f"요청 날짜={start_date}."
                    ),
                    "form_state": form_state,
                }
            if form_state.get("date0") and form_state["date0"] != start_date:
                await self._close_all_popups()
                return {
                    "success": False,
                    "error": (
                        f"날짜 입력 실패: 요청={start_date}, "
                        f"실제 폼 값={form_state['date0']} — OBTDatePicker 상태 갱신 실패. "
                        f"등록을 중단했습니다(잘못된 날짜 저장 방지)."
                    ),
                    "form_state": form_state,
                }
            if form_state.get("date1") and form_state["date1"] != end_date:
                await self._close_all_popups()
                return {
                    "success": False,
                    "error": (
                        f"종료일 입력 실패: 요청={end_date}, "
                        f"실제 폼 값={form_state['date1']} — OBTDatePicker 상태 갱신 실패."
                    ),
                    "form_state": form_state,
                }

            # 활성 폼 패널 안의 "등록" 버튼 JS 클릭 (사이드바 "일정 등록" 버튼 혼동 방지)
            reg_clicked = await self.page.evaluate("""() => {
                // 1순위: 활성 슬라이드 패널 안에서 탐색 (부분 클래스 매칭 — CSS hash 불안정)
                for (const panel of document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active')) {
                    for (const btn of panel.querySelectorAll('button[class*="themeblue"]')) {
                        if ((btn.innerText || '').trim() === '등록') { btn.click(); return true; }
                    }
                }
                // 2순위: 페이지 전체 (정확히 "등록" 텍스트만)
                for (const btn of document.querySelectorAll('button[class*="themeblue"]')) {
                    if ((btn.innerText || '').trim() === '등록') { btn.click(); return true; }
                }
                return false;
            }""")
            if not reg_clicked:
                register_btn = self.page.locator('.pubScLayer.active button[class*="themeblue"], .pubLayerSlide.active button[class*="themeblue"]').filter(has_text="등록")
                await register_btn.first.click(force=True, timeout=5000)
            await self.page.wait_for_timeout(1000)  # 실패 다이얼로그 뜰 시간 확보

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_after_save.png"))

            # ── 실패 다이얼로그 감지 (광범위하게: 모든 알림성 텍스트 포함) ──
            fail_msg = await self.page.evaluate("""() => {
                // 팝업/다이얼로그 요소들
                const selectors = [
                    '[class*="OBTDialog_dialogRoot"]',
                    '[data-orbit-component="OBTDialog"]',
                    '.obtdialog',
                    '[class*="OBTDialog"]',
                    '[class*="dialog"]',
                    '[role="dialog"]',
                    '[role="alert"]',
                    '[class*="toast"]',
                    '[class*="Toast"]',
                ];
                for (const sel of selectors) {
                    for (const el of document.querySelectorAll(sel)) {
                        const t = (el.innerText || '').trim();
                        if (!t) continue;
                        if (t.includes('실패') || t.includes('오류') || t.includes('error') || t.includes('Error')
                            || t.includes('필수') || t.includes('입력하세요') || t.includes('입력해주세요')
                            || t.includes('확인해') || t.includes('잘못') || t.includes('invalid')) {
                            return t.substring(0, 300);
                        }
                    }
                }
                return null;
            }""")
            if fail_msg:
                if DEBUG_SCREENSHOTS:
                    await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_schedule_error.png"))
                await self._close_dialogs()
                return {"success": False, "error": f"일정 등록 실패 (앱 오류): {fail_msg}"}

            # ── 등록 완료 확인: 폼 패널(.pubScLayer.active)이 닫혔는지 대기 ──
            panel_closed = False
            try:
                await self.page.wait_for_function(
                    "() => !document.querySelector('.pubScLayer.active, .pubLayerSlide.active')",
                    timeout=5000
                )
                panel_closed = True
            except Exception:
                pass

            if not panel_closed:
                # 패널이 여전히 열려 있음 → 폼 제출 실패 (false positive 방지)
                if DEBUG_SCREENSHOTS:
                    await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_schedule_error.png"))
                # 어떤 오류 텍스트든 모두 수집
                any_msg = await self.page.evaluate("""() => {
                    const panel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                    return panel ? panel.innerText.substring(0, 500) : null;
                }""")
                logger.warning(f"폼 패널 미닫힘 — 등록 실패 가능성. 패널 내용: {any_msg}")
                await self._close_dialogs()
                return {"success": False, "error": f"일정 등록 실패: 폼이 닫히지 않음 (검증 필요). 패널 내용: {(any_msg or '')[:200]}"}

            logger.info(f"일정 등록 완료: {event_data['title']}")
            result = {
                "success": True,
                "message": f"일정이 등록되었습니다: {event_data['title']}",
                "calendar": cal_name,
                "calendar_set": cal_ok,
                "memo_set": bool(event_data.get("description")),
                "memo": event_data.get("description", ""),
                "registered_date": form_state.get("date0"),
                "registered_time": f"{form_state.get('startTime') or ''}~{form_state.get('endTime') or ''}",
            }
            if time_warnings:
                result["time_warnings"] = time_warnings
            return result
        except Exception as e:
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_schedule_error.png"))
            # 타임아웃/오류 시 열려 있는 폼 패널을 강제로 닫아 다음 호출에 영향 없도록 정리
            try:
                await self.page.reload(wait_until="domcontentloaded", timeout=15000)
            except Exception:
                pass
            return {"success": False, "error": str(e)}

    async def _click_schedule_row(self, event_id: str, calendar_name: str = "기술2팀") -> dict:
        """event_id("YYYY-MM-DD" 또는 "YYYY-MM-DD:N") 로 목록 뷰 이벤트 행 클릭 후 뷰 패널 열기.
        read_schedule와 동일한 방식으로 날짜+인덱스 기반 탐색."""

        # event_id 파싱: "2026-03-25" 또는 "2026-03-25:2"
        parts = event_id.split(":")
        date = parts[0].strip()
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', date):
            return {"success": False, "error": f"event_id 형식 오류: '{date}' (YYYY-MM-DD 또는 YYYY-MM-DD:N 필요)"}
        try:
            event_index = int(parts[1]) if len(parts) > 1 else 0
        except ValueError:
            return {"success": False, "error": f"event_id 인덱스 오류: '{parts[1]}' (정수 필요)"}

        await self._navigate_to_schedule()
        if calendar_name:
            await self._select_sidebar_calendar(calendar_name)

        # 목록 뷰 전환 (이미 목록 뷰면 스킵)
        if self._sched_view != 'list':
            if await self._switch_to_list_view():
                self._sched_view = 'list'

        # 날짜로 이동 (같은 월이면 스킵)
        await self._goto_schedule_date(date)

        # JS 단일 evaluate로 행 탐색 + bbox 반환 (Python 반복 제거로 대폭 단축)
        row_info = await self.page.evaluate("""([targetDate, eventIndex]) => {
            const rows = document.querySelectorAll('.fc-calendarListMonth-view tr, .fc-view tr');
            let curDate = '';
            const dateRows = [];
            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length < 4) continue;
                const c0 = (cells[0].innerText || '').trim();
                if (['일자','날짜','시간'].includes(c0)) continue;
                if (cells.length >= 7) {
                    const m = c0.match(/(\\d{2})\\.(\\d{2})\\.(\\d{2})/);
                    if (m) curDate = `20${m[1]}-${m[2]}-${m[3]}`;
                }
                if (curDate === targetDate) dateRows.push(row);
            }
            if (!dateRows.length) return {error: `날짜 ${targetDate} 이벤트 없음`};
            if (eventIndex >= dateRows.length)
                return {error: `인덱스 ${eventIndex} 없음 (총 ${dateRows.length}개)`};
            const r = dateRows[eventIndex].getBoundingClientRect();
            // 화면 밖이면 스크롤
            if (r.top < 0 || r.top > window.innerHeight)
                dateRows[eventIndex].scrollIntoView({block:'center'});
            const r2 = dateRows[eventIndex].getBoundingClientRect();
            return {x: r2.left + r2.width / 2, y: r2.top + r2.height / 2,
                    count: dateRows.length};
        }""", [date, event_index])

        if "error" in row_info:
            return {"success": False, "error": row_info["error"]}

        await self.page.mouse.click(row_info["x"], row_info["y"])
        try:
            await self.page.wait_for_selector(".pubScDetails .layer_div_in, .layer_wrap.pubScDetails", timeout=4000)
        except Exception:
            await self.page.wait_for_timeout(800)
        return {"success": True, "date": date, "event_index": event_index}

    async def update_schedule(self, event_id: str, updates: dict) -> dict:
        """일정 수정. event_id = 'YYYY-MM-DD' 또는 'YYYY-MM-DD:N' (N=목록 인덱스, 0부터)
        개인캘린더 포함 전체 일정 기준 인덱스를 사용하려면 calendar_name 없이 전체 뷰 사용."""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        # 날짜 형식 검증 (업데이트 항목이 있는 경우만)
        for key in ("start_date", "end_date"):
            val = updates.get(key)
            if val and not re.match(r'^\d{4}-\d{2}-\d{2}$', val):
                return {"success": False, "error": f"{key} 형식 오류: '{val}' (YYYY-MM-DD 필요)"}
        sd = updates.get("start_date")
        ed = updates.get("end_date")
        if sd and ed and ed < sd:
            return {"success": False, "error": f"end_date({ed})가 start_date({sd})보다 앞입니다."}
        # 시간 형식 검증 (30분 단위만 지원)
        for tkey in ("start_time", "end_time"):
            tval = updates.get(tkey, "")
            if tval and not re.match(r'^\d{2}:(00|30)$', tval):
                return {"success": False, "error": f"{tkey} 형식 오류: '{tval}' (HH:00 또는 HH:30 형식, 30분 단위만 지원)"}
        try:
            # ── 잔여 다이얼로그 먼저 닫기 ──
            try:
                for btn_name in ["아니요", "아니오", "확인", "닫기"]:
                    loc = self.page.get_by_role("button", name=btn_name, exact=True)
                    if await loc.count() > 0:
                        await loc.first.click(timeout=1500)
                        await self.page.wait_for_timeout(800)
                        break
            except Exception:
                pass

            # 기술2팀 필터 기준으로 인덱스 탐색 (gw_list_schedule과 동일 기준)
            click_result = await self._click_schedule_row(event_id, calendar_name="기술2팀")
            if not click_result["success"]:
                # 폴백: 전체 캘린더 뷰로 재시도
                click_result = await self._click_schedule_row(event_id, calendar_name="")
            if not click_result["success"]:
                return click_result

            # ── 편집 폼 진입: 뷰 패널 내 수정 버튼 클릭 (delete_schedule과 동일 방식) ──
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_detail_panel.png"))

            edit_btn = None
            for sel in [
                '.pubScDetails button:has-text("수정")',
                '.layer_div_in button:has-text("수정")',
                '.pubScLayer.active button:has-text("수정")',
                '.pubLayerSlide.active button:has-text("수정")',
                '.pubScDetails button[class*="themeblue"]',
                '.layer_div_in button[class*="themeblue"]',
            ]:
                el = await self.page.query_selector(sel)
                if el:
                    edit_btn = el
                    break

            if edit_btn:
                # JS click 사용 (visibility 체크 우회 — 오버레이/숨김 버튼 대응)
                await self.page.evaluate("(btn) => btn.click()", edit_btn)
            else:
                # 폴백: Playwright locator로 시도
                try:
                    await self.page.locator('.pubScDetails, .layer_div_in').get_by_text("수정", exact=True).first.click(force=True, timeout=3000)
                    edit_btn = True
                except Exception:
                    # 수정 버튼을 찾지 못함 — 진단 정보 수집 후 에러 반환 (좌표 폴백 없음)
                    panel_content = await self.page.evaluate("""() => {
                        const p = document.querySelector('.pubScDetails') ||
                                  document.querySelector('.pubScLayer.active') ||
                                  document.querySelector('.pubLayerSlide.active');
                        return p ? p.innerText.substring(0, 500) : 'panel not found';
                    }""")
                    return {"success": False, "error": f"수정 버튼을 찾을 수 없습니다. 패널 내용: {panel_content[:300]}"}

            await self.page.wait_for_timeout(1500)  # 편집 폼 완전 렌더링 대기

            # 편집 폼이 열렸는지 확인 (#scTitleInput input 만 사용 — 오탐 방지)
            has_form = await self.page.evaluate("""() =>
                !!document.querySelector('#scTitleInput input')
            """)
            # 편집 폼 DOM 구조 진단 (일시 섹션 펼치기 전)
            edit_form_diag = await self.page.evaluate("""() => {
                const info = {};
                // scUnitTop 목록
                info.scUnitTops = [...document.querySelectorAll('.scUnitTop')].map(el => ({
                    text: (el.innerText || '').substring(0, 50),
                    cls: el.className.substring(0, 80),
                    visible: el.getBoundingClientRect().width > 0
                }));
                // 날짜 입력 visible 상태
                info.dateInputs = [...document.querySelectorAll('.scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy')].map(el => ({
                    val: el.value,
                    w: Math.round(el.getBoundingClientRect().width)
                }));
                // 시간 드롭다운 버튼 visible 상태
                const startBtn = document.querySelector('#startTimeComplete [class*="OBTComplete2_dropDownButton"]');
                info.startTimeBtn = startBtn ? {
                    w: Math.round(startBtn.getBoundingClientRect().width),
                    h: Math.round(startBtn.getBoundingClientRect().height)
                } : null;
                // startTimeComplete input 값
                const startInp = document.querySelector('#startTimeComplete input');
                info.startTimeVal = startInp ? startInp.value : null;
                return info;
            }""")
            logger.info(f"편집 폼 DOM 진단: {edit_form_diag}")
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_edit_form_opened.png"))
            if not has_form:
                panel_content = await self.page.evaluate("""() => {
                    const p = document.querySelector('.pubScDetails') ||
                              document.querySelector('.pubScLayer.active') ||
                              document.querySelector('.pubLayerSlide.active');
                    return p ? p.innerText.substring(0, 500) : 'panel not found';
                }""")
                return {"success": False, "error": f"편집 폼 열기 실패. 패널 내용: {panel_content[:300]}"}

            # ── 제목 수정 ──
            inp_state = None
            if updates.get("title"):
                # 제목 입력: .last로 visible element 선택 (hidden/visible 2개 존재)
                title_loc = self.page.locator('input[placeholder="제목을 입력하세요."]').last
                try:
                    await title_loc.wait_for(state="visible", timeout=3000)
                    await title_loc.click(timeout=3000)
                    await title_loc.fill(updates["title"], timeout=5000)
                    inp_state = "fill_ok"
                except Exception as e:
                    logger.warning(f"제목 fill 실패: {e}")
                    # querySelectorAll로 visible element 찾아 native setter 사용
                    filled = await self.page.evaluate(f"""() => {{
                        for (const el of document.querySelectorAll('input[placeholder="제목을 입력하세요."]')) {{
                            const r = el.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0) {{
                                const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                                setter.call(el, {repr(updates['title'])});
                                el.dispatchEvent(new Event('input', {{bubbles: true}}));
                                el.dispatchEvent(new Event('change', {{bubbles: true}}));
                                return {{x: Math.round(r.x + r.width/2), y: Math.round(r.y + r.height/2)}};
                            }}
                        }}
                        return null;
                    }}""")
                    inp_state = f"native_setter:{'ok' if filled else 'not_found'}"

            # ── 날짜/시간 섹션 펼치기 ──
            date_locators = self.page.locator(".scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy")
            date_count = await date_locators.count()
            need_datetime_section = (
                updates.get("start_date") or updates.get("end_date")
                or updates.get("start_time") or updates.get("end_time")
            )
            # 화면에 보이는 날짜 입력 수 확인
            datetime_visible = await self.page.evaluate("""() => {
                const inputs = [...document.querySelectorAll(
                    '.scUnitChild .OBTDatePickerRebuild_inputYMD__PtxMy'
                )];
                return inputs.filter(el => {
                    const r = el.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                }).length;
            }""")
            expand_result = None
            if datetime_visible < 2 and need_datetime_section:
                # 일시 섹션이 접혀 있음 → 보이는 '일시' scUnitTop 좌표를 구해 mouse.click() 사용
                # ※ el.click()은 Vue.js 핸들러를 트리거하지 않으므로 반드시 mouse.click() 사용
                expand_coords = await self.page.evaluate("""() => {
                    for (const el of document.querySelectorAll('.scUnitTop')) {
                        if ((el.innerText || '').includes('일시')) {
                            const r = el.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0) {
                                return {
                                    found: true,
                                    x: Math.round(r.left + r.width / 2),
                                    y: Math.round(r.top + r.height / 2),
                                    cls: el.className.substring(0, 60)
                                };
                            }
                        }
                    }
                    return { found: false };
                }""")
                if expand_coords and expand_coords.get('found'):
                    await self.page.mouse.click(expand_coords['x'], expand_coords['y'])
                    expand_result = {
                        'result': 'expanded:mouse_click',
                        'x': expand_coords['x'],
                        'y': expand_coords['y'],
                        'cls': expand_coords.get('cls', '')
                    }
                else:
                    expand_result = {'result': 'expand_failed_no_element'}
                logger.info(f"일시 섹션 펼치기: {expand_result}")
                await self.page.wait_for_timeout(1000)
                if DEBUG_SCREENSHOTS:
                    await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_after_expand.png"))
                date_count = await date_locators.count()
            # OBTComplete2 시간 드롭다운 visible 여부 — 임의 버튼 가시성 확인
            time_visible = await self.page.evaluate("""() => {
                for (const btn of document.querySelectorAll('[class*="OBTComplete2_dropDownButton"]')) {
                    const r = btn.getBoundingClientRect();
                    if (r.width > 0 && r.height > 0) return true;
                }
                return false;
            }""")
            if updates.get("start_date") and date_count >= 1:
                await date_locators.nth(0).fill(updates["start_date"].replace("-", ""))
                await self.page.keyboard.press("Tab")
                await self.page.wait_for_timeout(300)
            if updates.get("end_date") and date_count >= 2:
                await date_locators.nth(1).fill(updates["end_date"].replace("-", ""))
                await self.page.keyboard.press("Tab")
                await self.page.wait_for_timeout(300)

            # ── 시간 수정 ──
            # 임시 진단: 실제 visible btn 목록 확인
            time_btn_debug = await self.page.evaluate("""() => {
                const all = [];
                for (const btn of document.querySelectorAll('[class*="OBTComplete2_dropDownButton"]')) {
                    const r = btn.getBoundingClientRect();
                    const anc = btn.closest('[id]');
                    all.push({
                        ancId: anc ? anc.id : 'none',
                        w: Math.round(r.width), h: Math.round(r.height),
                        x: Math.round(r.left + r.width/2), y: Math.round(r.top + r.height/2)
                    });
                }
                return all;
            }""")
            time_warnings = []
            if updates.get("start_time"):
                ok = await self._select_time_from_dropdown("startTimeComplete", updates["start_time"])
                if not ok:
                    # 1회 재시도
                    await self.page.wait_for_timeout(500)
                    ok = await self._select_time_from_dropdown("startTimeComplete", updates["start_time"])
                if not ok:
                    time_warnings.append(f"start_time={updates['start_time']} 선택 실패")
            if updates.get("end_time"):
                # start_time 변경 시 ERP 자동갱신 대기 + 섹션 재확인/재전개
                if updates.get("start_time"):
                    await self.page.wait_for_timeout(800)
                # 드롭다운 선택 후 scUnitTop 토글로 섹션이 닫혔을 수 있음 → 재전개
                end_btn_visible = await self.page.evaluate("""() => {
                    for (const el of document.querySelectorAll('#endTimeComlpete [class*="OBTComplete2_dropDownButton"]')) {
                        const r = el.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) return true;
                    }
                    return false;
                }""")
                if not end_btn_visible:
                    logger.info("endTimeComlpete 버튼 미노출 → 일시 섹션 재전개 시도")
                    expand2 = await self.page.evaluate("""() => {
                        for (const el of document.querySelectorAll('.scUnitTop')) {
                            if ((el.innerText || '').includes('일시')) {
                                const r = el.getBoundingClientRect();
                                if (r.width > 0 && r.height > 0)
                                    return {found: true,
                                            x: Math.round(r.left + r.width / 2),
                                            y: Math.round(r.top + r.height / 2)};
                            }
                        }
                        return {found: false};
                    }""")
                    if expand2 and expand2.get('found'):
                        await self.page.mouse.click(expand2['x'], expand2['y'])
                        await self.page.wait_for_timeout(800)
                ok = await self._select_time_from_dropdown("endTimeComlpete", updates["end_time"])
                if not ok:
                    # 1회 재시도
                    await self.page.wait_for_timeout(500)
                    ok = await self._select_time_from_dropdown("endTimeComlpete", updates["end_time"])
                if not ok:
                    time_warnings.append(f"end_time={updates['end_time']} 선택 실패")
            if updates.get("start_time") or updates.get("end_time"):
                # 드롭다운 닫기: 버튼 위쪽 중립 영역 클릭 (좌표 고정값 대신 상대 위치 사용)
                await self.page.evaluate("""() => {
                    const title = document.querySelector('#scTitleInput input');
                    if (title) { title.click(); return; }
                    // 폼 헤더 영역 클릭 (드롭다운 외부)
                    const header = document.querySelector('.pubScLayerIn .scTitleWrap, .pubScLayerIn');
                    if (header) { const r = header.getBoundingClientRect(); header.click(); }
                }""")
                await self.page.wait_for_timeout(400)
                # 저장 전 최종 검증
                # .pubLayerSlide.active 내부로 범위를 제한하여 중복 매칭 방지
                actual_start = await self.page.locator(
                    ".pubLayerSlide.active #startTimeComplete input, "
                    ".pubScLayer.active #startTimeComplete input, "
                    "#startTimeComplete input"
                ).first.input_value() if updates.get("start_time") else None
                actual_end = await self.page.locator(
                    ".pubLayerSlide.active #endTimeComlpete input, "
                    ".pubScLayer.active #endTimeComlpete input, "
                    "#endTimeComlpete input"
                ).first.input_value() if updates.get("end_time") else None
                if actual_start and actual_start != updates["start_time"]:
                    time_warnings.append(f"⚠️ start_time 불일치: 요청={updates['start_time']}, 실제={actual_start}")
                    logger.warning(time_warnings[-1])
                if actual_end and actual_end != updates["end_time"]:
                    time_warnings.append(f"⚠️ end_time 불일치: 요청={updates['end_time']}, 실제={actual_end}")
                    logger.warning(time_warnings[-1])

            # ── 장소 수정 (JS fill — hidden/invisible 상태에서도 동작) ──
            if updates.get("location"):
                loc_val = updates["location"]
                filled = await self.page.evaluate(f"""() => {{
                    const inp = document.querySelector('input[placeholder="장소/주소를 입력하세요."]');
                    if (!inp) return false;
                    inp.scrollIntoView({{block: 'center'}});
                    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    setter.call(inp, {repr(loc_val)});
                    inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                    inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                    return true;
                }}""")
                if not filled:
                    try:
                        await self.page.locator('input[placeholder="장소/주소를 입력하세요."]').fill(
                            loc_val, timeout=2000, force=True
                        )
                    except Exception:
                        pass

            # ── 메모 수정 ──
            if updates.get("description"):
                await self._fill_memo_textarea(updates["description"])

            # ── 저장 준비 ──
            await self._hide_update_banner()
            # ⚠️ _close_dialogs() 미사용: _dimClicker 클릭이 수정 폼의 React 이벤트 ID를
            # 초기화하여 저장 시 새 일정이 생성되는 버그를 유발함.
            # 드롭다운만 안전하게 닫기: 제목 입력란 클릭 (Escape 금지)
            await self.page.evaluate("""() => {
                const inp = document.querySelector('input[placeholder="제목을 입력하세요."]');
                if (inp && inp.getBoundingClientRect().width > 0) inp.click();
            }""")
            await self.page.wait_for_timeout(300)

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_before_save.png"))

            # 저장 전 폼 패널 활성화 확인
            # — 편집 폼은 #scTitleInput input 또는 .pubScLayer.active 중 하나로 감지
            panel_active = await self.page.evaluate("""() =>
                !!document.querySelector(
                    '#scTitleInput input, .pubScLayer.active, .pubLayerSlide.active'
                )
            """)
            if not panel_active:
                logger.warning("수정 저장 전 폼 패널이 닫혀 있음")
                return {"success": False, "error": "편집 폼 패널이 닫혀 있음 (수정 불가)"}

            # 저장 버튼 탐색 전략:
            # 1순위: .pubScLayer.active / .pubLayerSlide.active 내 '저장'/'수정'/'확인'
            # 2순위: #scTitleInput 를 포함하는 폼 컨테이너 내 themeblue 버튼
            # 3순위: 화면 전체 visible themeblue 버튼 중 저장/수정/확인
            save_clicked = False
            save_clicked = await self.page.evaluate("""() => {
                const labels = ['저장', '수정', '확인'];
                // 1순위: 활성 슬라이드 패널
                for (const panel of document.querySelectorAll('.pubScLayer.active, .pubLayerSlide.active')) {
                    for (const btn of panel.querySelectorAll('button[class*="themeblue"], button')) {
                        const t = (btn.innerText||'').trim();
                        if (labels.includes(t)) { btn.click(); return '1순위:' + t; }
                    }
                }
                // 2순위: #scTitleInput 폼 컨테이너 기반 탐색
                const titleEl = document.getElementById('scTitleInput');
                if (titleEl) {
                    let container = titleEl;
                    for (let i = 0; i < 12; i++) {
                        container = container.parentElement;
                        if (!container) break;
                        const cls = container.className || '';
                        if (cls.includes('layer_div') || cls.includes('pubSc') || cls.includes('layerSlide')) break;
                    }
                    if (container) {
                        for (const btn of container.querySelectorAll('button[class*="themeblue"], button')) {
                            const t = (btn.innerText||'').trim();
                            if (labels.includes(t) && btn.getBoundingClientRect().width > 0) {
                                btn.click(); return '2순위:' + t;
                            }
                        }
                    }
                }
                // 3순위: 화면 전체 visible themeblue 버튼
                for (const btn of document.querySelectorAll('button[class*="themeblue"]')) {
                    const t = (btn.innerText||'').trim();
                    const r = btn.getBoundingClientRect();
                    if (labels.includes(t) && r.width > 0) { btn.click(); return '3순위:' + t; }
                }
                return false;
            }""")
            await self.page.wait_for_timeout(1000)

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_after_save.png"))

            # ── 알림 발송 다이얼로그 처리: 좌표 기반 mouse.click (React 버튼 대응) ──
            dialog_coords = await self.page.evaluate("""() => {
                const targets = ['미발송', '아니요', '아니오', '발송', '보내기', '확인'];
                for (const target of targets) {
                    for (const btn of document.querySelectorAll('button')) {
                        const t = (btn.innerText || '').trim();
                        if (t === target) {
                            const r = btn.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0) {
                                return {x: r.x + r.width / 2, y: r.y + r.height / 2, label: t};
                            }
                        }
                    }
                }
                return null;
            }""")
            if dialog_coords:
                x, y = dialog_coords['x'], dialog_coords['y']
                # 해당 좌표에 실제로 어떤 element가 있는지 확인
                elem_info = await self.page.evaluate(f"() => {{ const el = document.elementFromPoint({x}, {y}); return el ? {{tag: el.tagName, text: (el.innerText||'').substring(0,50), cls: el.className}} : null; }}")
                logger.info(f"다이얼로그 처리: mouse.click ({dialog_coords['label']}) at ({x:.0f},{y:.0f}), element at point: {elem_info}")
                await self.page.mouse.click(x, y)
                await self.page.wait_for_timeout(800)
            else:
                # dialog_coords가 null — 버튼 목록 덤프
                btn_dump = await self.page.evaluate("""() => [...document.querySelectorAll('button')].map(b => ({t: (b.innerText||'').trim().substring(0,30), w: Math.round(b.getBoundingClientRect().width), h: Math.round(b.getBoundingClientRect().height)})).filter(b => b.w > 0)""")
                logger.warning(f"다이얼로그 버튼 미발견. 전체 visible 버튼: {btn_dump}")

            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_after_dialog.png"))

            # ── 저장 완료 확인: 활성 패널 소멸 대기 ──
            panel_closed = False
            try:
                await self.page.wait_for_function(
                    "() => !document.querySelector('.pubScLayer.active, .pubLayerSlide.active')",
                    timeout=5000
                )
                panel_closed = True
            except Exception:
                pass

            if not panel_closed:
                # 패널이 여전히 열려 있음 → 저장 실패 (false positive 방지)
                if DEBUG_SCREENSHOTS:
                    await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_error.png"))
                # 진단: 현재 페이지의 모든 visible 버튼 텍스트 수집
                diag = await self.page.evaluate("""() => {
                    const panel = document.querySelector('.pubScLayer.active, .pubLayerSlide.active');
                    const btns = [...document.querySelectorAll('button')].filter(b => b.getBoundingClientRect().width > 0).map(b => (b.innerText||'').trim().substring(0,20));
                    return {panel: panel ? panel.innerText.substring(0, 300) : null, btns: btns};
                }""")
                await self._close_dialogs()
                return {"success": False, "error": f"저장 실패. visible 버튼: {diag.get('btns', [])}", "panel": (diag.get('panel') or '')[:200]}

            result = {"success": True, "event_id": event_id, "message": "일정이 수정되었습니다.", "saved_via_button": save_clicked}
            if time_warnings:
                result["time_warnings"] = time_warnings
            result["_diag"] = {
                "edit_form_dom": edit_form_diag,
                "datetime_visible": datetime_visible,
                "expand_result": expand_result,
                "time_visible": time_visible,
                "time_btn_debug": time_btn_debug,
            }
            return result
        except Exception as e:
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_update_error.png"))
            # 타임아웃/오류 시 열려 있는 폼 패널을 강제로 닫아 다음 호출에 영향 없도록 정리
            try:
                await self.page.reload(wait_until="domcontentloaded", timeout=15000)
            except Exception:
                pass
            return {"success": False, "error": str(e)}

    async def delete_schedule(self, event_id: str) -> dict:
        """일정 삭제. event_id = 'YYYY-MM-DD' 또는 'YYYY-MM-DD:N' (N=목록 인덱스, 0부터)"""
        if not await self.ensure_logged_in():
            return {"success": False, "error": "로그인 실패"}
        try:
            # 기술2팀 필터 기준으로 인덱스 탐색 (gw_list_schedule과 동일 기준)
            click_result = await self._click_schedule_row(event_id, calendar_name="기술2팀")
            if not click_result["success"]:
                # 폴백: 전체 캘린더 뷰로 재시도
                click_result = await self._click_schedule_row(event_id, calendar_name="")
            if not click_result["success"]:
                return click_result

            # 뷰 패널의 "삭제" 버튼 클릭 (JS click — visibility 우회)
            delete_btn = None
            for sel in [
                '.pubScDetails button:has-text("삭제")',
                '.layer_div_in button:has-text("삭제")',
                '.pubScLayer.active button:has-text("삭제")',
                '.pubLayerSlide.active button:has-text("삭제")',
            ]:
                el = await self.page.query_selector(sel)
                if el:
                    delete_btn = el
                    break

            if delete_btn:
                await self.page.evaluate("(btn) => btn.click()", delete_btn)  # JS click — visibility 우회
            else:
                # 폴백: locator + 진단 정보 수집
                try:
                    await self.page.locator('.pubScDetails, .layer_div_in, .pubScLayer.active').get_by_text("삭제", exact=True).first.click(force=True, timeout=5000)
                except Exception:
                    panel_content = await self.page.evaluate("""() => {
                        const p = document.querySelector('.pubScDetails') ||
                                  document.querySelector('.pubScLayer.active') ||
                                  document.querySelector('.pubLayerSlide.active');
                        return p ? p.innerText.substring(0, 500) : 'panel not found';
                    }""")
                    return {"success": False, "error": f"삭제 버튼을 찾을 수 없습니다. 패널 내용: {panel_content[:300]}"}

            try:
                await self.page.wait_for_selector("[class*='OBTDialog_dialogRoot'], [class*='OBTDialog'], .obtdialog", timeout=3000)
            except Exception:
                await self.page.wait_for_timeout(500)

            # 확인 다이얼로그 처리
            for confirm_text in ["확인", "삭제", "예", "OK"]:
                try:
                    confirm_btn = self.page.get_by_text(confirm_text, exact=True).first
                    if await confirm_btn.is_visible(timeout=2000):
                        await confirm_btn.click()
                        break
                except Exception:
                    pass

            await self.page.wait_for_timeout(800)
            return {"success": True, "event_id": event_id, "message": "일정이 삭제되었습니다."}
        except Exception as e:
            if DEBUG_SCREENSHOTS:
                await self.page.screenshot(path=os.path.join(SCREENSHOT_DIR, "_delete_error.png"))
            return {"success": False, "error": str(e)}

    # ---------- 리소스 정리 ----------

    async def cleanup(self):
        try:
            if self.browser:
                await self.browser.close()
            if hasattr(self, "playwright") and self.playwright:
                await self.playwright.stop()
            logger.info("브라우저 리소스 정리 완료")
        except Exception as e:
            logger.error(f"정리 중 오류: {e}")


# ============================================================
# MCP 서버 정의
# ============================================================

@asynccontextmanager
async def app_lifespan(app):
    client = GroupwareClient()
    await client.initialize()
    yield {"gw_client": client}
    await client.cleanup()

mcp = FastMCP("erp_groupware_mcp", lifespan=app_lifespan)




# ============================================================
# 일정 메모 구조화 헬퍼
# ============================================================

def _build_memo_text(
    sales: Optional[str] = None,
    customer: Optional[str] = None,
    project: Optional[str] = None,
    target: Optional[str] = None,
    content: Optional[str] = None,
    description: Optional[str] = None,
) -> Optional[str]:
    """구조화 메모 항목(영업/고객/사업/대상/내용)을 포맷된 문자열로 조합.

    구조화 항목이 하나라도 있으면 아래 형식으로 반환:
        영업: ...
        고객: ...
        사업: ...
        대상: ...
        내용: ...
    구조화 항목이 없으면 description을 그대로 반환.
    """
    fields = [
        ("영업", sales),
        ("고객", customer),
        ("사업", project),
        ("대상", target),
        ("내용", content),
    ]
    structured_lines = [f"{label}: {val}" for label, val in fields if val]
    if not structured_lines:
        return description or None
    if description:
        structured_lines.append(description)
    return "\n".join(structured_lines)


def _parse_memo_fields(memo_text: str) -> dict:
    """구조화 메모 텍스트에서 영업/고객/사업/대상/내용 필드를 파싱."""
    fields: dict = {}
    extras: list = []
    for line in memo_text.split("\n"):
        m = re.match(r'^(영업|고객|사업|대상|내용)\s*[:\uff1a]\s*(.*)$', line.strip())
        if m:
            fields[m.group(1)] = m.group(2).strip()
        elif line.strip():
            extras.append(line.strip())
    if extras:
        fields["기타"] = "\n".join(extras)
    return fields


# ============================================================
# MCP 도구 - 메일
# ============================================================

async def _reset_browser_impl(client) -> int:
    """브라우저를 홈으로 이동하고 모달/오버레이를 제거합니다. 제거된 요소 수를 반환합니다."""
    await client.page.goto(ERP_BASE_URL, wait_until="domcontentloaded", timeout=30000)
    await client.page.wait_for_timeout(3000)
    if "/login" in client.page.url:
        await client.login()
    removed = await client.page.evaluate("""() => {
        let count = 0;
        const banner = document.querySelector('.systemAlertUpdate');
        if (banner) { banner.style.display = 'none'; count++; }
        const loading = document.getElementById('a10-first-loading');
        if (loading) { loading.style.display = 'none'; count++; }
        for (const el of document.querySelectorAll(
            '[class*="loading"], [class*="Loading"], [class*="overlay"], [class*="Overlay"], ' +
            '[class*="spinner"], [class*="Spinner"], [class*="mask"], [class*="Mask"], ' +
            '[class*="modal"], [class*="Modal"], [class*="dialog"], [class*="Dialog"]'
        )) {
            const s = window.getComputedStyle(el);
            if ((s.position === 'fixed' || s.position === 'absolute')) {
                const r = el.getBoundingClientRect();
                if (r.width > 200 && r.height > 200) {
                    el.style.display = 'none'; count++;
                }
            }
        }
        for (const el of document.querySelectorAll('*')) {
            try {
                const s = window.getComputedStyle(el);
                const z = parseInt(s.zIndex, 10);
                if ((s.position === 'fixed' || s.position === 'absolute') && z > 9000) {
                    const r = el.getBoundingClientRect();
                    if (r.width > 400 && r.height > 400 && s.pointerEvents !== 'none') {
                        el.style.pointerEvents = 'none'; count++;
                    }
                }
            } catch(e) {}
        }
        const btns = document.querySelectorAll(
            '[class*="OBTDialog_dialogRoot"] button, [data-orbit-component="OBTDialog"] button'
        );
        for (const btn of btns) {
            const txt = (btn.innerText || '').trim();
            if (['확인','닫기','취소','OK','Close'].includes(txt)) { btn.click(); count++; break; }
        }
        return count;
    }""")
    client._sched_calendar = None
    client._sched_view = None
    client._sched_month = None
    return removed


@mcp.tool(
    name="gw_reset_browser",
    annotations={"title": "브라우저 상태 초기화", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_reset_browser(ctx: Context) -> str:
    """Playwright 브라우저를 홈페이지로 이동하고 모든 모달/오버레이/팝업을 강제 제거합니다.
    ElementHandle.click Timeout 등 브라우저 상태 문제가 반복될 때 호출하세요.

    Returns:
        str: JSON - 초기화 결과
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    try:
        removed = await _reset_browser_impl(client)
        return json.dumps({
            "success": True,
            "message": f"브라우저 초기화 완료 (제거된 blocking 요소: {removed}개)",
            "url": client.page.url
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)}, ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_check_login",
    annotations={"title": "그룹웨어 로그인 확인", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_check_login(ctx: Context) -> str:
    """그룹웨어 로그인 상태를 확인하고, 필요 시 자동 로그인합니다.

    Returns:
        str: JSON - 로그인 성공/실패 상태
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    # ensure_logged_in: 이미 로그인된 경우 재로그인 없이 True 반환
    success = await client.ensure_logged_in()
    if success:
        return json.dumps({"success": True, "message": "로그인 성공"}, ensure_ascii=False, indent=2)
    # 세션 만료 시 강제 재로그인 시도
    return json.dumps(await client.login(), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_list_inbox",
    annotations={"title": "수신 메일 목록 조회", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_list_inbox(ctx: Context, page: int = 1, per_page: int = 20) -> str:
    """수신함의 메일 목록을 조회합니다. 발신자, 제목, 날짜, 읽음 여부를 반환합니다.

    Args:
        page: 페이지 번호 (기본 1)
        per_page: 페이지당 표시 수 (기본 20)

    Returns:
        str: JSON - 메일 목록 [{mail_id, from, subject, date, is_read, has_attachment}]
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(await client.list_inbox(page, per_page), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_read_mail",
    annotations={"title": "메일 상세 읽기", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_read_mail(ctx: Context, mail_id: str = "") -> str:
    """메일 ID로 메일의 전체 내용(제목, 본문, 첨부파일 등)을 읽습니다.

    Args:
        mail_id: 메일 고유 ID

    Returns:
        str: JSON - 메일 상세 {from, to, cc, subject, body, attachments}
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(await client.read_mail(mail_id), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_send_mail",
    annotations={"title": "메일 발송", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True}
)
async def gw_send_mail(ctx: Context, to: str = "", subject: str = "", body: str = "",
                       cc: Optional[str] = None,
                       attachments: Optional[List[str]] = None) -> str:
    """그룹웨어를 통해 메일을 발송합니다.

    [파일 첨부 흐름]
      1. gw_list_recent_files 로 최근 생성/다운로드 파일 목록 확인
      2. 해당 파일명 또는 경로를 attachments에 전달
      3. 서버가 자동으로 C:/mcp/temp 로 복사 → 첨부 → 발송 후 temp 파일 삭제

    첨부파일 자동 탐색 순서:
      1. Downloads 폴더  ← 최우선
      2. Desktop, Documents
      3. C:/mcp/temp 공유 폴더
      4. Claude Desktop outputs / uploads 폴더 (생성된 파일 포함)
    파일을 찾지 못하면 발송을 중단하고 위치 확인을 요청합니다.

    Args:
        to: 수신자 이메일 (여러 명은 ; 구분)
        subject: 메일 제목
        body: 메일 본문
        cc: 참조 수신자 (여러 명은 ; 구분, 선택)
        attachments: 첨부파일 경로 목록. Windows 경로(C:/...) 또는 파일명만 지정 가능.
                     파일명만 지정하면 Downloads 폴더부터 자동 검색합니다.
                     예: ["C:/Users/jssul/Downloads/report.xlsx"] 또는 ["report.xlsx"]

    Returns:
        str: JSON - 발송 결과. 파일 미발견 시 needs_file_location=true 반환.
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    mail_data = {"to": to, "subject": subject, "body": body}
    if cc:
        mail_data["cc"] = cc
    if attachments:
        mail_data["attachments"] = attachments
    return json.dumps(await client.send_mail(mail_data), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_search_mail",
    annotations={"title": "메일 검색", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_search_mail(ctx: Context, keyword: str = "", folder: str = "inbox", page: int = 1) -> str:
    """키워드로 메일을 검색합니다. 수신함/발신함/임시보관함/휴지통에서 검색 가능합니다.

    Args:
        keyword: 검색 키워드 (제목/발신자/본문)
        folder: 검색 폴더 (inbox, sent, draft, trash)
        page: 페이지 번호

    Returns:
        str: JSON - 검색 결과 메일 목록
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(await client.search_mail(keyword, folder, page), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_list_recent_files",
    annotations={"title": "최근 생성 파일 목록 조회", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
)
async def gw_list_recent_files(ctx: Context, hours: int = 24, limit: int = 10) -> str:
    """Claude Desktop이 생성한 결과 파일 및 최근 다운로드 파일 목록을 조회합니다.

    메일에 파일을 첨부할 때, 어떤 파일이 사용 가능한지 먼저 이 도구로 확인하세요.
    조회된 파일의 filename 또는 path를 gw_send_mail의 attachments 파라미터에 전달하면
    자동으로 탐색하여 C:/mcp/temp로 복사 후 첨부됩니다.

    검색 위치 (최신 파일 우선):
      - Claude Desktop 세션 outputs 폴더 (생성된 Excel/PDF/Word 등)
      - Claude Desktop pending-uploads 폴더
      - 로컬 PC Downloads 폴더
      - C:/mcp/temp 공유 폴더

    Args:
        hours: 최근 N시간 이내 파일만 조회 (기본 24시간)
        limit: 최대 반환 파일 수 (기본 10)

    Returns:
        str: JSON - 최근 파일 목록 [{filename, path, source, modified, size_kb}]
             source 값: claude_outputs(Claude생성), claude_pending(업로드중),
                        downloads(다운로드폴더), mcp_temp(공유폴더)
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(client.list_recent_output_files(hours, limit), ensure_ascii=False, indent=2)


# ============================================================
# MCP 도구 - 일정
# ============================================================

@mcp.tool(
    name="gw_list_schedule",
    annotations={"title": "일정 조회", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_list_schedule(ctx: Context, date_from: str = "", date_to: str = "",
                           calendar_name: str = "기술2팀") -> str:
    """지정 기간의 일정을 조회합니다. 제목, 시간, 장소, 참석자 정보를 반환합니다.

    Args:
        date_from: 조회 시작일 (YYYY-MM-DD)
        date_to: 조회 종료일 (YYYY-MM-DD)
        calendar_name: 조회할 캘린더 이름 (기본: '기술2팀'). 전체 조회 시 빈 문자열 전달.

    Returns:
        str: JSON - 일정 목록 [{event_id, title, start, end, location, all_day}]
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(await client.list_schedule(date_from, date_to, calendar_name), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_read_schedule",
    annotations={"title": "일정 상세 조회 (메모 포함)", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_read_schedule(ctx: Context, date: str = "", event_index: int = 0,
                            calendar_name: str = "기술2팀") -> str:
    """특정 날짜의 이벤트를 클릭해 메모/설명을 포함한 상세 정보를 조회합니다.
    먼저 gw_list_schedule로 이벤트 목록을 확인 후, 보고 싶은 이벤트의 순서(0부터)를 event_index로 전달하세요.

    Args:
        date: 조회할 날짜 (YYYY-MM-DD)
        event_index: 해당 날짜의 이벤트 순서 (0부터, 기본 0)
        calendar_name: 캘린더 이름 (기본: '기술2팀')

    Returns:
        str: JSON - 이벤트 상세 {title, start_date, end_date, location, memo, attendees}
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    result = await client.read_schedule(date, event_index, calendar_name)
    # 구조화 메모 파싱 (영업/고객/사업/대상/내용)
    if result.get("success") and result.get("memo"):
        parsed = _parse_memo_fields(result["memo"])
        if parsed:
            result["memo_fields"] = parsed
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_create_schedule",
    annotations={"title": "일정 등록", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True}
)
async def gw_create_schedule(ctx: Context, title: str = "", start_date: str = "", start_time: str = "09:00",
                              end_date: Optional[str] = None, end_time: str = "10:00",
                              location: Optional[str] = None,
                              sales: Optional[str] = None,
                              customer: Optional[str] = None,
                              project: Optional[str] = None,
                              target: Optional[str] = None,
                              content: Optional[str] = None,
                              description: Optional[str] = None,
                              attendees: Optional[List[str]] = None, all_day: bool = False,
                              reminder_minutes: Optional[int] = None,
                              calendar_name: str = "기술2팀") -> str:
    """그룹웨어 캘린더에 새 일정을 등록합니다. 기본 캘린더: 기술2팀.

    메모는 아래 구조화 항목으로 입력합니다 (모두 선택):
        영업(sales), 고객(customer), 사업(project), 대상(target), 내용(content)
    입력된 항목은 다음 형식으로 메모에 자동 저장됩니다:
        영업: ...
        고객: ...
        사업: ...
        대상: ...
        내용: ...

    Args:
        title: 일정 제목
        start_date: 시작 날짜 (YYYY-MM-DD)
        start_time: 시작 시간 (HH:MM, 기본 09:00)
        end_date: 종료 날짜 (YYYY-MM-DD, 미입력 시 시작일과 동일)
        end_time: 종료 시간 (HH:MM, 기본 10:00)
        location: 장소
        sales: 메모 - 영업 담당자
        customer: 메모 - 고객사명
        project: 메모 - 사업
        target: 메모 - 대상
        content: 메모 - 업무 내용
        description: 메모 - 기타 자유형식 추가 내용 (구조화 항목과 함께 사용 가능)
        attendees: 참석자 이메일 목록
        all_day: 종일 일정 여부
        reminder_minutes: 사전 알림 (분 단위)
        calendar_name: 등록할 캘린더 이름 (기본: '기술2팀').

    Returns:
        str: JSON - 생성된 일정 ID 및 결과
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    await _reset_browser_impl(client)
    event_data = {"title": title, "start_date": start_date, "start_time": start_time, "end_time": end_time}
    if end_date:
        event_data["end_date"] = end_date
    if location:
        event_data["location"] = location
    memo = _build_memo_text(sales, customer, project, target, content, description)
    if memo:
        event_data["description"] = memo
    if attendees:
        event_data["attendees"] = attendees
    if all_day:
        event_data["all_day"] = all_day
    if reminder_minutes is not None:
        event_data["reminder_minutes"] = reminder_minutes
    event_data["calendar_name"] = calendar_name
    return json.dumps(await client.create_schedule(event_data), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_update_schedule",
    annotations={"title": "일정 수정", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True}
)
async def gw_update_schedule(ctx: Context, event_id: str = "", title: Optional[str] = None,
                              start_date: Optional[str] = None, start_time: Optional[str] = None,
                              end_date: Optional[str] = None, end_time: Optional[str] = None,
                              location: Optional[str] = None,
                              sales: Optional[str] = None,
                              customer: Optional[str] = None,
                              project: Optional[str] = None,
                              target: Optional[str] = None,
                              content: Optional[str] = None,
                              description: Optional[str] = None,
                              attendees: Optional[List[str]] = None) -> str:
    """기존 일정의 정보를 수정합니다. 변경할 항목만 입력하면 됩니다.

    메모는 구조화 항목으로 수정합니다 (변경할 항목만 입력):
        영업(sales), 고객(customer), 사업(project), 대상(target), 내용(content)
    구조화 항목을 하나라도 입력하면 메모 전체가 아래 형식으로 덮어써집니다:
        영업: ...
        고객: ...
        사업: ...
        대상: ...
        내용: ...

    Args:
        event_id: 대상 일정. 'YYYY-MM-DD' (첫 번째 일정) 또는 'YYYY-MM-DD:N' (N번째, 0부터).
                  gw_list_schedule로 날짜와 순서를 먼저 확인하세요.
        title: 변경할 제목
        start_date: 변경할 시작 날짜 (YYYY-MM-DD)
        start_time: 변경할 시작 시간 (HH:MM)
        end_date: 변경할 종료 날짜 (YYYY-MM-DD)
        end_time: 변경할 종료 시간 (HH:MM)
        location: 변경할 장소
        sales: 메모 - 영업 담당자
        customer: 메모 - 고객사명
        project: 메모 - 사업
        target: 메모 - 대상
        content: 메모 - 업무 내용
        description: 메모 - 기타 자유형식 추가 내용
        attendees: 변경할 참석자 목록

    Returns:
        str: JSON - 수정 결과
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    await _reset_browser_impl(client)
    updates = {}
    for key, val in [("title", title), ("start_date", start_date), ("start_time", start_time),
                     ("end_date", end_date), ("end_time", end_time), ("location", location),
                     ("attendees", attendees)]:
        if val is not None:
            updates[key] = val
    memo = _build_memo_text(sales, customer, project, target, content, description)
    if memo:
        updates["description"] = memo
    return json.dumps(await client.update_schedule(event_id, updates), ensure_ascii=False, indent=2)


@mcp.tool(
    name="gw_delete_schedule",
    annotations={"title": "일정 삭제", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": True}
)
async def gw_delete_schedule(ctx: Context, event_id: str = "") -> str:
    """일정을 삭제합니다. 삭제 전 확인 다이얼로그를 자동 처리합니다.

    Args:
        event_id: 삭제할 일정. 'YYYY-MM-DD' (첫 번째 일정) 또는 'YYYY-MM-DD:N' (N번째, 0부터).
                  gw_list_schedule로 날짜와 순서를 먼저 확인하세요.

    Returns:
        str: JSON - 삭제 결과
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    return json.dumps(await client.delete_schedule(event_id), ensure_ascii=False, indent=2)


# ============================================================
# MCP 도구 - 일정 DOM 진단 (임시)
# ============================================================

@mcp.tool(
    name="gw_diag_schedule",
    annotations={"title": "[진단] 일정 DOM 구조 분석", "readOnlyHint": True, "destructiveHint": False}
)
async def gw_diag_schedule(ctx: Context, target_date: str = "2026-01-01") -> str:
    """[진단용] 일정 페이지의 실제 DOM 구조, 목록뷰 전환 상태, gotoDate 동작,
    tr/td 셀 구조를 분석하여 반환합니다.

    Args:
        target_date: 이동할 날짜 (YYYY-MM-DD, 기본: 2026-01-01)
    """
    client: GroupwareClient = ctx.request_context.lifespan_context["gw_client"]
    if not await client.ensure_logged_in():
        return json.dumps({"success": False, "error": "로그인 실패"}, ensure_ascii=False)

    page = client.page
    result = {}

    # ── 1. 일정 페이지 이동 ──────────────────────────────────
    await client._navigate_to_schedule()
    result["current_url"] = page.url

    # ── 2. 툴바 제목 확인 ────────────────────────────────────
    result["toolbar_title_after_nav"] = await page.evaluate(
        "() => document.querySelector('.fc-toolbar-title, [class*=\"toolbar\"] [class*=\"title\"]')?.innerText || 'NOT_FOUND'"
    )

    # ── 3. 전체 버튼 목록 ────────────────────────────────────
    result["all_buttons"] = await page.evaluate("""() =>
        [...document.querySelectorAll('button')].map(b => ({
            text: (b.innerText || b.textContent || '').trim().substring(0, 30),
            cls:  b.className.substring(0, 60)
        })).filter(b => b.text)
    """)

    # ── 4. 목록뷰 전환 시도 ──────────────────────────────────
    switched = False
    for btn_text in ["목록", "List", "list"]:
        try:
            btn = page.get_by_role("button", name=btn_text, exact=True).first
            if await btn.count() > 0:
                await btn.click()
                await page.wait_for_timeout(1500)
                switched = True
                result["list_view_switch_method"] = f"button text='{btn_text}'"
                break
        except Exception:
            pass
    if not switched:
        for sel in [".fc-calendarListMonth-button", ".fc-listMonth-button", ".fc-list-button", "[class*='listMonth']"]:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click()
                    await page.wait_for_timeout(1500)
                    switched = True
                    result["list_view_switch_method"] = f"selector '{sel}'"
                    break
            except Exception:
                pass
    result["switched_to_list"] = switched

    # ── 5. 전환 후 tr/td 구조 ────────────────────────────────
    result["dom_after_switch"] = await page.evaluate("""() => {
        const trs = [...document.querySelectorAll('tr')];
        return {
            total_tr: trs.length,
            rows: trs.slice(0, 20).map((tr, i) => ({
                row_idx: i,
                td_count: tr.querySelectorAll('td').length,
                th_count: tr.querySelectorAll('th').length,
                texts: [...tr.querySelectorAll('td, th')].map(c =>
                    (c.innerText || c.textContent || '').trim().substring(0, 40))
            }))
        };
    }""")

    # ── 6. FullCalendar API 탐색 ─────────────────────────────
    result["fc_api"] = await page.evaluate("""() => {
        const el = document.querySelector('.fc');
        if (!el) return { found: false };
        const info = { found: true, cls: el.className.substring(0, 120) };
        info['_calendar']  = typeof el._calendar;
        info['calendar']   = typeof el.calendar;
        info['getApi']     = typeof el.getApi;
        try { const a = el.getApi?.(); info['getApi_keys'] = a ? Object.keys(a).slice(0,15) : null; } catch(e) { info['getApi_err'] = e.message; }
        try { const c = el._calendar || el.calendar; info['gotoDate_ok'] = !!(c?.gotoDate); info['cal_keys'] = c ? Object.keys(c).slice(0,20) : null; } catch(e) { info['cal_err'] = e.message; }
        return info;
    }""")

    # ── 7. gotoDate 동작 테스트 ──────────────────────────────
    title_before = await page.evaluate("() => document.querySelector('.fc-toolbar-title')?.innerText || ''")
    await page.evaluate(f"""() => {{
        const el = document.querySelector('.fc');
        const cal = el && (el._calendar || el.calendar);
        if (cal?.gotoDate) cal.gotoDate('{target_date}');
    }}""")
    await page.wait_for_timeout(2000)
    title_after = await page.evaluate("() => document.querySelector('.fc-toolbar-title')?.innerText || ''")
    result["goto_date_test"] = {
        "target": target_date,
        "title_before": title_before,
        "title_after": title_after,
        "changed": title_before != title_after
    }

    # ── 8. gotoDate 후 tr/td 구조 ────────────────────────────
    result["dom_after_goto"] = await page.evaluate("""() => {
        const trs = [...document.querySelectorAll('tr')];
        return {
            total_tr: trs.length,
            rows: trs.slice(0, 40).map((tr, i) => ({
                row_idx: i,
                td_count: tr.querySelectorAll('td').length,
                th_count: tr.querySelectorAll('th').length,
                texts: [...tr.querySelectorAll('td, th')].map(c =>
                    (c.innerText || c.textContent || '').trim().substring(0, 50))
            }))
        };
    }""")

    # ── 9. .fc 컨테이너 내부 뷰 클래스 확인 ─────────────────
    result["fc_view_info"] = await page.evaluate("""() => {
        const fc = document.querySelector('.fc');
        if (!fc) return null;
        return {
            fc_class: fc.className.substring(0, 200),
            list_view_present: !!fc.querySelector('[class*="list-view"], [class*="listView"], .fc-list, .fc-calendarListMonth-view'),
            list_event_count: fc.querySelectorAll('.fc-list-event, [class*="listEvent"], [class*="list-event"]').length,
            total_tr_in_fc: fc.querySelectorAll('tr').length,
            view_containers: [...fc.querySelectorAll('[class*="fc-view"]')].slice(0, 5).map(e => ({
                tag: e.tagName, cls: e.className.substring(0, 80)
            }))
        };
    }""")

    # ── 10. 사이드바 캘린더 항목 ──────────────────────────────
    result["sidebar_calendars"] = await page.evaluate("""() => {
        const sels = [
            '.calendarList label', '.lnbCalList label',
            '[class*="calList"] label', '[class*="calGroup"] label',
            '[class*="calendarGroup"] span', '[class*="calItem"] span',
            '[class*="lnb"] span', 'aside span', 'aside label'
        ];
        const out = {};
        for (const s of sels) {
            const els = document.querySelectorAll(s);
            if (els.length > 0)
                out[s] = [...els].map(e => (e.innerText || '').trim()).filter(Boolean);
        }
        return out;
    }""")

    return json.dumps({"success": True, **result}, ensure_ascii=False, indent=2)


# ============================================================
# 서버 실행
# ============================================================

if __name__ == "__main__":
    if "--http" in sys.argv:
        mcp.run(transport="streamable-http")
    else:
        mcp.run()
