# mcp

Claude Code MCP 서버 모음. 현재는 더존 아마란스 ERP 그룹웨어 자동화용 `erp_groupware` MCP 서버를 포함.

## 구성

### erp_groupware
더존 아마란스 기반 그룹웨어의 메일·일정을 Playwright로 자동화하는 MCP 서버.

주요 도구:
- `gw_check_login` / `gw_reset_browser` — 로그인 상태 확인, 브라우저 세션 리셋
- `gw_list_inbox` / `gw_read_mail` / `gw_search_mail` / `gw_send_mail` — 메일
- `gw_list_schedule` / `gw_read_schedule` / `gw_create_schedule` / `gw_update_schedule` / `gw_delete_schedule` — 일정
- `gw_list_recent_files` — 첨부 후보 파일 조회

## 설정

1. 레포 루트의 `mcp.json.example`을 `mcp.json`으로 복사.
2. `ERP_BASE_URL`, `ERP_USERNAME`, `ERP_PASSWORD`를 실제 값으로 채움.
3. `mcp.json`은 `.gitignore`에 의해 커밋되지 않음 (비밀번호 평문 포함).

```bash
cp mcp.json.example mcp.json
# mcp.json 편집 — 실제 계정 정보 입력
```

## 의존성

- Python 3.12+ (Windows에서는 Python 3.14 경로를 기본으로 참조)
- `fastmcp`, `playwright`, `python-dotenv`

```bash
cd erp_groupware
pip install -r requirements.txt
playwright install chromium
```

## 알려진 이슈

`gw_create_schedule`의 날짜/시간 쓰기 경로 — OBTDatePickerRebuild (Vue 기반) 컴포넌트 상태 갱신 실패. 하드-페일 검증이 포함되어 잘못된 날짜로 저장되지 않지만, 오늘 외 날짜 등록은 현재 차단됨. 해결 방향은 캘린더 팝업 클릭 방식으로 전환 검토 중.

## 라이선스

Private / 내부용.
