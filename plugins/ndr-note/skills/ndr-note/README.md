# ndr-note 스킬 — 다른 PC로 이식하기

기업 탐방(NDR/IR) 노트를 Nexus 양식 Word/PDF 보고서로 자동 작성하는 Claude Code 스킬.
**이 폴더만 복사하면 어느 PC에서도 동작**하도록 구성되어 있습니다(절대경로 의존 없음).

## 1. 폴더 복사
이 `ndr-note/` 폴더 전체를 새 PC의 다음 위치에 둡니다.
- Windows: `C:\Users\<사용자>\.claude\skills\ndr-note\`
- macOS / Linux: `~/.claude/skills/ndr-note/`

복사 수단은 무엇이든 OK (USB·클라우드·git). 여러 PC를 쓰면 `.claude/skills/`를 git으로 관리하면 동기화가 편합니다.

## 2. 의존성 설치
| 항목 | 용도 | 설치/조건 |
|---|---|---|
| Python 3.x + python-docx | docx 생성 | `pip install -r scripts/requirements.txt` |
| LibreOffice **또는** MS Word | docx→PDF | 둘 중 하나. LibreOffice면 전 OS 가능, Windows면 Word도 가능 |
| NotebookLM MCP (`nlm`) | 데이터 소스 | 새 PC에서 `nlm login` 재인증 + Claude Code에 notebooklm MCP 등록 |

> PDF는 `scripts/convert_to_pdf.py`가 LibreOffice(soffice) → Windows+Word 순으로 자동 시도합니다.
> 둘 다 없으면 docx만 생성됩니다(PDF 단계 생략).

## 3. 동작 확인 (선택)
```
cd .claude/skills/ndr-note
pip install -r scripts/requirements.txt
python scripts/build_ndr_note.py scripts/config_amotech.json   # 바탕화면에 docx 생성
python scripts/convert_to_pdf.py "<위에서 출력된 .docx 경로>"   # PDF 변환
```
아모텍 예시 보고서가 바탕화면에 생성되면 정상입니다.

## 4. 사용
Claude Code에서 `/ndr-note` 또는 "○○ 탐방노트 만들어줘"로 호출. 워크플로우는 `SKILL.md` 참고.

## 머신별로 바꿀 수 있는 값 (config에서)
- `meta.firm` (회사명/브랜드), `reported_by`(이름·이메일)
- `output`은 파일명만 적으면 그 PC 바탕화면에 저장됩니다(절대경로도 가능).
- `reference/`의 양식 파일은 디자인 근거용 참고 자료이며 생성에는 필요 없습니다(디자인은 제너레이터에 내장).
