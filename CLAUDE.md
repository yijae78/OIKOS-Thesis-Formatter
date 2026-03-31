# CLAUDE.md — Thesis Editor (Claude Code Instructions)

## Project
**Thesis-Editor-AgenticWorkflow**: Word 논문 편집 시스템
- 박사학위논문의 쪽번호, 섹션 순서, 스타일, 머리글/바닥글을 편집
- python-docx 기반 Python 도구로 조작
- 원본 파일은 절대 수정하지 않음 — 항상 작업 복사본 사용

## Hub Reference
- `AGENTS.md` — 모든 에이전트 공통 규칙 (반드시 준수)
- `soul.md` — 시스템 철학 및 가치관
- `workflow.md` — 편집 워크플로우 정의

## Absolute Criteria
1. **문서 무결성**: 편집 후 .docx가 Word에서 정상적으로 열려야 함
2. **SOT**: `state.yaml`에 모든 세션 상태 기록
3. **안전**: 모든 편집 전 자동 백업

## Project Structure
```
논문수정/
├── CLAUDE.md                         ← You are here
├── AGENTS.md                         ← Universal rules
├── soul.md                           ← Philosophy & values
├── workflow.md                       ← Editing workflow
├── state.yaml                        ← Session SOT
├── 1. 신이재 박사학위논문 원본.docx     ← ORIGINAL (NEVER MODIFY)
├── tools/
│   ├── __init__.py                   ← Package init + shared utilities
│   ├── analyzer.py                   ← Document structure analysis
│   ├── section_manager.py            ← Section reordering & breaks
│   ├── page_number_manager.py        ← Page number control
│   ├── style_manager.py              ← Style management
│   └── header_footer_manager.py      ← Header/footer editing
├── scripts/
│   ├── editor.py                     ← Main interactive editor
│   └── backup_restore.py             ← Backup management
├── backups/                          ← Timestamped backups
└── prompt/
    └── editing-commands.md           ← Command reference
```

## How to Use

### Quick Start
```bash
# 1. 문서 구조 확인
python scripts/editor.py --show-structure

# 2. 특정 섹션 쪽번호 변경
python scripts/editor.py --page-numbers

# 3. 섹션 순서 변경
python scripts/editor.py --reorder-sections

# 4. 스타일 확인/변경
python scripts/editor.py --styles
```

### Claude Code에서 사용
사용자가 요청하면:
1. `tools/` 모듈을 import하여 Python으로 직접 조작
2. 항상 `backup_restore.py`로 먼저 백업
3. 변경 후 `analyzer.py`로 결과 검증
4. `state.yaml` 업데이트

## Safety Rules
- 원본 파일(`1. 신이재 박사학위논문 원본.docx`)은 **절대** 수정 금지
- 모든 작업은 `working_copy.docx`에서 수행
- 편집 전 반드시 `backups/`에 타임스탬프 백업 생성
- 편집 후 반드시 L0 무결성 검증 수행
- 실패 시 최근 백업에서 자동 복원

## Language
- 사용자 소통: 한국어
- 코드/변수/주석: English
- 파일명: 한국어 허용 (원본 파일명 유지)
