# OIKOS Thesis Formatter

OIKOS(오이코스) 신학대학원 박사학위논문을 규정 서식에 맞게 자동 편집하는 Python 도구입니다.

## 주요 기능

- **쪽번호 관리** — 섹션별 쪽번호 형식(로마숫자/아라비아숫자) 및 시작값 설정
- **문서 구조 분석** — Word 문서의 섹션, 마진, 헤더/푸터 구조를 시각적으로 분석
- **섹션 관리** — 섹션 순서 변경 및 구역 나누기 제어
- **스타일 관리** — 논문 규정에 맞는 폰트, 크기, 간격 등 스타일 적용
- **머리글/바닥글 편집** — 양면 인쇄용 짝/홀수 페이지 헤더·푸터 설정

## 프로젝트 구조

```
OIKOS-Thesis-Formatter/
├── OIKOS 논문 템플릿(샘플).docx    ← 서식 기준 템플릿
├── OIKOS-template-spec.md         ← 템플릿 사양서
├── tools/                         ← 핵심 편집 도구
│   ├── analyzer.py                  문서 구조 분석
│   ├── page_number_manager.py       쪽번호 제어
│   ├── section_manager.py           섹션 순서/구역 나누기
│   ├── style_manager.py             스타일 관리
│   ├── header_footer_manager.py     머리글/바닥글 편집
│   └── page_mapper.py               페이지 매핑
├── scripts/                       ← 실행 스크립트
├── prompt/                        ← Claude Code 연동 프롬프트
└── backups/                       ← 자동 백업 저장소
```

## 요구 사항

- Python 3.10+
- [python-docx](https://python-docx.readthedocs.io/)

```bash
pip install python-docx
```

## 사용 방법

### 1. 문서 구조 확인

```python
from tools.analyzer import load_document, get_section_info

doc = load_document("논문파일.docx")
sections = get_section_info(doc)
for s in sections:
    print(s)
```

### 2. 쪽번호 설정 변경

```python
from tools.page_number_manager import get_page_number_settings, set_page_number

doc = load_document("논문파일.docx")

# 현재 설정 확인
settings = get_page_number_settings(doc)

# 특정 섹션 쪽번호 형식 변경 (예: 로마숫자 → 아라비아숫자)
set_page_number(doc, section_index=5, fmt='decimal', start=1)
doc.save("수정된논문.docx")
```

### 3. Claude Code에서 사용

이 프로젝트는 [Claude Code](https://claude.ai/claude-code)와 함께 사용하도록 설계되었습니다. `CLAUDE.md`에 정의된 지침에 따라 Claude가 `tools/` 모듈을 직접 호출하여 논문을 편집합니다.

## 안전 규칙

- 원본 논문 파일은 **절대 수정하지 않음** — 항상 작업 복사본에서 편집
- 모든 편집 전 `backups/`에 타임스탬프 백업 자동 생성
- 편집 후 문서 무결성 검증 수행

## 라이선스

MIT License
