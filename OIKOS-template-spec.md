# OIKOS 논문 템플릿 사양서 (Template Specification)

> 기준 파일: `OIKOS 논문 템플릿(샘플).docx`
> 분석일: 2026-03-31
> 목적: 어떤 논문이든 이 사양에 맞춰 OIKOS 양면/단면 형식으로 자동 변환

---

## 1. 문서 기본 정보

| 항목 | 값 |
|---|---|
| 총 섹션 | 47 |
| 총 단락 | 1,174 |
| 페이지 크기 | 21.6 × 27.9 cm (Letter) |
| 페이지 방향 | 세로 (portrait) — 전 섹션 동일 |
| 단 설정 | 1단 (전 섹션 동일) |
| mirrorMargins | 없음 |
| evenAndOddHeaders | 있음 (양면 인쇄용 짝/홀수 구분) |
| 기본 탭 간격 | 240 twips (0.42cm) |
| 호환성 모드 | 15 (Word 2013+) |
| 자동 하이픈 | 사용, 연속 제한 3회 |

---

## 2. 테마 폰트

| 구분 | Latin | Korean (Hang) |
|---|---|---|
| Major (제목) | 맑은 고딕 | 맑은 고딕 |
| Minor (본문) | 맑은 고딕 | 맑은 고딕 |

> 테마 폰트는 기본값이며, 실제 스타일에서 개별 폰트로 오버라이드됨

---

## 3. 마진 패턴 (2가지)

### 패턴 A — 장/섹션 시작 페이지 (nextPage break)

| 항목 | 값 |
|---|---|
| Top | 5.08 cm |
| Bottom | **2.97 cm** (전면부/본문 통일) |
| Left | 3.81 cm |
| Right | 2.54 cm |
| Header | 1.27 cm |
| Footer | **2.23 cm** (전면부/본문 통일) |
| Gutter | 0 |

### 패턴 B — 연속 페이지 (continuous break)

| 항목 | 값 |
|---|---|
| Top | 2.54 cm |
| Bottom | **2.97 cm** (전면부/본문 통일) |
| Left | 3.81 cm |
| Right | 2.54 cm |
| Header | 2.23 cm (일부 1.91cm, 2.03cm) |
| Footer | **2.23 cm** (전면부 통일) |
| Gutter | 0 |

---

## 4. 스타일 상세 (사용된 16개 스타일)

### 4.1 Normal (기본 스타일) — 46건

| 항목 | 값 |
|---|---|
| 폰트 크기 | 12pt (sz=24 half-pts) |
| 줄간격 | 240 twips, lineRule=atLeast |
| autoSpaceDE | 0 (한영 자동간격 꺼짐) |
| autoSpaceDN | 0 (한숫자 자동간격 꺼짐) |
| 언어 | eastAsia=en-US |

### 4.2 *¶ Indented (본문) — 734건 (가장 많이 사용)

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 정렬 | both (양쪽 정렬) |
| 첫줄 들여쓰기 | 720 twips (1.27cm) |
| 줄간격 | 480 twips, lineRule=atLeast (더블스페이스) |

### 4.3 *¶ Continued (no indent) — 1건

| 항목 | 값 |
|---|---|
| 기반 | *¶ Indented |
| 첫줄 들여쓰기 | 0 (들여쓰기 없음) |
| 나머지 | *¶ Indented와 동일 |

### 4.4 *Heading: Main (장 제목) — 33건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | *¶ Indented |
| 볼드 | true |
| 대문자 | true (caps) |
| 정렬 | center (가운데) |
| 좌/우 들여쓰기 | 1080 twips (1.91cm) 양쪽 |
| 단락 뒤 간격 | 480 twips |
| keepLines | true (단락 내 분리 금지) |
| outlineLvl | 1 |

### 4.5 *Heading: 1st-Level (절 제목) — 9건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | *¶ Indented |
| 볼드 | true |
| szCs | 12pt |
| 글자 색상 | 000000 (검정) |
| 언어 | eastAsia=ko-KR |
| 정렬 | center (가운데) |
| 좌 들여쓰기 | 1627 twips (2.87cm) |
| 우 들여쓰기 | 1077 twips (1.90cm) |
| 단락 앞 간격 | 720 twips |
| keepNext | true |
| outlineLvl | 2 |

### 4.6 Heading 3 (소제목) — 49건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | Normal |
| 볼드 | true |
| 정렬 | center (가운데) |
| 좌/우 들여쓰기 | 1267 twips (2.23cm) 양쪽 |
| 단락 앞 간격 | 720 twips |
| keepNext | true |
| outlineLvl | 2 |

### 4.7 Heading 4 (하위 소제목) — 42건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | Normal |
| 볼드 | true |
| 이탤릭 | true |
| 단락 앞 간격 | 720 twips |
| keepNext | true |
| outlineLvl | 3 |

### 4.8 인용1 (블록 인용) — 20건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 한글 폰트 | 바탕체 (eastAsia=바탕체) |
| 정렬 | both (양쪽 정렬) |
| 좌/우 들여쓰기 | 720 twips (1.27cm) 양쪽 |
| 줄간격 | 240 twips, lineRule=auto (싱글스페이스) |
| 탭 | left@1440 |
| 언어 | eastAsia=zh-CN |

### 4.9 EndNote Bibliography (참고문헌) — 84건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| (Normal 서식 그대로 상속) | |

### 4.10 *Abbreviations (약어 목록) — 3건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | Normal |
| 좌 들여쓰기 | 1440 twips (2.54cm) |
| 내어쓰기 | 1440 twips (2.54cm) |
| 단락 앞 간격 | 240 twips |

### 4.11 *Biblio: Author (참고문헌 저자) — 2건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | *Biblio: Entry |
| 단락 앞 간격 | 240 twips |
| keepNext | true |

### 4.12 *Biblio: Entry (참고문헌 항목) — 2건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 좌 들여쓰기 | 1426 twips (2.52cm) |
| 내어쓰기 | 994 twips (1.75cm) |
| 단락 뒤 간격 | 240 twips |
| 탭 | left@270 |

### 4.13 toc 2 (목차 2레벨) — 37건

| 항목 | 값 |
|---|---|
| 기반 | toc 1 |
| 단락 앞 간격 | 240 twips |

### 4.14 toc 3 (목차 3레벨) — 58건

| 항목 | 값 |
|---|---|
| 기반 | toc 2 |
| 다음 스타일 | Normal |
| 좌 들여쓰기 | 1440 twips (2.54cm) |
| 단락 앞 간격 | 0 |

### 4.15 toc 4 (목차 4레벨) — 42건

| 항목 | 값 |
|---|---|
| 기반 | toc 3 |
| 다음 스타일 | Normal |
| 좌 들여쓰기 | 2160 twips (3.81cm) |

### 4.16 table of figures (표/그림 목차) — 12건

| 항목 | 값 |
|---|---|
| 기반 | Normal |
| 다음 스타일 | Normal |
| 폰트 | ascii=Times, hAnsi=Times |
| 대문자 | true (caps) |
| 좌 들여쓰기 | 1440 twips (2.54cm) |
| 우 들여쓰기 | 720 twips (1.27cm) |
| 내어쓰기 | 1440 twips (2.54cm) |
| 단락 앞 간격 | 240 twips |
| 탭 | left@1368, left@1440, right@8640(leader=dot) |

---

## 5. 문서 구조 (47 섹션 매핑)

### 5.1 전면부 (S0~S17) — lowerRoman

| 섹션 | break | 내용 | titlePg | 마진 패턴 |
|---|---|---|---|---|
| S0 | nextPage | 표지 | no | A |
| S1 | continuous | (빈 페이지) | no | A |
| S2 | continuous | ABSTRACT (영문) | YES | A (F=2.23) |
| S3 | continuous | (영문요약 계속) | YES | B |
| S4 | nextPage | 개요 (국문 ABSTRACT) | YES | A |
| S5 | continuous | (국문요약 계속) | YES | B |
| S6 | nextPage | 목차 (TABLE OF CONTENTS) | YES | A |
| S7 | continuous | (목차 계속) | YES | B (H=2.23, F=2.23) |
| S8 | continuous | (빈) | YES | B |
| S9 | nextPage | 표 목록 (LIST OF TABLES) | YES | A |
| S10 | continuous | (그림 목록) | YES | B |
| S11 | continuous | (빈) | YES | B |
| S12 | nextPage | 약어 목록 | YES | A |
| S13 | continuous | (빈) | YES | B |
| S14 | continuous | (빈) | YES | B |
| S15 | nextPage | 헌정 (DEDICATION) | YES | A |
| S16 | nextPage | 감사의 말 | YES | A |
| S17 | continuous | (감사의 말 계속) | YES | B |

- 쪽번호: lowerRoman, ABSTRACT(S2)에서 start=1 (= i)
- 표지(S0~S1): 쪽번호 없음
- S2~S17: 하단 중앙 PAGE 필드 (style=a6, jc=center, PAGE field — 본문 장시작과 동일 방식/위치)
- even_page_footer: S2에서 설정, S3~S17은 linked로 상속

### 5.2 본문 (S18~S28) — decimal

| 섹션 | break | 내용 | titlePg |
|---|---|---|---|
| S18 | nextPage | 제1장 서론 | YES |
| S19 | continuous | (1장 계속) | YES |
| S20 | nextPage | 제2장 선행연구고찰 | YES |
| S21 | continuous | (2장 계속) | YES |
| S22 | nextPage | 제3장 연구 방법론 | YES |
| S23 | continuous | (3장 계속) | YES |
| S24 | nextPage | 제4장 사례 분석 | YES |
| S25 | continuous | (4장 계속) | YES |
| S26 | continuous | 제5장 논의 | YES |
| S27 | nextPage | 제6장 결론 | YES |
| S28 | nextPage | 제7장 제언 | YES |

- S18에서 start=1 (decimal)
- 장 시작 페이지: 하단 중앙 (first_page_footer)
- 일반 홀수 페이지: 상단 오른쪽 (default_header, jc=right)
- 일반 짝수 페이지: 상단 왼쪽 (even_header)
- 장 시작 페이지 마진: 패턴 A (top=5.08cm)
- 연속 페이지 마진: 패턴 B (top=2.54cm)

### 5.3 영문 요약 (S29~S44) — decimal 연속

| 섹션 | break | 내용 |
|---|---|---|
| S29 | nextPage | ENGLISH SUMMARY |
| S30 | continuous | (계속) |
| S31 | nextPage | TABLE OF CONTENTS |
| S32 | continuous | (계속) |
| S33 | nextPage | TEXT OF SUMMARY |
| S34 | continuous | CHAPTER 1 |
| S35 | nextPage | CHAPTER 2 |
| S36 | continuous | (계속) |
| S37 | nextPage | CHAPTER 3 |
| S38 | continuous | (계속) |
| S39 | nextPage | CHAPTER 4 |
| S40 | continuous | (계속) |
| S41 | nextPage | CHAPTER 5 |
| S42 | continuous | (계속) |
| S43 | nextPage | CHAPTER 6 |
| S44 | continuous | (계속) |

- 쪽번호: **하단 중앙만** (상단 없음)

### 5.4 후면부 (S45~S46) — decimal 연속

| 섹션 | break | 내용 |
|---|---|---|
| S45 | nextPage | 인용 문헌 (REFERENCES CITED) |
| S46 | continuous | 부록 (Appendices) |

- 쪽번호: 하단 중앙

---

## 6. 쪽번호 규칙 종합

| 영역 | 형식 | 시작값 | 위치 | 방식 |
|---|---|---|---|---|
| 표지 (S0~S1) | - | - | **없음** | - |
| 전면부 (S2~S17) | lowerRoman | i (start=1) | **하단 중앙** | PAGE field (style=a6, jc=center) |
| 본문 장 시작 | decimal | 1 (S18 start=1) | **하단 중앙** | PAGE field (style=a6, jc=center) |
| 본문 일반 홀수 | decimal | 연속 | **상단 오른쪽** | PAGE field (jc=right) |
| 본문 일반 짝수 | decimal | 연속 | **상단 왼쪽** | PAGE field |
| 영문 요약~부록 (S29~S46) | decimal | 연속 | **하단 중앙** | PAGE field (style=a6, jc=center) |

> **통일 원칙**: 하단 중앙 쪽번호는 전면부/본문/영문부 모두 동일한 footer 거리(2.23cm), bottom 마진(2.97cm), 스타일(a6), 정렬(center), 방식(PAGE field)을 사용한다.

---

## 7. 각주/미주 설정

| 항목 | 값 |
|---|---|
| 각주 수 | 2개 (separator/continuation 포함) |
| 번호 형식 | (기본값 — decimal) |
| 미주 | 설정 없음 |

---

## 8. 목록/번호 매기기

- abstractNum 정의: 18개
- 주요 번호 형식:
  - `%1.` (decimal, 1단계)
  - `%1.%2.` (decimal, 2단계)
  - `%1.%2.%3.` (decimal, 3단계)
- 불릿: 2단계에서 bullet 사용

---

## 9. 표 서식

- 표 수: 6개
- 표 정렬: center
- 표 너비: auto
- 테두리 패턴:
  - 학술 표: top/bottom만 single (좌우/내부 none)
  - 데이터 표: 전체 single border

---

## 10. 이미지/도형

- drawing 요소: 9개
- (이미지 배치 방식은 개별 분석 필요)

---

## 11. 문서 격자

| 섹션 | 격자 설정 |
|---|---|
| S0~S25 | 격자 없음 (기본값) |
| S26~S28 | linePitch=326, charSpace 있음 |
| S29~S46 | 격자 없음 |

---

## 12. 하이퍼링크 스타일

| 항목 | 값 |
|---|---|
| 스타일 ID | ab |
| 색상 | 000000 (검정) |
| 밑줄 | none |
| (목차 하이퍼링크가 검정, 밑줄 없음) | |

---

## 13. 스타일 상속 체인

```
Normal (기본)
├── *¶ Indented (본문)
│   └── *¶ Continued (no indent) (들여쓰기 없는 본문)
├── *Heading: Main (장 제목) → next: *¶ Indented
├── *Heading: 1st-Level (절 제목) → next: *¶ Indented
├── Heading 3 (소제목) → next: Normal
├── Heading 4 (하위 소제목) → next: Normal
├── 인용1 (블록 인용)
├── *Abbreviations (약어) → next: Normal
├── *Biblio: Author (저자) → next: *Biblio: Entry
├── *Biblio: Entry (참고문헌 항목)
├── EndNote Bibliography (참고문헌)
├── table of figures (표/그림 목차) → next: Normal
└── toc 1
    └── toc 2
        └── toc 3 → next: Normal
            └── toc 4 → next: Normal
```

---

## 14. 장 시작 페이지 vs 연속 페이지 마진 차이

장(Chapter) 시작 시 상단 여백이 넓고(5.08cm), 연속 페이지는 좁음(2.54cm).
이는 장 제목이 페이지 상단에서 아래로 내려와 시작하는 학술 논문 관행을 반영.

| 항목 | 장 시작 (패턴 A) | 연속 (패턴 B) |
|---|---|---|
| Top margin | 5.08 cm | 2.54 cm |
| Header distance | 1.27 cm | 2.23 cm |
| Footer distance | 2.23 cm | 1.27 cm |

---

## 15. 특이사항 및 주의점

1. **S13~S14**: 약어 목록 뒤 빈 섹션. fmt=lowerRoman으로 정리됨
2. **S26 (제5장)**: break=continuous (nextPage 아님) — 특수 처리
3. **S31, S33**: titlePg=no — 영문 목차/요약 시작에서 다른 패턴
4. **인용1 스타일**: 한글 폰트가 바탕체, 언어가 zh-CN으로 설정 (특이)
5. **Normal 스타일**: eastAsia 언어가 en-US로 설정
6. **autoSpaceDE/DN**: 꺼져 있음 (한글-영문/숫자 자동 간격 비활성)

---

## 16. 수정 이력

| 날짜 | 내용 |
|---|---|
| 2026-03-31 | 전면부(S2~S17) footer 통일: bottom=2.97cm, footer거리=2.23cm, PAGE field 방식으로 변경. 본문 장시작과 동일 위치/방식 |
| 2026-03-31 | 전면부 even_page_footer: S2에 PAGE(center) 설정, S3~S17 linked 상속 |
| 2026-03-31 | S0~S1 표지: framePr/PAGE 제거, 쪽번호 없음 처리 |
| 2026-03-31 | 하이퍼링크 스타일: 파란색→검정, 밑줄 제거 |
| 2026-03-31 | 목차에서 논문 제목(영문/한글) 2줄 삭제 |
