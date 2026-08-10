# 수술통계 대시보드 (2017~)

iCloud 의 수술 엑셀 DB를 읽어 **비식별 집계값만** 뽑아, GitHub Pages 로 배포하는 정적 대시보드입니다.

## 구조

```
build.py       엑셀 → 정규화 → docs/data.js + docs/data.json  (집계값만)
mapping.yaml   분류 매핑표 — 분류를 바꾸려면 이 파일만 수정
update.sh      빌드 → 커밋 → push 한 번에
docs/          GitHub Pages 로 배포되는 폴더
  index.html   대시보드 (단일 파일, 외부 라이브러리 없음)
  data.js      집계 데이터
```

## 개인정보

`docs/` 에 올라가는 데이터에는 다음이 **포함되지 않습니다**.

- 환자 성명, 등록번호
- 수술 날짜 (연·월까지만 집계에 사용)
- 나이 (10세 단위 구간으로만 저장)
- Remarks, OP findings 등 자유기술 내용

원본 엑셀은 iCloud 에만 있으며 이 저장소에 복사되지 않습니다.

## 환자를 추가한 뒤

엑셀에 행을 추가하고 저장하면 자동 감시가 갱신합니다. 즉시 반영하려면:

```bash
~/Projects/surgery-dashboard/update.sh
```

## 분류를 바꾸려면

`mapping.yaml` 을 열어 해당 항목을 고치고 `update.sh` 를 실행합니다.
매핑에 없는 새 값이 들어오면 `미분류` 로 집계되고 `update.log` 에 경고가 남습니다.

주요 설정:

| 항목 | 위치 |
|---|---|
| 원본 엑셀 경로 | `source_file` |
| 동맥류 집계 기준 (현재 809건) | `aneurysm` |
| 질환군 분류 | `disease_group` |
| 세부 술기 이름 | `procedure` |
| 데이터 공백 기간 | `gap_periods` |

## 집계 기준 메모

- **총 1,944건** — 시술자에 C 가 없는 45건(2018~2023, OSW/LJK/CWC 등)도 팀 실적으로 포함
- **동맥류 809건** — `Dx = UIA 또는 SAH` 이면서 clip(202) / coil(602) / 시도 후 조영술만 종료(5)
  - `Op name = 1(coil)` 675건 중 74건은 DAVF·CCF·AVF·종양이므로 동맥류에서 제외됨
  - 질환군 `뇌동맥류`(진단 기준, 883건)와 동맥류 치료 건수(809건)는 다른 숫자입니다
- **2024-07 ~ 2025-08** — 수술 미시행 기간. 해당 연도 건수는 다른 연도와 직접 비교하지 말 것

## 로컬에서 보기

```bash
python3 -m http.server 8777 --directory ~/Projects/surgery-dashboard/docs
```
