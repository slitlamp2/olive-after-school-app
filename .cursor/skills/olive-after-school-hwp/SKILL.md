---
name: olive-after-school-hwp
description: "방과후센터 앱에서 일일활동일지·품위서·계획서 HWP(.hwp) 문서를 HwpController/hwp_service로 생성·편집할 때 사용하는 스킬. .hwp 템플릿 구조를 유지한 채 내용만 채우거나 수정할 때 사용."
---

# 올리브 방과후센터 HWP 문서 스킬 (.hwp 전용)

이 스킬은 **`C:\Users\윤미란\Desktop\방과후센터` 프로젝트**에서,
`일일활동일지`, `품위서`, `계획서` 한글 **`.hwp` 템플릿**을
이미 구현된 `hwp_service.py` / HwpController를 통해 채우거나 수정할 때의
사용 규칙을 정의한다.

## 1. 관련 파일

- `app/services/hwp_service.py`
  - `create_daily_log(data, output_dir)`
  - `create_purchase_doc(data, output_dir, image_paths=None)`
  - `create_plan(data, output_dir, image_paths=None)`
- `app/routes/daily_log.py`
- `app/routes/purchase_doc.py`
- `app/routes/plan.py`
- 템플릿 경로 (CONTEXT.md 기준)
  - `D:\일일활동일지양식.hwp`
  - `D:\품위서양식.hwp`
  - `D:\계획서양식.hwp`

## 2. 공통 원칙 (.hwp 작업 시)

1. **템플릿은 그대로 두고, 내용만 채운다.**
   - 표 구조, 셀 병합, 글꼴, 여백, 쪽수는 템플릿에서 제어한다.
   - 코드는 `_fill_label`, `_fill_purchase_table` 등 **텍스트/셀 내용**만 바꾼다.
2. **항상 `hwp_service.py`의 public 함수만 통해서 HWP를 생성한다.**
   - 직접 COM을 새로 여는 코드(HwpController 인스턴스 생성 등)를 여기저기 추가하지 않는다.
   - 필요한 동작이 없으면 `hwp_service.py` 안에 **도우미 함수를 추가**한 뒤,
     public 함수에서 호출하도록 구조를 유지한다.
3. **경로와 템플릿은 하드코딩된 값(위 3개 .hwp) 기준으로 맞춘다.**
   - 템플릿 경로를 바꿔야 하면 `CONTEXT.md`와 `hwp_service.py`를 함께 수정한다.
4. **에러 발생 시 사용자에게는 Flask 라우트에서 메시지를 보여주고,
   HWP 내부에서는 팝업이 뜨지 않도록 `_controller_session()` 컨텍스트를 사용한다.**

## 3. 품위서 (`create_purchase_doc`) 사용 가이드

### 3.1. 입력 데이터 구조

`create_purchase_doc(data, output_dir, image_paths=None)`의 `data` 딕셔너리는 다음 키를 가진다:

- `purchase_date` (str): 구입일자 (`YYYY.MM.DD` 또는 `YYYY-MM-DD`)
- `store_name` (str): 매장명 (현재는 템플릿 안 특정 셀에 직접 매핑하지 않지만, 필요 시 `_fill_label`로 확장)
- `items` (list[dict]):
  - 각 원소: `{ "name", "qty", "unit", "unit_price", "amount", "note" }`
- `total_amount` (str): 합계 금액

`image_paths`:
- 영수증 이미지 파일 경로 리스트
- 2페이지 “영수증 사진” 표에 한 장만 대표로 삽입

### 3.2. 동작 요약

- 템플릿: `D:\품위서양식.hwp`
- 동작 순서:
  1. `_replace_first_date_in_doc`로 문서 내 첫 `YYYY.MM.DD` 날짜를 `purchase_date`로 치환
  2. `_fill_label("구입일자", purchase_date)`로 레이블 기반 날짜 셀 채우기
  3. `_fill_purchase_table(items)`로 **표의 "품명" 헤더 아래 행을 전부 채움**
  4. `_fill_label("합  계", total_amount)` 및 동일 레이블 두 번째 occurrence도 채움
  5. `_insert_receipt_images(image_paths)`로 2페이지 마지막 표에 영수증 이미지 삽입

### 3.3. 수정 시 지켜야 할 점

- 표 구조를 바꾸지 말고, `_fill_purchase_table` 내부에서
  **각 열에 들어갈 값 매핑만 수정**한다.
- 템플릿에서 헤더 이름(예: `"품명"`)이 바뀌면,
  `_fill_purchase_table`의 `ctrl.find_text("품명")` 부분만 템플릿에 맞게 조정한다.

## 4. 계획서 (`create_plan`) 사용 가이드

### 4.1. 입력 데이터 구조

`create_plan(data, output_dir, image_paths=None)`의 `data`는 현재 다음 키를 사용한다:

- `plan_date` (str): 작성일자
- `purpose` (str): 목적
- `goal` (str): 목표
- `program_content` (str): 프로그램 내용
- `expected_effect` (str): 기대 효과
- `purchase_summary` (str): 품위서 요약 텍스트 (화면 표시/선택 칸에 사용)
- `purchase_items` (list[dict]): 품위서와 동일한 품목 리스트
- `purchase_total_amount` (str): 품위서 합계 금액

`image_paths`:
- 계획서 2페이지 “사진1, 사진2, …” 셀에 넣을 사진 경로 리스트

### 4.2. 동작 요약

- 템플릿: `D:\계획서양식.hwp`
- 동작 순서:
  1. `_fill_label("작성일자", plan_date)`
  2. `_fill_label("목적", purpose)`
  3. `_fill_label("목표", goal)`
  4. `_fill_label("프로그램 내용", program_content)`
  5. `_fill_label("기대효과", expected_effect)`
  6. `purchase_items`가 있으면 `_fill_purchase_table(purchase_items)`로
     **계획서 구입내역 표(헤더 '품명')도 품위서와 동일하게 채움**
  7. `purchase_total_amount`가 있으면 `"합계"`, `"합  계"` 레이블 옆 칸에 합계 입력
  8. (선택) `purchase_summary`가 있으면 `"품위서 내용"`, `"구입내역"` 등
     템플릿 내 요약용 셀 레이블에 줄바꿈 유지(`\n`→`\r`) 상태로 삽입
  9. `_insert_plan_photos(image_paths)`로 2페이지 표의 `"사진1"`, `"사진2"` 셀에 사진 삽입

### 4.3. 품위서 → 계획서 연동 원칙

- `routes/purchase_doc.py`에서 품위서 생성 시:
  - `session['purchase_items']` ← 영수증 분석 결과 `items`
  - `session['purchase_total_amount']` ← `total_amount`
  - `session['purchase_summary']` ← 화면용 요약 텍스트
- `routes/plan.py`에서 계획서 생성 시:
  - 위 세션 값을 읽어 `data` 딕셔너리의 `purchase_items`, `purchase_total_amount`,
    `purchase_summary` 세 필드에 그대로 전달한다.
- **스킬 사용 시**:
  - 품위서 → 계획서 플로우를 깨지 말고,
  - 계획서에서 품위서 표를 새로 해석하지 말고
    **이미 정제된 `items`/`total_amount`를 재사용**한다.

## 5. 일일활동일지 (`create_daily_log`) 사용 가이드

### 5.1. 입력 데이터 구조

`create_daily_log(data, output_dir)`의 `data`는 다음 키를 사용한다:

- `date` (str): 활동 날짜
- `time` (str): 이용 시간 (예: `"14:30~18:30"`)
- `place` (str): 장소 (필요 시)
- `student_names` (str): 활동 명단 (이름 문자열)
- `activities` (dict): 시간대별 활동 내용
  - `activity_1430`, `activity_1500`, `activity_1600`, `activity_1700`, `activity_1800`
  - `special_note`
- `activity_content` (str): 브라우저에서 보여줄 전체 활동 텍스트 (요약)
- `photo_paths` (list): 사진 경로 리스트

### 5.2. 동작 요약

- 템플릿: `D:\일일활동일지양식.hwp`
- 동작 순서:
  1. `_build_activity_text(activities)`로 시간대별 내용을 하나의 텍스트 블록으로 합침
     - 시간대 헤더 제거
     - “없음/특이사항 없음” 등은 제외
  2. `_fill_label("일시", f"{date}  {time}")`
  3. `_fill_label("이용시간", time)`
  4. `_fill_label("이용자", student_names)`
  5. “참여 내용” 헤더를 찾아, 바로 아래 **프로그램 참여 내용 셀**에
     `full_activity`를 줄바꿈 유지(`\n`→`\r`) 상태로 삽입
  6. (현재) 사진 자동 삽입은 비활성화:
     - `_spawn_daily_log_photo_worker`를 통한 비동기 처리만 준비되어 있으며,
       `ENABLE_DAILY_LOG_PHOTO_INSERTION`가 `True`일 때만 사용한다.

### 5.3. 편집 시 주의사항

- 활동 텍스트 규칙(“없음” 제거, 따뜻한 말투 등)은 `CONTEXT.md`에서 정한 요구사항을 따른다.
- 셀 위치를 직접 좌표로 조작하지 말고, 반드시
  - 레이블 찾기(`ctrl.find_text("참여 내용")`)
  - `TableSelCell` / `InsertText` 조합을 써서 템플릿 구조를 보존한다.

## 6. 새로운 HWP 요구사항이 생겼을 때

새로운 `.hwp` 문서를 다뤄야 할 때는 다음 순서를 따른다.

1. **한글에서 템플릿을 먼저 만든다.**
   - 표/머리글/바닥글/쪽수를 모두 템플릿에서 디자인한다.
2. `hwp_service.py`에 **전용 생성 함수**를 추가한다.
   - `_controller_session()`을 사용해 HWP를 열고,
   - `_fill_label` / 전용 테이블 채우기 함수로 내용만 입력한 뒤
   - `_make_output_path`로 파일명을 정해 저장한다.
3. Flask 라우트(`app/routes/*.py`)에서 이 함수를 호출한다.
4. 이 SKILL.md에 새 섹션(예: `## 7. 신규 문서명`)을 추가해
   - 입력 데이터 구조
   - 템플릿 경로
   - 채우는 순서
   를 간단히 문서화한다.

이 스킬을 따를 때, `.hwp` 문서는 **항상 템플릿 구조를 보존한 채 내용만 교체**하도록 구현해야 한다.

