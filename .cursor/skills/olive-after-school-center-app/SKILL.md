---
name: olive-after-school-center-app
description: Guides development of the Olive Youth After-School Center Activity App (올리브청소년방과후 센터 활동앱). Use when building or modifying this app, adding 일일활동일지/품위서/계획서 features, integrating OpenAI image analysis, or generating HWP documents for social workers.
---

# 올리브청소년방과후 센터 활동앱 개발

## 적용 시점

- 이 프로젝트(방과후센터 앱)를 만들거나 수정할 때
- 일일활동일지생성, 품위서, 계획서, 기타 메뉴 관련 기능을 구현·변경할 때
- OpenAI로 사진 분석·영수증 인식·계획서 문구 생성이 필요할 때
- 한글(HWP) 문서를 생성·저장·템플릿 채우기가 필요할 때

## 필수 참고 문서

1. **CONTEXT.md** (프로젝트 루트): 앱 개요, 첫 화면(제목·버튼 4개), 기능별 스펙(일지/품위서/계획서/기타), 기술·제약, 플로우. 구현 전 해당 기능 섹션을 반드시 확인한다.
2. **한글_템플릿_내용만_바꾸기.md**: 예시 한글 파일을 열어 표·형식 유지하고 내용만 바꿀 때(hwp_open → hwp_replace_text / hwp_fill_cells → hwp_save) 절차와 채팅 요청 예시.

## 화면·플로우 준수

- **메인**: 제목 `올리브청소년방과후 센터 활동앱`, 버튼 4개(일일활동일지생성, 품위서, 계획서, 기타).
- **일일활동일지**: 사진 3~4장 업로드, 시간/장소/특이사항 입력 가능 → OpenAI Vision으로 활동기록 생성 → 정해진 서식 HWP 출력.
- **품위서**: 영수증 사진 업로드 → OpenAI로 내용 추출 → 품위서 HWP 생성 → 계획서 생성 여부 묻기, 요청 시 사진 최대 6장 추가 후 구간별 계획서 생성.
- **계획서**: 품위서 내용(2번에서 작성된 것) 자동 반영 + 사진 → 구간별 계획서(목적·목표·진행내용) HWP 생성.
- **기타**: 버튼만 두고 화면/기능은 추후 정의.

## 기술 선택

- **문서 출력**: 한글(HWP) 기본. **앱에서 만드는 모든 최종 결과물은 HWP 파일(.hwp)로 저장**한다. HWP MCP로 생성·저장·템플릿 채우기(hwp_create, hwp_open, hwp_replace_text, hwp_fill_cells, hwp_save 등).
- **이미지·텍스트 생성**: OpenAI API 사용(일지: 사진→활동 설명, 품위서: 영수증→품목/금액, 계획서: 품위서+사진→목적/목표/진행내용). API 키는 환경 변수 또는 설정에서 로드, 코드에 노출 금지.

## 품질 체크

- **최종 산출물**: 일일활동일지·품위서·계획서 등 생성 문서는 반드시 **.hwp 파일**로 저장했는지 확인.
- UI·메시지는 한국어.
- 일지 3~4장, 품위서 후 계획서 시 사진 최대 6장 등 CONTEXT.md의 사진 수·플로우를 유지한다.
- 서식·구간(구대) 정의가 필요하면 CONTEXT.md의 "추후 정의할 항목"을 참고해 한 번에 확정하지 말고, 필요한 최소 필드만 구현 후 확장한다.
