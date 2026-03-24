"""
HWP 문서 서비스 – HwpController를 직접 사용하여 템플릿을 채우고 저장합니다.

템플릿의 각 레이블 셀을 찾아 옆 셀(또는 아래 셀)에 내용을 입력합니다.
보안 DLL이 없으면 SetMessageBoxMode로 대화상자를 억제합니다.
"""
import sys
import os
import re
import subprocess
import tempfile
import threading
import time
from datetime import datetime
from contextlib import contextmanager

import pythoncom


def _resolve_hwp_mcp_path() -> str:
    """
    현재 Windows 사용자(AppData\Roaming)를 기준으로
    Cursor MCP 서버의 hwp-mcp 경로를 동적으로 계산합니다.

    - 기본값: %APPDATA%\Cursor\mcp-servers\hwp-mcp
    - 환경변수 HWP_MCP_PATH 가 설정되어 있으면 그 값을 우선 사용
    """
    env_path = os.environ.get("HWP_MCP_PATH")
    if env_path:
        return env_path

    appdata = os.environ.get("APPDATA")
    if appdata:
        return os.path.join(appdata, "Cursor", "mcp-servers", "hwp-mcp")

    # APPDATA 가 없을 경우, 기존 하드코딩 경로를 최후 fallback 으로 사용
    return r'C:\Users\user\AppData\Roaming\Cursor\mcp-servers\hwp-mcp'


HWP_MCP_PATH = _resolve_hwp_mcp_path()
if HWP_MCP_PATH not in sys.path:
    sys.path.insert(0, HWP_MCP_PATH)


try:
    from src.tools.hwp_controller import HwpController
    HWP_AVAILABLE = True
except ImportError:
    HWP_AVAILABLE = False

HWP_LOCK = threading.Lock()
ENABLE_DAILY_LOG_PHOTO_INSERTION = False


_SECURITY_MODULE_DLL = os.path.join(
    HWP_MCP_PATH,
    'security_module', 'FilePathCheckerModuleExample.dll'
)


def _get_controller():
    if not HWP_AVAILABLE:
        raise RuntimeError("HWP 모듈을 찾을 수 없습니다. hwp-mcp 경로를 확인해 주세요.")
    ctrl = HwpController()
    if not ctrl.connect(visible=True, register_security_module=False):
        raise RuntimeError("한글 프로그램에 연결하지 못했습니다. 한글이 설치되어 있는지 확인해 주세요.")

    # 보안 모듈 등록 – 파일 열기 보안 경고창 방지 (한컴 문서: 레지스트리 + RegisterModule 타입/이름)
    try:
        # 한컴 공식: RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample") — 경로는 레지스트리에서 조회
        ctrl.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        print(f"[hwp_service] 보안 모듈 등록 성공 (레지스트리)")
    except Exception as e1:
        # 레지스트리 미등록 환경 대비: 경로로 직접 등록 시도
        if os.path.exists(_SECURITY_MODULE_DLL):
            try:
                ctrl.hwp.RegisterModule("FilePathCheckerModuleExample", _SECURITY_MODULE_DLL)
                print(f"[hwp_service] 보안 모듈 등록 성공 (경로 직접)")
            except Exception as e2:
                print(f"[hwp_service] 보안 모듈 등록 실패 (무시): {e2}")
        else:
            print(f"[hwp_service] 보안 모듈 등록 실패 (무시): {e1}")
    return ctrl


@contextmanager
def _controller_session():
    """
    COM 초기화 + HWP 컨트롤러를 안전하게 연결/해제합니다.
    단일 스레드 잠금으로 HWP COM 안정성을 유지합니다.
    """
    pythoncom.CoInitialize()
    with HWP_LOCK:
        ctrl = _get_controller()
        # 연결 즉시 대화상자 억제 (파일 접근 보안 경고 포함)
        # 0x00010000: 예/아니오에서 '예' 자동 선택
        # 0x00010010: 특정 대화상자(파일 접근 경고 등) 억제
        # 0x00000001: 확인 버튼 자동 선택
        try:
            ctrl.hwp.SetMessageBoxMode(0x00010000 | 0x00010010 | 0x00000001)
        except Exception:
            pass
        try:
            yield ctrl
        finally:
            try:
                ctrl.hwp.SetMessageBoxMode(0x00000000)
            except Exception:
                pass
            try:
                ctrl.close_all_documents(save=False, suppress_dialog=True)
            except Exception:
                pass
            try:
                ctrl.hwp.SetMessageBoxMode(0x00100000)
                try:
                    ctrl.hwp.Quit()
                    print("[hwp_service] 한글 종료(Quit) 실행")
                except (AttributeError, Exception):
                    try:
                        ctrl.hwp.Run("FileExit")
                        print("[hwp_service] 한글 종료(Run FileExit) 실행")
                    except Exception:
                        ctrl.hwp.HAction.Run("FileExit")
                        print("[hwp_service] 한글 종료(HAction FileExit) 실행")
            except Exception as e:
                print(f"[hwp_service] 한글 종료 시도 실패: {e}")
            try:
                time.sleep(0.2)
            except Exception:
                pass
            try:
                ctrl.disconnect()
            except Exception:
                pass
            pythoncom.CoUninitialize()


def _make_output_path(prefix: str, output_dir: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_dir, f"{prefix}_{stamp}.hwp")


def get_photo_insert_log_path(output_path: str) -> str:
    return f"{output_path}.photo.log"


def _spawn_photo_worker(doc_type: str, output_path: str, photo_paths: list):
    """
    사진 삽입은 별도 프로세스(클립보드 BMP + Ctrl+V)에서 처리합니다.
    doc_type: daily_log | purchase_doc | plan
    """
    valid = [os.path.abspath(p) for p in photo_paths if p and os.path.isfile(p)]
    if not valid:
        return

    worker_path = os.path.join(os.path.dirname(__file__), "photo_insert_worker.py")
    if not os.path.exists(worker_path):
        print(f"[hwp_service] photo worker 없음: {worker_path}")
        return

    n = 6 if doc_type == "plan" else 2 if doc_type == "daily_log" else 3
    to_use = valid[:n]
    log_path = get_photo_insert_log_path(output_path)
    try:
        with open(log_path, "w", encoding="utf-8") as log_fp:
            log_fp.write(f"[photo_worker] STATUS:STARTED doc_type={doc_type}\n")
            log_fp.write(f"[photo_worker] output={output_path}\n")
            log_fp.write(f"[photo_worker] photos={len(to_use)}\n")
        log_fp = open(log_path, "a", encoding="utf-8")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            [sys.executable, worker_path, doc_type, output_path, *to_use],
            stdout=log_fp,
            stderr=log_fp,
            creationflags=creationflags,
            close_fds=False,
        )
        print(f"[hwp_service] 사진 삽입 워커 시작 ({doc_type}): {os.path.basename(output_path)}")
    except Exception as e:
        print(f"[hwp_service] 사진 삽입 워커 시작 실패: {e}")



def _format_number_with_comma(val) -> str:
    """숫자를 천 단위 쉼표가 있는 문자열로 포맷합니다."""
    try:
        if isinstance(val, (int, float)):
            return f"{int(val):,}"
        s_val = str(val).strip()
        if s_val.isdigit():
            return f"{int(s_val):,}"
        return s_val
    except ValueError:
        return str(val)


def _fill_label(ctrl, label: str, value: str, direction: str = "right"):
    """레이블 옆 셀에 값을 입력합니다. 표에서 못 찾으면 문서에서 해당 문구를 값으로 치환합니다. 실패해도 예외를 던지지 않습니다."""
    if not value:
        return
    try:
        ok, msg = ctrl.fill_cell_next_to_label(label, value, direction, mode="replace")
        if not ok:
            # 표 레이블이 없으면, 문서에 있는 해당 문구를 값으로 치환 (품위서 등 플레이스홀더 양식 대응)
            try:
                if ctrl.replace_text(label, value, replace_all=True):
                    return  # 치환 성공 시 로그 생략
            except Exception:
                pass
            print(f"[hwp_service] fill_label '{label}': {msg}")
    except Exception as e:
        print(f"[hwp_service] fill_label '{label}' 오류: {e}")


def _replace_first_date_in_doc(ctrl, value: str):
    """문서에서 처음 발견되는 YYYY.MM.DD 날짜를 지정 값으로 치환합니다."""
    if not value:
        return

    normalized = value.strip().replace("-", ".")
    if not normalized:
        return

    try:
        doc_text = ctrl.get_text() or ""
        match = re.search(r"\d{4}\.\d{2}\.\d{2}", doc_text)
        if match:
            ctrl.replace_text(match.group(0), normalized, replace_all=True)
        else:
            print(f"[hwp_service] 문서 날짜를 찾지 못했습니다: {normalized}")
    except Exception as e:
        print(f"[hwp_service] 문서 날짜 치환 오류: {e}")


def _fill_korean_parenthesized_date(ctrl, raw_date: str):
    """
    품위서/계획서 양식에 있는 '( 2026년. 월. 일)' 같은 레이블을
    실제 날짜로 치환합니다.

    raw_date 예시:
      - '2026-03-17'
      - '2026.03.17'
    결과 예시:
      - '( 2026년 3월 17일)'
    """
    if not raw_date:
        return

    s = (raw_date or "").strip().replace("-", ".").replace("/", ".")
    parts = [p for p in s.split(".") if p]
    if len(parts) < 3:
        return

    year, month, day = parts[0], parts[1][:2], parts[2][:2]
    # 사용자 요청: '2026년 3월 16일' 처럼 점(.)과 앞의 0 없이 출력
    try:
        month_int = int(month)
        day_int = int(day)
    except ValueError:
        month_int, day_int = month, day
    formatted = f"( {year}년 {month_int}월 {day_int}일)"

    # 템플릿에 고정으로 들어 있는 예시 텍스트를 실제 날짜로 치환
    try:
        ctrl.replace_text("( 2026년. 월. 일)", formatted, replace_all=True)
    except Exception:
        # 다른 연도 예시가 있을 경우를 대비해, '년. 월. 일)' 패턴 전체를 통째로 교체 시도
        try:
            doc_text = ctrl.get_text() or ""
            import re as _re
            m = _re.search(r"\(\s*\d{4}년\.?\s*월\.?\s*일\)", doc_text)
            if m:
                ctrl.replace_text(m.group(0), formatted, replace_all=True)
        except Exception:
            pass

def _fill_purchase_placeholder_cell(ctrl, placeholder: str, value: str):
    """
    품위서 양식의 플레이스홀더 문자열(예: 영수증1합계)을 값으로 바꿉니다.
    값이 비어 있으면 해당 문구만 제거합니다.
    """
    if not placeholder:
        return
    v = value if value is not None else ""
    try:
        ctrl.replace_text(placeholder, v, replace_all=True)
    except Exception as e:
        print(f"[hwp_service] replace '{placeholder}': {e}")
    try:
        ok, msg = ctrl.fill_cell_next_to_label(placeholder, v, "right", mode="replace")
        if not ok and v:
            print(f"[hwp_service] fill next to '{placeholder}': {msg}")
    except Exception:
        pass


def _fill_purchase_receipt_slots(ctrl, receipt_parts: list):
    """
    품위서양식.hwp 내 장별 칸 채움 (템플릿에 동일 문자열이 있어야 함).
    - 영수증1합계 / 영수증1거래처 … 영수증3
    """
    slots = (
        ("영수증1합계", "영수증1거래처"),
        ("영수증2합계", "영수증2거래처"),
        ("영수증3합계", "영수증3거래처"),
    )
    parts = receipt_parts or []
    for i, (sum_ph, store_ph) in enumerate(slots):
        part = parts[i] if i < len(parts) else None
        amt = (part.get("total_amount") or "") if part else ""
        store = (part.get("store_name") or "") if part else ""
        amt_fmt = _format_number_with_comma(amt) if amt else ""
        _fill_purchase_placeholder_cell(ctrl, sum_ph, amt_fmt)
        _fill_purchase_placeholder_cell(ctrl, store_ph, store)


def _fill_purchase_table(ctrl, items: list):
    """품위서 표의 헤더(품명)를 기준으로 품목 행 데이터를 채웁니다. 추출된 품목 전부를 채웁니다."""
    rows = []
    for item in items:
        rows.append([
            str(item.get("name", "") or ""),
            str(item.get("qty", "") or ""),
            str(item.get("unit", "") or ""),
            _format_number_with_comma(item.get("unit_price", "") or ""),
            _format_number_with_comma(item.get("amount", "") or ""),
            str(item.get("note", "") or ""),
        ])

    if not rows:
        return

    try:
        if not ctrl.find_text("품명"):
            print("[hwp_service] 품위서 표 헤더 '품명'을 찾을 수 없습니다.")
            return
        if not ctrl.fill_table_with_data(rows, start_row=2, start_col=1, has_header=False):
            print("[hwp_service] 품위서 표 데이터 입력에 실패했습니다.")
    except Exception as e:
        print(f"[hwp_service] 품위서 표 입력 오류: {e}")


# ──────────────────────────────────────────────
# 일일활동일지
# ──────────────────────────────────────────────
def _strip_time_prefix(text: str) -> str:
    """내용 앞에 붙은 '14:30~15:00 ' 형태의 시간대 접두어를 제거합니다."""
    import re
    return re.sub(r'^\d{1,2}:\d{2}[~\-]\d{1,2}:\d{2}\s*', '', text.strip())


_EMPTY_WORDS = {"없음", "해당없음", "없음.", "해당 없음", "없음,", "-", "N/A", "n/a",
                "특이사항 없음", "이상 없음", "특이사항없음", "이상없음"}


def _is_empty_content(text: str) -> bool:
    """내용이 비어 있거나 '없음' 계열 단어인지 확인합니다."""
    t = text.strip().strip("[]().")
    return not t or t in _EMPTY_WORDS


def _build_activity_text(acts: dict) -> str:
    """시간대별 활동 내용을 하나의 텍스트 블록으로 합칩니다.
    - 시간대 헤더([14:30~15:00 등원 및 자유활동]) 없이 내용만 출력
    - 내용이 없거나 '없음'인 항목은 건너뜁니다.
    """
    keys = [
        "activity_1430",
        "activity_1500",
        "activity_1600",
        "activity_1700",
        "activity_1800",
    ]
    parts = []
    for key in keys:
        content = _strip_time_prefix(acts.get(key, ""))
        if not _is_empty_content(content):
            parts.append(content)

    special = acts.get("special_note", "").strip()
    if special and not _is_empty_content(special):
        parts.append(f"특이사항: {special}")

    return "\n\n".join(parts)


def create_daily_log(data: dict, output_dir: str) -> str:
    """
    일일활동일지 HWP를 생성합니다.

    data keys:
        date (str), time (str), place (str), special_note (str),
        activity_content (str)   ← 브라우저 표시용 전체 텍스트
        activities (dict)        ← 시간대별 구조화 데이터
        photo_paths (list)       ← 업로드된 사진 경로 목록
    """
    from config import Config
    template = Config.DAILY_LOG_TEMPLATE
    output_path = _make_output_path("일일활동일지", output_dir)

    acts          = data.get("activities", {})
    date_str      = data.get("date", "")
    time_str      = data.get("time", "")
    datetime_str  = f"{date_str}  {time_str}".strip()
    student_names = data.get("student_names", "").strip()
    photo_paths   = data.get("photo_paths", [])

    # 전체 활동 내용을 하나의 블록으로 합치기
    full_activity = _build_activity_text(acts)

    with _controller_session() as ctrl:
        try:
            ctrl.open_document(template)

            # 문서 내 고정 문구 정리
            try:
                ctrl.replace_text("활동명단입력", "활동명단 대상자 이름", replace_all=True)
                ctrl.replace_text("활동일지 생성내역 입력", "", replace_all=True)
            except Exception:
                pass

            # ── 헤더 정보 채우기: 일시, 이용시간, 이용자(활동명단) ──
            _fill_label(ctrl, "일시", datetime_str)
            _fill_label(ctrl, "이용시간", time_str)
            _fill_label(ctrl, "이용자", student_names)

            # ── 프로그램 참여 내용 셀에 전체 활동 내역 삽입 ──
            # "참여 내용" 헤더에서 MoveDown → 실제 참여내용 셀(ListId=15)로 이동
            if full_activity and ctrl.find_text("참여 내용"):
                ctrl.hwp.HAction.Run("MoveDown")          # 헤더 → 내용 셀로 이동
                content_pos = ctrl.hwp.GetPos()           # 셀 위치(ListId) 저장

                # 기존 셀 내용 삭제
                ctrl.hwp.HAction.Run("TableSelCell")
                ctrl.hwp.HAction.Run("EditCut")

                # EditCut 후 커서가 이동할 수 있으므로 저장한 위치로 복귀
                ctrl.hwp.SetPos(content_pos[0], 0, 0)

                # HWP 표 안에서 BreakPara가 동작하지 않는 경우가 있으므로
                # \r 을 포함한 텍스트를 한 번에 삽입 (HWP InsertText에서 \r = 줄바꿈)
                text_for_hwp = full_activity.replace('\n', '\r')
                ctrl.hwp.HAction.GetDefault(
                    "InsertText",
                    ctrl.hwp.HParameterSet.HInsertText.HSet
                )
                ctrl.hwp.HParameterSet.HInsertText.Text = text_for_hwp
                ctrl.hwp.HAction.Execute(
                    "InsertText",
                    ctrl.hwp.HParameterSet.HInsertText.HSet
                )

            # 사진 삽입 자동화는 현재 HWP 포커스/붙여넣기 안정성 이슈로 잠시 비활성화합니다.
            if photo_paths and ENABLE_DAILY_LOG_PHOTO_INSERTION:
                print("[hwp_service] 활동 사진 자동 삽입 활성화")

            ctrl.save_document(output_path)
            ctrl.close_document(save=False, suppress_dialog=True)
        except Exception as e:
            try:
                ctrl.hwp.SetMessageBoxMode(0x00000000)
            except Exception:
                pass
            raise RuntimeError(f"일일활동일지 생성 오류: {e}")

    # 사진 삽입은 저장 후 별도 프로세스(클립보드 BMP 붙여넣기)에서 비동기 처리
    _spawn_photo_worker("daily_log", output_path, photo_paths)
    return output_path


# ──────────────────────────────────────────────
# 품위서
# ──────────────────────────────────────────────
def create_purchase_doc(data: dict, output_dir: str, image_paths: list = None) -> str:
    """
    품위서 HWP를 생성합니다.
    2페이지에 영수증 사진을 중앙에 배치합니다.

    data keys:
        purchase_date (str), store_name (str),
        items (list[dict]):  [{name, qty, unit, unit_price, amount, note}, ...]
        total_amount (str)
        receipt_parts (list, optional): 장별 {store_name, total_amount} — 영수증N합계·거래처 칸용
    image_paths: 업로드 순서대로 최대 3장 — 영수증1사진, 영수증2사진(또는 영수증사진2), 영수증3사진 슬롯에 삽입
    """
    from config import Config
    template = Config.PURCHASE_DOC_TEMPLATE
    output_path = _make_output_path("품위서", output_dir)
    image_paths = image_paths or []

    with _controller_session() as ctrl:
        try:
            ctrl.open_document(template)

            # 실제 양식 기준:
            # - 상단 문장의 날짜는 첫 YYYY.MM.DD 날짜 문자열을 구입일자로 치환
            # - 만약 '구입일자' 레이블형 양식이면 그 값도 함께 채움
            # - 표는 '품명' 헤더 아래부터 직접 채움
            # - 합계는 '합  계'가 나오는 두 곳 모두 오른쪽 셀에 입력
            purchase_date = data.get("purchase_date", "")
            _replace_first_date_in_doc(ctrl, purchase_date)
            _fill_label(ctrl, "구입일자", purchase_date)
            _fill_korean_parenthesized_date(ctrl, purchase_date)
            receipt_parts = data.get("receipt_parts") or []
            _fill_purchase_receipt_slots(ctrl, receipt_parts)
            # 여러 장일 때 store_name은 "매장1 | 매장2" 병합값이라,
            # 양식의 공통 '거래처' 칸에 넣으면 첫 행 거래처 셀까지 덮어쓸 수 있음 → 1번 영수증 매장만 사용
            if receipt_parts:
                store_for_label = (receipt_parts[0].get("store_name") or "").strip()
            else:
                store_for_label = (data.get("store_name") or "").strip()
            _fill_label(ctrl, "거래처", store_for_label, "down")
            _fill_purchase_table(ctrl, data.get("items", []))
            total_amount = _format_number_with_comma(data.get("total_amount", ""))
            _fill_label(ctrl, "합  계", total_amount)
            if total_amount:
                try:
                    # 두 번째 "합  계" 레이블에 값 채우기 시도
                    ctrl.fill_cell_next_to_label("합  계", total_amount, "right", occurrence=2, mode="replace")
                except Exception:
                    # 두 번째 레이블이 없으면 다음 시도
                    pass
                try:
                    # "합계" 레이블에 값 채우기 시도
                    ctrl.fill_cell_next_to_label("합계", total_amount, "right", mode="replace")
                except Exception:
                    pass

            ctrl.save_document(output_path)
            ctrl.close_document(save=False, suppress_dialog=True)
        except Exception as e:
            try:
                ctrl.hwp.SetMessageBoxMode(0x00000000)
            except Exception:
                pass
            raise RuntimeError(f"품위서 생성 오류: {e}")

    # 영수증 사진 삽입은 저장 후 별도 프로세스(클립보드 BMP 붙여넣기)에서 처리
    if image_paths:
        _spawn_photo_worker("purchase_doc", output_path, image_paths)
    return output_path


# ──────────────────────────────────────────────
# 계획서
# ──────────────────────────────────────────────
def _fill_plan_total(ctrl, formatted_total: str):
    """계획서의 합계 필드를 두 곳 모두 채웁니다."""
    if not formatted_total:
        return
    
    # 첫 번째 "합  계"
    try:
        ctrl.fill_cell_next_to_label("합  계", formatted_total, "right", occurrence=1, mode="replace")
    except Exception as e:
        print(f"[hwp_service] _fill_plan_total '합  계'(1) 오류: {e}")

    # 두 번째 "합  계"
    try:
        ctrl.fill_cell_next_to_label("합  계", formatted_total, "right", occurrence=2, mode="replace")
    except Exception as e:
        print(f"[hwp_service] _fill_plan_total '합  계'(2) 오류: {e}")

    # "합계" (공백 없는 버전)
    try:
        ctrl.fill_cell_next_to_label("합계", formatted_total, "right", mode="replace")
    except Exception as e:
        print(f"[hwp_service] _fill_plan_total '합계' 오류: {e}")


def create_plan(data: dict, output_dir: str, image_paths: list = None) -> str:
    """
    계획서 HWP를 생성합니다.
    2페이지 '사진 1', '사진 2', ... 셀에 업로드 사진을 넣습니다.

    data keys:
        plan_date (str), purpose (str), goal (str),
        program_content (str), expected_effect (str),
        purchase_summary (str),
        purchase_items (list[dict]), purchase_total_amount (str)
    image_paths: 업로드된 사진 경로 목록 (2페이지 표에 삽입)
    """
    from config import Config
    template = Config.PLAN_TEMPLATE
    output_path = _make_output_path("계획서", output_dir)
    image_paths = image_paths or []

    with _controller_session() as ctrl:
        try:
            ctrl.open_document(template)

            plan_date = data.get("plan_date", "")
            _fill_label(ctrl, "작성일자",   plan_date)
            _fill_korean_parenthesized_date(ctrl, plan_date)
            _fill_label(ctrl, "목적",       data.get("purpose", ""))
            _fill_label(ctrl, "목표",       data.get("goal", ""))
            _fill_label(ctrl, "프로그램 내용", data.get("program_content", ""))
            _fill_label(ctrl, "기대효과",   data.get("expected_effect", ""))

            # 거래처
            store_name = data.get("store_name", "")
            _fill_label(ctrl, "거래처", store_name, "down")

            # ── 품위서 표 내용을 계획서 표에도 동일하게 채우기 ──
            purchase_items = data.get("purchase_items") or []
            purchase_total_amount = _format_number_with_comma(data.get("purchase_total_amount", ""))
            if purchase_items:
                # 계획서 양식의 구입내역 표도 헤더가 '품명'으로 동일하다고 가정하고,
                # 품위서와 같은 방식으로 표 행을 채웁니다.
                _fill_purchase_table(ctrl, purchase_items)
            
            # 합계 필드 채우기 (두 곳 모두)
            _fill_plan_total(ctrl, purchase_total_amount)

            # (선택) 요약 텍스트를 별도 칸에 넣고 싶을 때 사용
            purchase_summary = (data.get("purchase_summary") or "").strip()
            if purchase_summary:
                summary_for_hwp = purchase_summary.replace("\n", "\r")
                _fill_label(ctrl, "품위서 내용", summary_for_hwp)
                _fill_label(ctrl, "구입내역", summary_for_hwp)

            ctrl.save_document(output_path)
            ctrl.close_document(save=False, suppress_dialog=True)
        except Exception as e:
            try:
                ctrl.hwp.SetMessageBoxMode(0x00000000)
            except Exception:
                pass
            raise RuntimeError(f"계획서 생성 오류: {e}")

    # 사진 삽입은 저장 후 별도 프로세스(클립보드 BMP 붙여넣기)에서 처리
    if image_paths:
        _spawn_photo_worker("plan", output_path, image_paths)
    return output_path
