"""
일일활동일지 / 품위서 / 계획서 공통 사진 삽입 워커

서버 프로세스와 분리된 별도 프로세스에서 HWP 문서를 다시 열고,
클립보드(BMP) + Ctrl+V 방식으로 사진을 붙여넣습니다.

사용법:
  python photo_insert_worker.py daily_log <output_path> <photo1> [photo2]
  python photo_insert_worker.py purchase_doc <output_path> <photo1> [photo2] [photo3]
  python photo_insert_worker.py plan <output_path> <photo1> [photo2] ... [photo6]
"""
import ctypes
import io
import os
import sys
import tempfile
import time

import pythoncom
import win32clipboard
import win32con
import win32gui
from PIL import Image, ImageOps


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
    return r"C:\Users\user\AppData\Roaming\Cursor\mcp-servers\hwp-mcp"


HWP_MCP_PATH = _resolve_hwp_mcp_path()
if HWP_MCP_PATH not in sys.path:
    sys.path.insert(0, HWP_MCP_PATH)

_SECURITY_MODULE_DLL = os.path.join(
    HWP_MCP_PATH,
    "security_module",
    "FilePathCheckerModuleExample.dll",
)

from src.tools.hwp_controller import HwpController


def _log(msg: str):
    print(f"[photo_worker] {msg}", flush=True)


def _find_hwp_window():
    hwnds = []

    def _cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and ("Hwp" in title or "HWP" in title or "한글" in title):
                hwnds.append(hwnd)
        return True

    win32gui.EnumWindows(_cb, None)
    return hwnds[0] if hwnds else None


def _force_foreground(hwnd):
    if not hwnd:
        return

    user32 = ctypes.windll.user32
    fg = user32.GetForegroundWindow()
    fg_tid = user32.GetWindowThreadProcessId(fg, None)
    my_tid = ctypes.windll.kernel32.GetCurrentThreadId()

    if fg_tid != my_tid:
        user32.AttachThreadInput(my_tid, fg_tid, True)
    user32.SetForegroundWindow(hwnd)
    if fg_tid != my_tid:
        user32.AttachThreadInput(my_tid, fg_tid, False)
    time.sleep(0.15)


def _send_ctrl_v():
    user32 = ctypes.windll.user32
    vk_control, vk_v = 0x11, 0x56
    user32.keybd_event(vk_control, 0, 0, 0)
    user32.keybd_event(vk_v, 0, 0, 0)
    time.sleep(0.03)
    user32.keybd_event(vk_v, 0, 2, 0)
    user32.keybd_event(vk_control, 0, 2, 0)
    time.sleep(0.5)


def _orient_image_for_display(img: Image.Image) -> Image.Image:
    """
    EXIF Orientation(휴대폰 세로/가로 촬영 메타데이터)을 픽셀에 반영합니다.
    붙여넣기 전에 적용해 사람·장면이 통상적인 '위' 방향으로 보이게 합니다.
    EXIF가 없거나 이미 반영된 파일은 그대로 둡니다.
    """
    try:
        return ImageOps.exif_transpose(img)
    except Exception:
        return img


def _prepare_image(image_path: str, index: int, target_w: int = 980, target_h: int = 680) -> str:
    """
    셀 크기에 맞게 이미지를 미리 축소합니다.
    비율을 유지한 채 흰 배경 캔버스에 맞춰 넣습니다. 원본 색상 유지.
    삽입 전 EXIF 방향을 보정합니다.
    """
    img = Image.open(image_path)
    img = _orient_image_for_display(img)
    img = img.convert("RGB")
    img.thumbnail((target_w, target_h))

    canvas = Image.new("RGB", (target_w, target_h), (255, 255, 255))
    x = (target_w - img.width) // 2
    y = (target_h - img.height) // 2
    canvas.paste(img, (x, y))

    out_path = os.path.join(tempfile.gettempdir(), f"olive_photo_{index}_{os.getpid()}.bmp")
    canvas.save(out_path, "BMP")
    return out_path


def _set_clipboard_bmp(bmp_path: str):
    with Image.open(bmp_path).convert("RGB") as img:
        buf = io.BytesIO()
        img.save(buf, "BMP")
        dib = buf.getvalue()[14:]

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, dib)
    finally:
        win32clipboard.CloseClipboard()


def _paste_photo(hwnd, bmp_path: str):
    _set_clipboard_bmp(bmp_path)
    time.sleep(0.1)
    _force_foreground(hwnd)
    _send_ctrl_v()


def _clear_current_cell(ctrl):
    """현재 셀의 내용을 비웁니다."""
    pos = ctrl.hwp.GetPos()
    ctrl.hwp.HAction.Run("TableSelCell")
    ctrl.hwp.HAction.Run("EditCut")
    time.sleep(0.1)
    ctrl.hwp.SetPos(pos[0], 0, 0)
    time.sleep(0.1)


def _insert_daily_log_photos(ctrl, hwnd, photo_paths: list):
    """일일활동일지: 사진1입력, 사진2입력 자리에 사진 붙이기(문구는 제거)."""
    if len(photo_paths) == 1:
        chosen = [photo_paths[0], photo_paths[0]]
    else:
        chosen = photo_paths[:2]

    prepared = [_prepare_image(p, i + 1) for i, p in enumerate(chosen)]

    # AllReplace는 문서 전체 치환이라, 찾기→붙여넣기 위치를 잡기 어렵습니다.
    # 플레이스홀더를 문서에 거의 없을 앵커 문자열로 바꾼 뒤 해당 셀만 비우고 이미지를 넣습니다.
    slot_markers = (
        ("사진1입력", "@@OLIVE_DAILY_PHOTO_1@@", prepared[0], "왼쪽"),
        ("사진2입력", "@@OLIVE_DAILY_PHOTO_2@@", prepared[1], "오른쪽"),
    )
    for placeholder, marker, prep, name in slot_markers:
        try:
            ctrl.replace_text(placeholder, marker, replace_all=True)
        except Exception as e:
            _log(f"'{placeholder}' → 앵커 치환 실패: {e}")

    for placeholder, marker, prep, name in slot_markers:
        ctrl.hwp.HAction.Run("MoveDocBegin")
        if not ctrl.find_text(marker):
            _log(f"앵커 '{marker}' 를 찾지 못했습니다. 양식에 '{placeholder}' 문구가 있는지 확인하세요.")
            continue
        _clear_current_cell(ctrl)
        _paste_photo(hwnd, prep)
        # 붙여넣기 후에도 같은 셀 아래 줄 등에 앵커 문자열이 남는 경우가 있어 전부 삭제
        try:
            ctrl.replace_text(marker, "", replace_all=True)
        except Exception as e:
            _log(f"앵커 문구 제거 실패 ({marker}): {e}")
        _log(f"{name} 셀 삽입 완료(플레이스홀더·앵커 제거)")

    # 남은 앵커가 있으면 최종 정리(한 번 더)
    for _, marker, _, _ in slot_markers:
        try:
            ctrl.replace_text(marker, "", replace_all=True)
        except Exception:
            pass


def _insert_purchase_doc_photos(ctrl, hwnd, photo_paths: list):
    """품의서: 양식의 플레이스홀더 셀에 영수증 이미지를 붙여넣습니다.

    - 신규(슬롯별 후보 순서대로 시도): 영수증1사진 / 영수증2사진(또는 영수증사진2) / 영수증3사진
    - 플레이스홀더 문구는 붙여넣기 후 제거(일일활동일지 사진 슬롯과 동일한 앵커 방식)
    - 구형: '영수증사진' 아래 영역 또는 '영수증입력' 등 (1장만)
    """
    # 양식마다 2번 슬롯 문자열이 '영수증2사진' 또는 '영수증사진2'로 다를 수 있음
    slot_label_groups = (
        ("영수증1사진",),
        ("영수증2사진", "영수증사진2"),
        ("영수증3사진", "영수증사진3"),
    )
    n = len(photo_paths)
    placed_count = 0

    # 업로드 장수만큼 슬롯 문구를 앵커로 바꾼 뒤 셀 비우기·붙여넣기·앵커 제거
    for idx in range(n):
        marker = f"@@OLIVE_PURCHASE_PHOTO_{idx + 1}@@"
        for label in slot_label_groups[idx]:
            try:
                ctrl.replace_text(label, marker, replace_all=True)
            except Exception as e:
                _log(f"'{label}' → 앵커 치환 실패: {e}")

    for idx, path in enumerate(photo_paths):
        marker = f"@@OLIVE_PURCHASE_PHOTO_{idx + 1}@@"
        ctrl.hwp.HAction.Run("MoveDocBegin")
        if not ctrl.find_text(marker):
            _log(f"앵커 '{marker}' 를 찾지 못했습니다. 슬롯 {idx + 1} placeholder 확인.")
            continue
        prepared = _prepare_image(path, idx + 1, target_w=980, target_h=1200)
        _clear_current_cell(ctrl)
        _paste_photo(hwnd, prepared)
        try:
            ctrl.replace_text(marker, "", replace_all=True)
        except Exception as e:
            _log(f"앵커 제거 실패 ({marker}): {e}")
        placed_count += 1
        _log(f"영수증 {idx + 1} 삽입 완료(플레이스홀더·앵커 제거): {os.path.basename(path)}")

    for i in range(1, n + 1):
        try:
            ctrl.replace_text(f"@@OLIVE_PURCHASE_PHOTO_{i}@@", "", replace_all=True)
        except Exception:
            pass

    if placed_count == len(photo_paths):
        return

    if placed_count > 0:
        _log(
            f"신규 슬롯에 {placed_count}/{len(photo_paths)}장만 삽입되었습니다. "
            "여러 장을 쓰려면 양식에 '영수증1사진', '영수증2사진'(또는 '영수증사진2'), '영수증3사진'을 두세요."
        )
        return

    path = photo_paths[0]
    prepared = _prepare_image(path, 1, target_w=980, target_h=1200)

    try:
        ctrl.hwp.HAction.Run("MoveDocBegin")
        if ctrl.find_text("영수증사진"):
            ctrl.hwp.HAction.Run("MoveLineEnd")
            ctrl.hwp.HAction.Run("MoveDown")
            ctrl.hwp.HAction.Run("MoveDown")
            _paste_photo(hwnd, prepared)
            _log(f"영수증사진 아래 영역에 영수증 삽입 완료: {os.path.basename(path)}")
            return
    except Exception as e:
        _log(f"영수증사진 기준 삽입 실패, fallback 사용: {e}")

    for placeholder, move_right in (("영수증입력", False), ("영수증 사진", True), ("영수증", True)):
        ctrl.hwp.HAction.Run("MoveDocBegin")
        if not ctrl.find_text(placeholder):
            continue
        ctrl.hwp.HAction.Run("TableSelCell")
        ctrl.hwp.HAction.Run("Cancel")
        if move_right:
            ctrl.hwp.HAction.Run("TableRightCell")
            ctrl.hwp.HAction.Run("TableSelCell")
            ctrl.hwp.HAction.Run("Cancel")
        _clear_current_cell(ctrl)
        _paste_photo(hwnd, prepared)
        _log(f"영수증 삽입 완료: {os.path.basename(path)} (placeholder='{placeholder}')")
        return

    _log(
        "영수증 사진 placeholder를 찾지 못했습니다. "
        "품의서양식.hwp에 '영수증1사진', '영수증2사진', '영수증3사진' 또는 '영수증사진'·'영수증입력'이 있는지 확인해 주세요."
    )


def _insert_plan_photos(ctrl, hwnd, photo_paths: list):
    """계획서: 사진1, 사진2, ... 사진6 placeholder"""
    for i, path in enumerate(photo_paths[:6], start=1):
        label = f"사진{i}"
        ctrl.hwp.HAction.Run("MoveDocBegin")
        if not ctrl.find_text(label):
            _log(f"'{label}' 셀을 찾을 수 없습니다.")
            continue
        prepared = _prepare_image(path, i, target_w=600, target_h=600)
        _clear_current_cell(ctrl)
        _paste_photo(hwnd, prepared)
        _log(f"{label} 삽입 완료: {os.path.basename(path)}")


def insert_photos(doc_type: str, output_path: str, photo_paths: list):
    photo_paths = [os.path.abspath(p) for p in photo_paths if p and os.path.isfile(p)]
    if not photo_paths:
        _log("유효한 사진이 없습니다.")
        _log("STATUS:ERROR")
        return

    pythoncom.CoInitialize()
    ctrl = None
    try:
        ctrl = HwpController()
        if not ctrl.connect(visible=True, register_security_module=False):
            _log("한글 연결 실패")
            _log("STATUS:ERROR")
            return

        # 보안 모듈 등록 – 사진 삽입 시 파일 접근 팝업 방지 (hwp_service와 동일)
        try:
            ctrl.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        except Exception:
            if os.path.exists(_SECURITY_MODULE_DLL):
                try:
                    ctrl.hwp.RegisterModule("FilePathCheckerModuleExample", _SECURITY_MODULE_DLL)
                except Exception:
                    pass

        # 파일 접근·확인 대화상자 억제 (0x00010000|0x00010010|0x00000001)
        try:
            ctrl.hwp.SetMessageBoxMode(0x00010000 | 0x00010010 | 0x00000001)
        except Exception:
            pass

        ctrl.open_document(output_path)
        time.sleep(0.9)

        hwnd = _find_hwp_window()
        _log(f"doc_type={doc_type}, hwnd={hwnd}")

        if doc_type == "daily_log":
            _insert_daily_log_photos(ctrl, hwnd, photo_paths)
        elif doc_type == "purchase_doc":
            _insert_purchase_doc_photos(ctrl, hwnd, photo_paths)
        elif doc_type == "plan":
            _insert_plan_photos(ctrl, hwnd, photo_paths)
        else:
            _log(f"알 수 없는 doc_type: {doc_type}")
            _log("STATUS:ERROR")
            return

        ctrl.save_document(output_path)
        _log(f"저장 완료: {output_path}")
        _log("STATUS:SUCCESS")
    except Exception as e:
        _log(f"오류: {e}")
        _log("STATUS:ERROR")
    finally:
        try:
            if ctrl:
                try:
                    ctrl.hwp.SetMessageBoxMode(0x00020000)
                except Exception:
                    pass
                ctrl.close_document(save=False, suppress_dialog=True)
                ctrl.close_all_documents(save=False, suppress_dialog=True)
                try:
                    ctrl.hwp.SetMessageBoxMode(0x00020000)
                    try:
                        ctrl.hwp.Quit()
                        _log("한글 종료(Quit) 실행")
                    except (AttributeError, Exception):
                        try:
                            ctrl.hwp.Run("FileExit")
                            _log("한글 종료(Run FileExit) 실행")
                        except Exception:
                            ctrl.hwp.HAction.Run("FileExit")
                            _log("한글 종료(HAction FileExit) 실행")
                except Exception as ex:
                    _log(f"한글 종료 시도 실패: {ex}")
            time.sleep(0.5)
        except Exception:
            pass
        try:
            if ctrl:
                ctrl.disconnect()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    if len(sys.argv) < 4:
        _log("usage: photo_insert_worker.py <doc_type> <output_path> <photo1> [photo2] ...")
        _log("  doc_type: daily_log | purchase_doc | plan")
        return

    doc_type = sys.argv[1].lower()
    output_path = sys.argv[2]
    photos = sys.argv[3:]

    if doc_type not in ("daily_log", "purchase_doc", "plan"):
        _log(f"doc_type은 daily_log, purchase_doc, plan 중 하나여야 합니다: {doc_type}")
        return

    insert_photos(doc_type, output_path, photos)


if __name__ == "__main__":
    main()
