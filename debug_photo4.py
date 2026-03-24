"""
HWP 사진 삽입 - 여러 방법 테스트
"""
import sys, time, io, glob, os, ctypes


def _resolve_hwp_mcp_path() -> str:
    env_path = os.environ.get("HWP_MCP_PATH")
    if env_path:
        return env_path
    appdata = os.environ.get("APPDATA")
    if appdata:
        return os.path.join(appdata, "Cursor", "mcp-servers", "hwp-mcp")
    return r'C:\Users\user\AppData\Roaming\Cursor\mcp-servers\hwp-mcp'


sys.path.insert(0, _resolve_hwp_mcp_path())
import pythoncom, win32clipboard, win32con, win32gui, win32api
from PIL import Image
from src.tools.hwp_controller import HwpController

TEMPLATE = r'D:\일일활동일지양식.hwp'
OUT = r'c:\Users\윤미란\Desktop\방과후센터\debug_photo4_result.hwp'

# 테스트 이미지 준비
test_img = r'c:\Users\윤미란\Desktop\방과후센터\test_red.png'
Image.new("RGB", (300, 200), (255, 100, 100)).save(test_img)

imgs = glob.glob(r'c:\Users\윤미란\Desktop\방과후센터\app\static\uploads\*.*')
imgs = [f for f in imgs if os.path.splitext(f)[1].lower() in ('.jpg','.jpeg','.png')]
if imgs:
    test_img = imgs[0]
print(f"Image: {test_img}")

pythoncom.CoInitialize()
ctrl = HwpController()
ctrl.connect(visible=True, register_security_module=False)
ctrl.hwp.SetMessageBoxMode(0x000F0000)
ctrl.open_document(TEMPLATE)
time.sleep(2)
hwp = ctrl.hwp

# 사진 헤더 찾기 + MoveDown
ctrl.find_text("활  동  사  진")
hwp.HAction.Run("MoveDown")
time.sleep(0.2)
p = hwp.GetPos()
print(f"Left cell: ListId={p[0]}")

# ============ Method A: hwp.InsertPicture() direct ============
print("\n=== Method A: hwp.InsertPicture() ===")
try:
    abs_p = os.path.abspath(test_img)
    # InsertPicture(path, pgno, sizetype, isEmbedded, sizeOption)
    # Some HWP versions: InsertPicture(path, Embedded, SizeType, ...)
    r = hwp.InsertPicture(abs_p, 1)
    print(f"  Result: {r}")
except Exception as e:
    print(f"  Error: {e}")

# ============ Method B: HParameterSet.HInsertPicture ============
print("\n=== Method B: HParameterSet.HInsertPicture ===")
try:
    abs_p = os.path.abspath(test_img)
    pset = hwp.HParameterSet.HInsertPicture
    print(f"  HInsertPicture object: {pset}")
    hwp.HAction.GetDefault("InsertPicture", pset.HSet)
    pset.FileName = abs_p
    pset.Embed = 1
    pset.Width = 5000
    pset.Height = 4000
    r = hwp.HAction.Execute("InsertPicture", pset.HSet)
    print(f"  Execute result: {r}")
except Exception as e:
    print(f"  Error: {e}")

# ============ Method C: Ctrl+V via keybd_event ============
print("\n=== Method C: keybd_event Ctrl+V ===")
try:
    # Move to right cell first
    hwp.HAction.Run("TableNextCell")
    time.sleep(0.2)

    img = Image.open(test_img).convert("RGB")
    img.thumbnail((600, 400))
    buf = io.BytesIO()
    img.save(buf, "BMP")
    bmp_data = buf.getvalue()[14:]
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_DIB, bmp_data)
    win32clipboard.CloseClipboard()
    time.sleep(0.3)

    # Find HWP window and force foreground
    hwp_hwnd = None
    def find_hwp(h, _):
        global hwp_hwnd
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t and ("Hwp" in t or "HWP" in t or "hwp" in t):
                hwp_hwnd = h
        return True
    win32gui.EnumWindows(find_hwp, None)

    if hwp_hwnd:
        # AttachThreadInput trick for foreground
        fg_thread = ctypes.windll.user32.GetWindowThreadProcessId(
            ctypes.windll.user32.GetForegroundWindow(), None)
        my_thread = ctypes.windll.kernel32.GetCurrentThreadId()
        ctypes.windll.user32.AttachThreadInput(my_thread, fg_thread, True)
        ctypes.windll.user32.SetForegroundWindow(hwp_hwnd)
        ctypes.windll.user32.AttachThreadInput(my_thread, fg_thread, False)
        time.sleep(0.5)

        # Ctrl+V
        VK_CONTROL = 0x11
        VK_V = 0x56
        ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 0, 0)
        ctypes.windll.user32.keybd_event(VK_V, 0, 0, 0)
        time.sleep(0.05)
        ctypes.windll.user32.keybd_event(VK_V, 0, 2, 0)
        ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 2, 0)
        time.sleep(1)
        print(f"  Ctrl+V sent to HWND={hwp_hwnd}")
    else:
        print("  HWP window not found!")
except Exception as e:
    print(f"  Error: {e}")

# Save
print(f"\nSaving to {OUT}")
ctrl.save_document(OUT)
ctrl.close_document(save=False, suppress_dialog=True)
pythoncom.CoUninitialize()
print("DONE")
