import sys, time, io, glob, os


def _resolve_hwp_mcp_path() -> str:
    env_path = os.environ.get("HWP_MCP_PATH")
    if env_path:
        return env_path
    appdata = os.environ.get("APPDATA")
    if appdata:
        return os.path.join(appdata, "Cursor", "mcp-servers", "hwp-mcp")
    return r'C:\Users\user\AppData\Roaming\Cursor\mcp-servers\hwp-mcp'


sys.path.insert(0, _resolve_hwp_mcp_path())
import pythoncom, win32clipboard, win32con, win32gui, ctypes
from PIL import Image
from src.tools.hwp_controller import HwpController

TEMPLATE = r'D:\일일활동일지양식.hwp'
OUT = r'c:\Users\윤미란\Desktop\방과후센터\debug_photo_result.hwp'

imgs = [f for f in glob.glob(r'c:\Users\윤미란\Desktop\방과후센터\app\static\uploads\*.*')
        if os.path.splitext(f)[1].lower() in ('.jpg','.jpeg','.png','.gif','.webp')]
if not imgs:
    imgs = [f for f in glob.glob(r'c:\Users\윤미란\Desktop\*.jpg')]
if not imgs:
    imgs = [f for f in glob.glob(r'c:\Users\윤미란\Desktop\*.png')]

print(f"test images: {len(imgs)}")
for im in imgs[:3]:
    print(f"  {im}")

if not imgs:
    print("NO IMAGES FOUND - creating a test image")
    test_img = r'c:\Users\윤미란\Desktop\방과후센터\test_photo.png'
    Image.new("RGB", (200, 200), (255, 0, 0)).save(test_img)
    imgs = [test_img]

pythoncom.CoInitialize()
ctrl = HwpController()
ctrl.connect(visible=True, register_security_module=False)
ctrl.hwp.SetMessageBoxMode(0x000F0000)
ctrl.open_document(TEMPLATE)
time.sleep(2)
hwp = ctrl.hwp

# HWP HWND
hwp_hwnd = None
def find_hwp(h, _):
    global hwp_hwnd
    if win32gui.IsWindowVisible(h):
        t = win32gui.GetWindowText(h)
        if t and ("Hwp" in t or "HWP" in t or "hwp" in t):
            hwp_hwnd = h
            return False
    return True
win32gui.EnumWindows(find_hwp, None)
print(f"HWP HWND: {hwp_hwnd}")

def focus_hwp():
    if hwp_hwnd:
        try:
            ctypes.windll.user32.SetForegroundWindow(hwp_hwnd)
        except:
            pass
    time.sleep(0.3)

def paste_bmp(img_path):
    img = Image.open(img_path).convert("RGB")
    img.thumbnail((800, 600))
    buf = io.BytesIO()
    img.save(buf, "BMP")
    bmp_data = buf.getvalue()[14:]
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_DIB, bmp_data)
    win32clipboard.CloseClipboard()
    time.sleep(0.3)
    focus_hwp()
    r = hwp.HAction.Run("Paste")
    print(f"  Paste result={r}")
    time.sleep(0.5)
    if not r:
        focus_hwp()
        r2 = hwp.HAction.Run("Paste")
        print(f"  Paste retry result={r2}")
        time.sleep(0.5)

print("\n=== Step 1: find header ===")
found = ctrl.find_text("활  동  사  진")
p = hwp.GetPos()
print(f"found={found} ListId={p[0]} Para={p[1]} Pos={p[2]}")

print("\n=== Step 2: MoveDown ===")
focus_hwp()
hwp.HAction.Run("MoveDown")
time.sleep(0.2)
p2 = hwp.GetPos()
print(f"after MoveDown: ListId={p2[0]} Para={p2[1]} Pos={p2[2]}")

print("\n=== Step 3: Paste photo 1 (left cell) ===")
paste_bmp(imgs[0])

print("\n=== Step 4: Move to right cell ===")
# try multiple approaches
for action in ["TableNextCell", "TableRightCell"]:
    r = hwp.HAction.Run(action)
    p3 = hwp.GetPos()
    print(f"  {action}: result={r} ListId={p3[0]}")
    if r and p3[0] != p2[0]:
        print(f"  -> moved to different cell!")
        break

if len(imgs) >= 2:
    print("\n=== Step 5: Paste photo 2 (right cell) ===")
    paste_bmp(imgs[1])
else:
    print("\n=== Step 5: Paste same photo (right cell) ===")
    paste_bmp(imgs[0])

print(f"\n=== Save to {OUT} ===")
ctrl.save_document(OUT)
ctrl.close_document(save=False, suppress_dialog=True)
pythoncom.CoUninitialize()
print("DONE - check debug_photo_result.hwp")
