import ctypes
import tkinter as tk
from tkinter import simpledialog
from ctypes import wintypes
import sys
import threading
from PIL import Image, ImageDraw
import pystray
import json
import os

# =============================================================================
# 0. 설정 관리 (저장/불러오기/초기화)
# =============================================================================
CONFIG_FILE = 'config.json'

# 기본값 정의 (초기화 시 이 값으로 복구됨)
DEFAULT_SETTINGS = {
    'alpha': 0.4,
    'bg': 'black',
    'fg': 'red',   # 한글 모드일 때의 기본 글자색
    'width': 70,
    'height': 40,
    'x': 0,
    'y': 750,
    'font_size': 24,
    'anchor': 'e' # 정렬: 'w'(좌), 'center'(중앙), 'e'(우)
}

# 현재 설정을 담을 전역 변수
current_settings = DEFAULT_SETTINGS.copy()

def load_settings():
    """파일에서 설정 불러오기"""
    global current_settings
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved_settings = json.load(f)
                for key, value in saved_settings.items():
                    current_settings[key] = value
        except Exception as e:
            print(f"설정 로드 실패: {e}")

def save_settings():
    """현재 설정을 파일에 저장"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(current_settings, f, indent=4)
    except Exception as e:
        print(f"설정 저장 실패: {e}")

def reset_to_defaults(icon=None, item=None):
    """모든 설정을 기본값으로 되돌림"""
    global current_settings
    current_settings = DEFAULT_SETTINGS.copy()
    save_settings()
    root.after(0, apply_gui_settings)

# =============================================================================
# 1. 윈도우 API 설정
# =============================================================================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
imm32 = ctypes.windll.imm32
shell32 = ctypes.windll.shell32

GWL_EXSTYLE = -20
WS_EX_TRANSPARENT = 0x20
WS_EX_LAYERED = 0x80000

IME_CMODE_NATIVE = 0x0001
WM_IME_CONTROL = 0x0283
IMC_GETCONVERSIONMODE = 0x0001
IMC_GETOPENSTATUS = 0x0005

class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT)
    ]

if not hasattr(wintypes, 'LRESULT'):
    wintypes.LRESULT = ctypes.c_ssize_t

user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.AttachThreadInput.restype = wintypes.BOOL
user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
user32.GetGUIThreadInfo.restype = wintypes.BOOL
user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
user32.GetKeyboardLayout.restype = wintypes.HANDLE
user32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
user32.GetClassNameW.restype = ctypes.c_int
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.SendMessageW.restype = wintypes.LRESULT 
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
imm32.ImmGetContext.restype = wintypes.HANDLE
imm32.ImmGetContext.argtypes = [wintypes.HWND]
imm32.ImmGetConversionStatus.restype = wintypes.BOOL
imm32.ImmGetConversionStatus.argtypes = [wintypes.HANDLE, ctypes.POINTER(ctypes.c_ulong), ctypes.POINTER(ctypes.c_ulong)]
imm32.ImmGetDefaultIMEWnd.restype = wintypes.HWND
imm32.ImmGetDefaultIMEWnd.argtypes = [wintypes.HWND]
imm32.ImmGetOpenStatus.restype = wintypes.BOOL
imm32.ImmGetOpenStatus.argtypes = [wintypes.HANDLE]
shell32.IsUserAnAdmin.restype = wintypes.BOOL
shell32.IsUserAnAdmin.argtypes = []

_conversion = ctypes.c_ulong()
_sentence = ctypes.c_ulong()
_gui_info = GUITHREADINFO()
_gui_info.cbSize = ctypes.sizeof(GUITHREADINFO)

# =============================================================================
# 2. 관리자 권한 및 IME 로직
# =============================================================================
def is_admin():
    try: return shell32.IsUserAnAdmin()
    except: return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

def get_ime_mode_from_imc(hIMC):
    try:
        if not imm32.ImmGetOpenStatus(hIMC): return 'A'
        if imm32.ImmGetConversionStatus(hIMC, ctypes.byref(_conversion), ctypes.byref(_sentence)):
            return '한' if _conversion.value & IME_CMODE_NATIVE else 'A'
        return None
    except: return None

def get_ime_mode_from_hwnd(hwnd, debug=False):
    if not hwnd: return None
    hIMC = imm32.ImmGetContext(hwnd)
    if hIMC:
        mode = get_ime_mode_from_imc(hIMC)
        imm32.ImmReleaseContext(hwnd, hIMC)
        if mode: return mode
    ime_hwnd = imm32.ImmGetDefaultIMEWnd(hwnd)
    if ime_hwnd:
        hIMC = imm32.ImmGetContext(ime_hwnd)
        if hIMC:
            mode = get_ime_mode_from_imc(hIMC)
            imm32.ImmReleaseContext(ime_hwnd, hIMC)
            if mode: return mode
        try:
            if not user32.SendMessageW(ime_hwnd, WM_IME_CONTROL, IMC_GETOPENSTATUS, 0): return 'A'
            if user32.SendMessageW(ime_hwnd, WM_IME_CONTROL, IMC_GETCONVERSIONMODE, 0) & IME_CMODE_NATIVE: return '한'
            else: return 'A'
        except: pass
    return None

def get_input_mode(verbose=False):
    try:
        # 1. 현재 활성화된 창 핸들 가져오기
        fg_hwnd = user32.GetForegroundWindow()
        if not fg_hwnd: return None
        
        # 2. 내 프로그램(트레이 아이콘/오버레이)인 경우 검사 건너뛰기
        pid_buffer = wintypes.DWORD()
        fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(pid_buffer))
        current_pid = os.getpid()
        
        if pid_buffer.value == current_pid:
            return None

        # [최적화 1] Attach 없이 먼저 시도 (Fast Path)
        # 대부분의 프로그램은 스레드 연결 없이 윈도우 핸들만으로도 IME 상태 조회가 가능합니다.
        # 여기서 성공하면 위험한 AttachThreadInput을 호출하지 않아 렉이 발생하지 않습니다.
        mode_without_attach = get_ime_mode_from_hwnd(fg_hwnd, verbose)
        if mode_without_attach:
            return mode_without_attach

        cur_tid = kernel32.GetCurrentThreadId()
        
        # [최적화 2] 키보드 레이아웃이 한국어가 아니면 깊게 검사하지 않음
        # 현재 입력 언어가 영어나 다른 언어라면 굳이 내부 포커스까지 뒤질 필요가 없습니다.
        if (user32.GetKeyboardLayout(fg_tid) & 0xFFFF) != 0x0412: 
            return 'A'
        
        # [최적화 3] 최후의 수단으로만 Attach 수행 (Slow Path)
        # 위 방법이 모두 실패했을 때(예: 크롬 주소창 등 깊숙한 포커스)만 스레드를 연결합니다.
        attached = False
        if fg_tid != cur_tid:
            attached = user32.AttachThreadInput(cur_tid, fg_tid, True)
        
        try:
            # 스레드가 연결된 상태이므로 정밀하게 포커스(Focus)와 캐럿(Caret) 정보를 가져옵니다.
            targets = []
            if user32.GetGUIThreadInfo(fg_tid, ctypes.byref(_gui_info)):
                # 캐럿(커서 깜빡임)이 있는 곳 -> 포커스 된 곳 -> 활성 창 순서로 우선순위
                if _gui_info.hwndCaret: targets.append(_gui_info.hwndCaret)
                if _gui_info.hwndFocus: targets.append(_gui_info.hwndFocus)
                if _gui_info.hwndActive: targets.append(_gui_info.hwndActive)
            
            targets.append(fg_hwnd)
            
            for h in targets:
                m = get_ime_mode_from_hwnd(h, verbose)
                if m: return m
            return None
        finally:
            # [중요] 반드시 연결을 즉시 해제해야 입력 멈춤 현상을 방지할 수 있음
            if attached: user32.AttachThreadInput(cur_tid, fg_tid, False)
    except: return None

# =============================================================================
# 3. GUI 설정 (Tkinter)
# =============================================================================
load_settings()

root = tk.Tk()
root.overrideredirect(True)
root.attributes('-topmost', True)
root.attributes('-alpha', current_settings['alpha'])
root.configure(bg=current_settings['bg'])
root.configure(cursor='arrow')
root.geometry(f"{current_settings['width']}x{current_settings['height']}+{current_settings['x']}+{current_settings['y']}")

label = tk.Label(root, text='A', font=('Arial', current_settings['font_size'], 'bold'), 
                 fg=current_settings['fg'], bg=current_settings['bg'], 
                 anchor=current_settings['anchor'])
label.pack(expand=True, fill='both', padx=5)

hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED | WS_EX_TRANSPARENT)

root.withdraw()

current_text = ""
current_fg_color = ""

def apply_gui_settings():
    """변경된 설정을 GUI에 반영"""
    root.attributes('-alpha', current_settings['alpha'])
    root.configure(bg=current_settings['bg'])
    root.configure(cursor='arrow')
    # 폰트, 정렬, 배경색 적용 (글자색은 update_label에서 동적으로 처리)
    label.configure(bg=current_settings['bg'], 
                    font=('Arial', current_settings['font_size'], 'bold'),
                    anchor=current_settings['anchor'])
    root.geometry(f"{current_settings['width']}x{current_settings['height']}+{current_settings['x']}+{current_settings['y']}")

def update_label():
    global current_text, current_fg_color
    new_mode = get_input_mode()

    if new_mode == '한':
        root.deiconify()
        text = 'ㅎ'
        # 한글 모드일 때는 사용자가 설정한 fg 색상 사용
        text_color = current_settings['fg']
    elif new_mode == 'A':
        root.withdraw() # 취향에 따라 'A'일 때 숨김 유지
        text = 'A'
        
        # [스마트 색상 처리]
        light_backgrounds = ['white', 'lightblue', 'orange', 'orchid', 'yellow']
        if current_settings['bg'] in light_backgrounds:
            text_color = 'black'
        else:
            text_color = 'white'
    else:
        root.withdraw()
        text = '?'
        text_color = 'red'

    # 텍스트나 색상이 변경되었을 때만 업데이트
    if text != current_text or text_color != current_fg_color:
        label.config(text=text, fg=text_color)
        current_text = text
        current_fg_color = text_color

    root.after(500, update_label)

# =============================================================================
# 4. 시스템 트레이 및 설정 변경 로직
# =============================================================================

def ask_and_set_value(setting_key, prompt_msg, is_float=False):
    def _ask():
        root.attributes('-topmost', False)
        current_val = current_settings[setting_key]
        if is_float:
            val = simpledialog.askfloat("설정", f"{prompt_msg}\n(현재: {current_val})", parent=root, minvalue=0.1, maxvalue=1.0)
        else:
            val = simpledialog.askinteger("설정", f"{prompt_msg}\n(현재: {current_val})", parent=root)
            
        if val is not None:
            current_settings[setting_key] = val
            save_settings()
            apply_gui_settings()
        root.attributes('-topmost', True)
    root.after(0, _ask)

def set_background(color_name):
    def inner(icon, item):
        current_settings['bg'] = color_name
        save_settings()
        root.after(0, apply_gui_settings)
    return inner

def set_text_color(color_name):
    """[추가됨] 텍스트 색상 변경 함수"""
    def inner(icon, item):
        current_settings['fg'] = color_name
        save_settings()
        # apply_gui_settings에서는 label.config만 하므로 
        # 실제 색상 반영은 update_label 루프에서 처리됨
        root.after(0, apply_gui_settings) 
    return inner

def set_alignment(anchor_val):
    def inner(icon, item):
        current_settings['anchor'] = anchor_val
        save_settings()
        root.after(0, apply_gui_settings)
    return inner

def is_alignment_checked(anchor_val):
    def inner(item):
        return current_settings['anchor'] == anchor_val
    return inner

def exit_program(icon, item):
    icon.stop()
    root.quit()

def create_tray_image():
    width = 64
    height = 64
    image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((0, 0, width - 1, height - 1), fill='white', outline='black', width=15)
    return image

def setup_tray_icon():
    image = create_tray_image()
    menu = pystray.Menu(
        pystray.MenuItem('위치/크기 설정', pystray.Menu(
            pystray.MenuItem('너비 변경', lambda i, item: ask_and_set_value('width', '새 너비 입력:')),
            pystray.MenuItem('높이 변경', lambda i, item: ask_and_set_value('height', '새 높이 입력:')),
            pystray.MenuItem('X 좌표 변경', lambda i, item: ask_and_set_value('x', '새 X 좌표 입력:')),
            pystray.MenuItem('Y 좌표 변경', lambda i, item: ask_and_set_value('y', '새 Y 좌표 입력:')),
        )),
        pystray.MenuItem('텍스트 정렬', pystray.Menu(
            pystray.MenuItem('좌측 정렬', set_alignment('w'), checked=is_alignment_checked('w'), radio=True),
            pystray.MenuItem('중앙 정렬', set_alignment('center'), checked=is_alignment_checked('center'), radio=True),
            pystray.MenuItem('우측 정렬', set_alignment('e'), checked=is_alignment_checked('e'), radio=True),
        )),
        pystray.MenuItem('스타일 설정', pystray.Menu(
            pystray.MenuItem('글자 크기 변경', lambda i, item: ask_and_set_value('font_size', '새 글자 크기 입력:')),
            pystray.MenuItem('투명도 변경', lambda i, item: ask_and_set_value('alpha', '새 투명도 입력 (0.1~1.0):', is_float=True)),
            pystray.Menu.SEPARATOR,
            
            # [추가됨] 글자색 설정 메뉴
            pystray.MenuItem('글자색 설정 (한글)', pystray.Menu(
                pystray.MenuItem('글자: Red', set_text_color('red')),
                pystray.MenuItem('글자: Black', set_text_color('black')),
                pystray.MenuItem('글자: White', set_text_color('white')),
                pystray.MenuItem('글자: LightBlue', set_text_color('lightblue')),
                pystray.MenuItem('글자: Olive', set_text_color('olive')),
                pystray.MenuItem('글자: Orange', set_text_color('orange')),
                pystray.MenuItem('글자: Orchid', set_text_color('orchid')),
            )),
            pystray.Menu.SEPARATOR,
            # 배경색 설정 메뉴
            pystray.MenuItem('배경색 설정', pystray.Menu(
                pystray.MenuItem('배경: Black', set_background('black')),
                pystray.MenuItem('배경: White', set_background('white')),
                pystray.MenuItem('배경: Red', set_background('red')),
                pystray.MenuItem('배경: LightBlue', set_background('lightblue')),
                pystray.MenuItem('배경: Olive', set_background('olive')),
                pystray.MenuItem('배경: Orange', set_background('orange')),
                pystray.MenuItem('배경: Orchid', set_background('orchid')),
            )),
        )),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('설정 초기화 (Default)', reset_to_defaults),
        pystray.MenuItem('종료', exit_program)
    )
    icon = pystray.Icon('HangulIndicator', image, '한/영 표시기', menu)
    icon.run()

if __name__ == "__main__":
    tray_thread = threading.Thread(target=setup_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()

    update_label()
    root.mainloop()