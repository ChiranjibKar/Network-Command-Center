"""
╔══════════════════════════════════════════════════════════╗
║             Network Command Center v2.0                  ║
║     Telemetry Department - Ground Station Toolkit        ║
╚══════════════════════════════════════════════════════════╝
Run as Administrator for full functionality.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
import subprocess
import socket
import uuid
import re
import threading
import ctypes
import sys
import os
import platform
import struct
import tempfile
import math


# ─────────────────────────── HIDE CONSOLE WINDOW ───────────────────────────
def hide_console():
    """Hide the black cmd window that appears behind the GUI."""
    if platform.system() == 'Windows':
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE = 0
        except:
            pass

hide_console()


# ─────────────────────────── ADMIN CHECK ───────────────────────────
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        try:
            # Get the absolute path to this script and quote it properly
            script = os.path.abspath(sys.argv[0])
            params = f'"{script}"'
            # Append any extra arguments (quoted individually)
            if len(sys.argv) > 1:
                extras = ' '.join(f'"{a}"' for a in sys.argv[1:])
                params += ' ' + extras
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, params, None, 1
            )
            sys.exit(0)
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            CustomDialog(root, "error", "Admin Required",
                         f"This app needs Administrator privileges.\n\n{e}")
            sys.exit(1)


# ─────────────────────────── ICON GENERATOR ───────────────────────────
def _make_icon_photoimage(root, size=32):
    """Generate a LAN network topology icon as a tk.PhotoImage.
    This works reliably for BOTH title bar AND taskbar on Windows."""
    W = H = size
    img = tk.PhotoImage(width=W, height=H)

    # Pre-compute pixel colors into a 2D grid
    # Colors in #RRGGBB
    C_TRANS   = None
    C_BG      = '#0d1117'
    C_ACCENT  = '#58a6ff'
    C_GREEN   = '#3fb950'
    C_WHITE   = '#f0f6fc'
    C_DIM     = '#30363d'
    C_DIMLINE = '#21262d'

    cx, cy = W // 2, H // 2
    # Scale factor for different sizes
    s = size / 32.0

    # Node positions (cx, cy, radius, color)
    nodes = [
        (cx, cy,           3.2 * s, C_WHITE),   # center
        (cx, 4 * s,        2.4 * s, C_GREEN),   # top
        ((W - 5) * s / s,  cy,      2.4 * s, C_GREEN),   # right
        (cx, (H - 6) * s / s,       2.4 * s, C_GREEN),   # bottom
        (4 * s,            cy,      2.4 * s, C_GREEN),   # left
    ]
    # Adjust for size
    nodes = [
        (cx,       cy,              3.0*s, C_WHITE),
        (cx,       int(4*s),        2.2*s, C_GREEN),
        (int(27*s/1), cy,           2.2*s, C_GREEN),
        (cx,       int(27*s/1),     2.2*s, C_GREEN),
        (int(4*s), cy,              2.2*s, C_GREEN),
    ]
    if size == 32:
        nodes = [
            (15, 15, 3.2, C_WHITE),
            (15, 4,  2.3, C_GREEN),
            (27, 15, 2.3, C_GREEN),
            (15, 26, 2.3, C_GREEN),
            (3,  15, 2.3, C_GREEN),
        ]
        sub_nodes = [(7, 7, 1.5, C_ACCENT), (24, 7, 1.5, C_ACCENT),
                     (24, 24, 1.5, C_ACCENT), (7, 24, 1.5, C_ACCENT)]
        main_lines = [(15,15,15,4), (15,15,27,15), (15,15,15,26), (15,15,3,15)]
        sub_lines = [(15,4,7,7), (15,4,24,7), (15,26,7,24), (15,26,24,24),
                     (3,15,7,7), (3,15,7,24), (27,15,24,7), (27,15,24,24)]
    else:
        # 16x16 simplified
        nodes = [(7,7,2, C_WHITE), (7,1,1.3, C_GREEN), (13,7,1.3, C_GREEN),
                 (7,13,1.3, C_GREEN), (1,7,1.3, C_GREEN)]
        sub_nodes = [(3,3,0.8,C_ACCENT),(11,3,0.8,C_ACCENT),
                     (11,11,0.8,C_ACCENT),(3,11,0.8,C_ACCENT)]
        main_lines = [(7,7,7,1),(7,7,13,7),(7,7,7,13),(7,7,1,7)]
        sub_lines = [(7,1,3,3),(7,1,11,3),(7,13,3,11),(7,13,11,11),
                     (1,7,3,3),(1,7,3,11),(13,7,11,3),(13,7,11,11)]

    # Build pixel grid
    grid = [[C_TRANS] * W for _ in range(H)]

    def set_px(x, y, color):
        if 0 <= x < W and 0 <= y < H and color is not None:
            grid[y][x] = color

    def draw_line_px(x0, y0, x1, y1, color):
        steps = max(abs(x1 - x0), abs(y1 - y0), 1) * 2
        for i in range(steps + 1):
            t = i / steps
            px = int(round(x0 + (x1 - x0) * t))
            py = int(round(y0 + (y1 - y0) * t))
            set_px(px, py, color)

    def draw_circle_px(cx, cy, r, color):
        ri = int(r) + 1
        for dy in range(-ri, ri + 1):
            for dx in range(-ri, ri + 1):
                if math.sqrt(dx*dx + dy*dy) <= r:
                    set_px(int(cx) + dx, int(cy) + dy, color)

    # Outer ring
    for y in range(H):
        for x in range(W):
            dist = math.sqrt((x - cx)**2 + (y - cy)**2)
            limit = 15 if size == 32 else 7.5
            if limit - 1 <= dist <= limit:
                grid[y][x] = C_DIM

    # Draw lines
    for x0, y0, x1, y1 in sub_lines:
        draw_line_px(x0, y0, x1, y1, C_DIMLINE)
    for x0, y0, x1, y1 in main_lines:
        draw_line_px(x0, y0, x1, y1, C_ACCENT)

    # Draw nodes
    for nx, ny, r, c in sub_nodes:
        draw_circle_px(nx, ny, r, c)
    for nx, ny, r, c in nodes:
        draw_circle_px(nx, ny, r, c)

    # Write pixels to PhotoImage using row-based put (MUCH faster than per-pixel)
    row_data = []
    for y in range(H):
        row = ""
        for x in range(W):
            c = grid[y][x]
            if c is None:
                row += " {} "
            else:
                row += f" {c} "
        row_data.append(row.strip())

    # PhotoImage.put() accepts a space-separated row format
    for y, row_str in enumerate(row_data):
        colors_in_row = []
        for x in range(W):
            c = grid[y][x]
            if c is None:
                colors_in_row.append(C_BG)  # fill transparent with bg
            else:
                colors_in_row.append(c)
        img.put("{" + " ".join(colors_in_row) + "}", to=(0, y))

    return img


def _make_icon_ico_file():
    """Generate .ico file as fallback for iconbitmap."""
    W, H = 32, 32
    T  = (0, 0, 0, 0)
    BG = (23, 17, 13, 255)
    AB = (255, 166, 88, 255)
    GR = (80, 185, 63, 255)
    WH = (252, 246, 240, 255)
    DM = (60, 54, 48, 255)

    pixels = [[BG] * W for _ in range(H)]

    def set_px(x, y, color):
        if 0 <= x < W and 0 <= y < H:
            pixels[y][x] = color

    def draw_line_px(x0, y0, x1, y1, color):
        steps = max(abs(x1 - x0), abs(y1 - y0), 1) * 2
        for i in range(steps + 1):
            t = i / steps
            px = int(round(x0 + (x1 - x0) * t))
            py = int(round(y0 + (y1 - y0) * t))
            set_px(px, py, color)

    def draw_circle_px(cx, cy, r, color):
        ri = int(r) + 1
        for dy in range(-ri, ri + 1):
            for dx in range(-ri, ri + 1):
                if math.sqrt(dx*dx + dy*dy) <= r:
                    set_px(int(cx) + dx, int(cy) + dy, color)

    # Outer ring
    cx, cy = 15, 15
    for y in range(H):
        for x in range(W):
            dist = math.sqrt((x - cx)**2 + (y - cy)**2)
            if 14 <= dist <= 15:
                pixels[y][x] = DM

    # Lines
    for x0,y0,x1,y1 in [(15,4,7,7),(15,4,24,7),(15,26,7,24),(15,26,24,24),
                          (3,15,7,7),(3,15,7,24),(27,15,24,7),(27,15,24,24)]:
        draw_line_px(x0,y0,x1,y1, DM)
    for x0,y0,x1,y1 in [(15,15,15,4),(15,15,27,15),(15,15,15,26),(15,15,3,15)]:
        draw_line_px(x0,y0,x1,y1, AB)

    # Nodes
    for nx,ny,r,c in [(7,7,1.5,AB),(24,7,1.5,AB),(24,24,1.5,AB),(7,24,1.5,AB)]:
        draw_circle_px(nx,ny,r,c)
    for nx,ny,r,c in [(15,15,3.2,WH),(15,4,2.3,GR),(27,15,2.3,GR),(15,26,2.3,GR),(3,15,2.3,GR)]:
        draw_circle_px(nx,ny,r,c)

    # Build ICO binary
    bmp_data = bytearray()
    for y in range(H - 1, -1, -1):
        for x in range(W):
            b, g, r, a = pixels[y][x]
            bmp_data.extend([b, g, r, a])
    and_mask = bytearray(W * H // 8)

    ico = bytearray()
    ico.extend(struct.pack('<HHH', 0, 1, 1))
    bmp_size = 40 + len(bmp_data) + len(and_mask)
    ico.extend(struct.pack('<BBBBHHIH', W, H, 0, 0, 1, 32, bmp_size, 22))
    ico.extend(struct.pack('<IiiHHIIiiII', 40, W, H*2, 1, 32, 0,
                            len(bmp_data)+len(and_mask), 0, 0, 0, 0))
    ico.extend(bmp_data)
    ico.extend(and_mask)

    ico_path = os.path.join(tempfile.gettempdir(), 'ncc_icon.ico')
    with open(ico_path, 'wb') as f:
        f.write(ico)
    return ico_path


# ─────────────────────────── HELPERS ───────────────────────────
def get_hostname():
    return socket.gethostname()

def get_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "N/A"

def get_mac():
    mac = ':'.join(f'{(uuid.getnode() >> i) & 0xff:02x}' for i in range(40, -1, -8))
    return mac.upper()

def run_cmd(cmd, timeout=15):
    try:
        # CREATE_NO_WINDOW prevents cmd flash, but only exists on Windows
        kwargs = dict(capture_output=True, text=True, timeout=timeout, shell=True)
        if platform.system() == 'Windows':
            kwargs['creationflags'] = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
        result = subprocess.run(cmd, **kwargs)
        return result.stdout + result.stderr
    except subprocess.TimeoutExpired:
        return "[Timed out]"
    except Exception as e:
        return f"[Error] {e}"

def get_adapters():
    adapters = []
    output = run_cmd('netsh interface ipv4 show addresses')
    blocks = output.strip().split('\n\n')
    for block in blocks:
        lines = block.strip().split('\n')
        adapter = {}
        for line in lines:
            line = line.strip()
            if line.startswith('Configuration for interface'):
                name = line.replace('Configuration for interface', '').strip().strip('"')
                adapter['name'] = name
            elif 'IP Address' in line or 'IP address' in line:
                parts = line.split(':')
                if len(parts) >= 2:
                    adapter['ip'] = parts[-1].strip()
            elif 'Subnet' in line or 'subnet' in line:
                parts = line.split(':')
                if len(parts) >= 2:
                    adapter['subnet'] = parts[-1].strip()
            elif 'Default Gateway' in line or 'default gateway' in line:
                parts = line.split(':')
                if len(parts) >= 2:
                    adapter['gateway'] = parts[-1].strip()
        if 'name' in adapter:
            status_out = run_cmd(f'netsh interface show interface name="{adapter["name"]}"')
            adapter['enabled'] = 'Connected' in status_out or 'Enabled' in status_out
            adapters.append(adapter)
    return adapters

def get_adapter_mac(adapter_name):
    output = run_cmd(f'getmac /v /fo csv')
    for line in output.split('\n'):
        if adapter_name.lower() in line.lower():
            parts = line.split(',')
            for part in parts:
                cleaned = part.strip().strip('"')
                if re.match(r'^([0-9A-Fa-f]{2}[-:]){5}[0-9A-Fa-f]{2}$', cleaned):
                    return cleaned
    return "N/A"


# ─────────────────────────── COLOR THEME ───────────────────────────
COLORS = {
    'bg_dark':       '#0d1117',
    'bg_card':       '#161b22',
    'bg_input':      '#21262d',
    'bg_overlay':    '#010409',
    'border':        '#30363d',
    'accent':        '#58a6ff',
    'accent_hover':  '#79c0ff',
    'green':         '#3fb950',
    'green_dark':    '#238636',
    'red':           '#f85149',
    'red_dark':      '#da3633',
    'orange':        '#d29922',
    'yellow':        '#e3b341',
    'purple':        '#bc8cff',
    'text':          '#f0f6fc',
    'text_dim':      '#8b949e',
    'text_muted':    '#484f58',
}

DIALOG_ICONS = {
    'info':    ('i',  COLORS['accent'], COLORS['accent']),
    'success': ('✓', COLORS['green'],  COLORS['green']),
    'warning': ('!',  COLORS['orange'], COLORS['orange']),
    'error':   ('✕',  COLORS['red'],    COLORS['red']),
    'confirm': ('?',  COLORS['purple'], COLORS['purple']),
}

# (Icon generation functions are above — _make_icon_photoimage and _make_icon_ico_file)


# ─────────────────────────── CUSTOM DIALOG ───────────────────────────
class CustomDialog(tk.Toplevel):
    """Dark-themed dialog replacing tkinter messagebox."""

    def __init__(self, parent, dialog_type, title, message,
                 yes_no=False, callback=None):
        super().__init__(parent)
        self.result = None
        self.callback = callback

        self.title(title)
        self.configure(bg=COLORS['bg_overlay'])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        try:
            ico = _make_icon_photoimage(self, 32)
            self.wm_iconphoto(True, ico)
        except:
            pass

        self.protocol("WM_DELETE_WINDOW", lambda: self._close(False))

        icon_sym, icon_fg, border_col = DIALOG_ICONS.get(dialog_type, DIALOG_ICONS['info'])

        # ── Main container ──
        outer = tk.Frame(self, bg=border_col, padx=2, pady=2)
        outer.pack(fill='both', expand=True, padx=6, pady=6)

        inner = tk.Frame(outer, bg=COLORS['bg_dark'])
        inner.pack(fill='both', expand=True)

        # Top accent strip
        tk.Frame(inner, bg=border_col, height=3).pack(fill='x')

        # ── Content: icon on left, text on right ──
        content = tk.Frame(inner, bg=COLORS['bg_dark'], padx=18, pady=16)
        content.pack(fill='both', expand=True)

        # Horizontal layout: icon | text
        body = tk.Frame(content, bg=COLORS['bg_dark'])
        body.pack(fill='both', expand=True)

        # Icon canvas on left
        icon_canvas = tk.Canvas(body, width=48, height=48,
                                 bg=COLORS['bg_dark'], highlightthickness=0)
        icon_canvas.pack(side='left', anchor='n', padx=(0, 14), pady=(2, 0))

        if dialog_type == 'warning':
            icon_canvas.create_polygon(24, 4, 46, 44, 2, 44,
                                        fill=COLORS['bg_card'], outline=border_col, width=2)
            icon_canvas.create_line(24, 18, 24, 30, fill=icon_fg, width=3, capstyle='round')
            icon_canvas.create_oval(22, 34, 26, 38, fill=icon_fg, outline='')
        elif dialog_type == 'error':
            icon_canvas.create_oval(2, 2, 46, 46, fill=COLORS['bg_card'], outline=border_col, width=2)
            icon_canvas.create_line(16, 16, 32, 32, fill=icon_fg, width=3, capstyle='round')
            icon_canvas.create_line(32, 16, 16, 32, fill=icon_fg, width=3, capstyle='round')
        elif dialog_type == 'success':
            icon_canvas.create_oval(2, 2, 46, 46, fill=COLORS['bg_card'], outline=border_col, width=2)
            icon_canvas.create_line(14, 24, 21, 32, 34, 16,
                                     fill=icon_fg, width=3, capstyle='round', joinstyle='round')
        elif dialog_type == 'confirm':
            icon_canvas.create_oval(2, 2, 46, 46, fill=COLORS['bg_card'], outline=border_col, width=2)
            icon_canvas.create_arc(17, 12, 31, 28, start=0, extent=180,
                                    style='arc', outline=icon_fg, width=3)
            icon_canvas.create_line(24, 26, 24, 30, fill=icon_fg, width=3, capstyle='round')
            icon_canvas.create_oval(22, 34, 26, 38, fill=icon_fg, outline='')
        else:
            icon_canvas.create_oval(2, 2, 46, 46, fill=COLORS['bg_card'], outline=border_col, width=2)
            icon_canvas.create_oval(22, 13, 26, 17, fill=icon_fg, outline='')
            icon_canvas.create_line(24, 22, 24, 35, fill=icon_fg, width=3, capstyle='round')

        # Text column on right
        text_col = tk.Frame(body, bg=COLORS['bg_dark'])
        text_col.pack(side='left', fill='both', expand=True)

        tk.Label(text_col, text=title, bg=COLORS['bg_dark'], fg=COLORS['text'],
                 font=('Segoe UI', 12, 'bold'), anchor='w').pack(fill='x', pady=(0, 6))

        msg_label = tk.Label(text_col, text=message, bg=COLORS['bg_dark'], fg=COLORS['text_dim'],
                              font=('Segoe UI', 10), anchor='nw', justify='left',
                              wraplength=360)
        msg_label.pack(fill='both', expand=True)

        # ── Buttons ──
        btn_frame = tk.Frame(inner, bg=COLORS['bg_dark'], padx=18, pady=12)
        btn_frame.pack(fill='x', side='bottom')

        if yes_no:
            tk.Button(btn_frame, text="  Cancel  ",
                      bg=COLORS['bg_card'], fg=COLORS['text_dim'],
                      font=('Segoe UI', 10), relief='flat',
                      padx=18, pady=5, cursor='hand2',
                      activebackground=COLORS['bg_input'], activeforeground=COLORS['text'],
                      command=lambda: self._close(False)).pack(side='right', padx=(6, 0))

            yes_btn = tk.Button(btn_frame, text="  Confirm  ",
                                bg=border_col, fg='white',
                                font=('Segoe UI', 10, 'bold'), relief='flat',
                                padx=18, pady=5, cursor='hand2',
                                activebackground=icon_fg, activeforeground='white',
                                command=lambda: self._close(True))
            yes_btn.pack(side='right')
            yes_btn.focus_set()
            self.bind('<Return>', lambda e: self._close(True))
            self.bind('<Escape>', lambda e: self._close(False))
        else:
            ok_btn = tk.Button(btn_frame, text="    OK    ",
                               bg=border_col, fg='white',
                               font=('Segoe UI', 10, 'bold'), relief='flat',
                               padx=20, pady=5, cursor='hand2',
                               activebackground=icon_fg, activeforeground='white',
                               command=lambda: self._close(True))
            ok_btn.pack(side='right')
            ok_btn.focus_set()
            self.bind('<Return>', lambda e: self._close(True))
            self.bind('<Escape>', lambda e: self._close(True))

        # ── Auto-size the window after content is laid out ──
        self.update_idletasks()
        req_w = max(420, inner.winfo_reqwidth() + 28)
        req_h = max(180, inner.winfo_reqheight() + 28)
        px = parent.winfo_rootx() + parent.winfo_width() // 2 - req_w // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - req_h // 2
        self.geometry(f'{req_w}x{req_h}+{px}+{py}')

        self.wait_window()

    def _close(self, value):
        self.result = value
        if self.callback:
            self.callback(value)
        self.destroy()


# ─────────────────────────── TOAST NOTIFICATION ───────────────────────────
class ToastNotification:
    """Slide-in notification that auto-dismisses."""

    def __init__(self, parent, message, toast_type='info', duration=3000):
        self.parent = parent
        icon_sym, icon_fg, border_col = DIALOG_ICONS.get(toast_type, DIALOG_ICONS['info'])

        self.frame = tk.Frame(parent, bg=border_col, padx=2, pady=2)
        inner = tk.Frame(self.frame, bg=COLORS['bg_card'], padx=12, pady=8)
        inner.pack(fill='both', expand=True)

        row = tk.Frame(inner, bg=COLORS['bg_card'])
        row.pack(fill='x')

        # Mini canvas icon
        ic = tk.Canvas(row, width=24, height=24, bg=COLORS['bg_card'], highlightthickness=0)
        ic.pack(side='left', padx=(0, 8))
        if toast_type == 'success':
            ic.create_oval(1, 1, 23, 23, fill='', outline=icon_fg, width=2)
            ic.create_line(7, 12, 10, 16, 17, 8, fill=icon_fg, width=2,
                           capstyle='round', joinstyle='round')
        elif toast_type == 'error':
            ic.create_oval(1, 1, 23, 23, fill='', outline=icon_fg, width=2)
            ic.create_line(8, 8, 16, 16, fill=icon_fg, width=2, capstyle='round')
            ic.create_line(16, 8, 8, 16, fill=icon_fg, width=2, capstyle='round')
        elif toast_type == 'warning':
            ic.create_polygon(12, 2, 23, 22, 1, 22, fill='', outline=icon_fg, width=2)
            ic.create_line(12, 9, 12, 15, fill=icon_fg, width=2, capstyle='round')
            ic.create_oval(11, 17, 13, 19, fill=icon_fg, outline='')
        else:
            ic.create_oval(1, 1, 23, 23, fill='', outline=icon_fg, width=2)
            ic.create_oval(11, 6, 13, 8, fill=icon_fg, outline='')
            ic.create_line(12, 11, 12, 18, fill=icon_fg, width=2, capstyle='round')

        tk.Label(row, text=message, bg=COLORS['bg_card'], fg=COLORS['text'],
                 font=('Segoe UI', 10), anchor='w').pack(side='left', fill='x', expand=True)

        close_btn = tk.Label(row, text="✕", bg=COLORS['bg_card'], fg=COLORS['text_muted'],
                              font=('Segoe UI', 10), cursor='hand2')
        close_btn.pack(side='right', padx=(8, 0))
        close_btn.bind('<Button-1>', lambda e: self._dismiss())

        self.frame.place(relx=1.0, rely=0.0, anchor='ne', x=-15, y=10)
        parent.after(duration, self._dismiss)

    def _dismiss(self):
        try:
            self.frame.destroy()
        except:
            pass


# ═══════════════════════════ MAIN APP ═══════════════════════════
class NetworkCommandCenter(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Network Command Center v2.0")
        self.geometry("1100x750")
        self.minsize(950, 650)
        self.configure(bg=COLORS['bg_dark'])

        # Set custom icon — use wm_iconphoto (reliable for BOTH title bar AND taskbar)
        try:
            # PhotoImage for wm_iconphoto — this is the reliable method
            self._icon_img_32 = _make_icon_photoimage(self, 32)
            self._icon_img_16 = _make_icon_photoimage(self, 16)
            self.wm_iconphoto(True, self._icon_img_32, self._icon_img_16)
        except Exception:
            # Fallback: try .ico file with iconbitmap
            try:
                ico_path = _make_icon_ico_file()
                self.iconbitmap(default=ico_path)
            except:
                pass
        # Set Windows taskbar grouping ID
        if platform.system() == 'Windows':
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    'NetworkCommandCenter.v2.0')
            except:
                pass

        self.style = ttk.Style()
        self._configure_styles()
        self._build_ui()

    # ── Dialog wrappers ──
    def show_info(self, title, msg):
        CustomDialog(self, 'info', title, msg)

    def show_success(self, title, msg):
        CustomDialog(self, 'success', title, msg)

    def show_warning(self, title, msg):
        CustomDialog(self, 'warning', title, msg)

    def show_error(self, title, msg):
        CustomDialog(self, 'error', title, msg)

    def ask_confirm(self, title, msg):
        dlg = CustomDialog(self, 'confirm', title, msg, yes_no=True)
        return dlg.result

    def toast(self, msg, toast_type='info', duration=3000):
        ToastNotification(self, msg, toast_type, duration)

    def _add_copy_menu(self, widget):
        """Add right-click 'Copy' context menu to a readonly widget."""
        menu = tk.Menu(widget, tearoff=0, bg=COLORS['bg_card'], fg=COLORS['text'],
                       activebackground=COLORS['accent'], activeforeground='white',
                       font=('Segoe UI', 9), relief='flat', bd=1)
        menu.add_command(label="  Copy  ", command=lambda: self._copy_widget_text(widget))

        def show_menu(e):
            menu.tk_popup(e.x_root, e.y_root)

        widget.bind('<Button-3>', show_menu)   # Right-click
        widget.bind('<Control-c>', lambda e: self._copy_widget_text(widget))

    def _copy_widget_text(self, widget):
        """Copy widget text to clipboard."""
        try:
            text = widget.get()
            self.clipboard_clear()
            self.clipboard_append(text)
            self.toast("Copied to clipboard!", 'success', 1500)
        except:
            pass

    def _typewriter_tick(self):
        """Animate the developer credit one character at a time."""
        if self._typewriter_idx <= len(self._typewriter_text):
            displayed = self._typewriter_text[:self._typewriter_idx]
            # Show blinking cursor during typing
            cursor = "▌" if self._typewriter_idx < len(self._typewriter_text) else ""
            self.dev_label.config(text=displayed + cursor)
            self._typewriter_idx += 1
            # Vary speed slightly for natural feel
            delay = 65
            char = self._typewriter_text[self._typewriter_idx - 1] if self._typewriter_idx > 0 and self._typewriter_idx <= len(self._typewriter_text) else ''
            if char == ' ':
                delay = 40
            self.after(delay, self._typewriter_tick)
        else:
            # Typing done — blink cursor 3 times then settle
            self._blink_count = 0
            self._typewriter_blink()

    def _typewriter_blink(self):
        """Blink the cursor a few times after typing finishes, then show final text."""
        if self._blink_count < 6:
            if self._blink_count % 2 == 0:
                self.dev_label.config(text=self._typewriter_text + "▌")
            else:
                self.dev_label.config(text=self._typewriter_text)
            self._blink_count += 1
            self.after(350, self._typewriter_blink)
        else:
            # Final state — settled, slight color upgrade
            self.dev_label.config(text=self._typewriter_text,
                                   fg=COLORS['text_dim'])

    def _configure_styles(self):
        self.style.theme_use('clam')
        self.style.configure('TNotebook', background=COLORS['bg_dark'],
                             borderwidth=0, padding=0)
        self.style.configure('TNotebook.Tab',
                             background=COLORS['bg_card'],
                             foreground=COLORS['text_dim'],
                             padding=[16, 8],
                             font=('Segoe UI', 10, 'bold'))
        self.style.map('TNotebook.Tab',
                       background=[('selected', COLORS['bg_dark'])],
                       foreground=[('selected', COLORS['accent'])])
        self.style.configure('TFrame', background=COLORS['bg_dark'])
        self.style.configure('TLabel',
                             background=COLORS['bg_dark'],
                             foreground=COLORS['text'],
                             font=('Segoe UI', 10))
        self.style.configure('TEntry',
                             fieldbackground=COLORS['bg_input'],
                             foreground=COLORS['text'],
                             insertcolor=COLORS['text'],
                             borderwidth=1)
        self.style.configure('Treeview',
                             background=COLORS['bg_card'],
                             foreground=COLORS['text'],
                             fieldbackground=COLORS['bg_card'],
                             font=('Consolas', 9),
                             rowheight=24)
        self.style.configure('Treeview.Heading',
                             background=COLORS['bg_input'],
                             foreground=COLORS['accent'],
                             font=('Segoe UI', 9, 'bold'))
        self.style.map('Treeview',
                       background=[('selected', COLORS['accent'])],
                       foreground=[('selected', 'white')])

    def _build_ui(self):
        # ── TOP BAR ──
        top_frame = tk.Frame(self, bg=COLORS['bg_card'], pady=10, padx=15)
        top_frame.pack(fill='x', padx=10, pady=(10, 5))

        # Title + developer credit area
        title_area = tk.Frame(top_frame, bg=COLORS['bg_card'])
        title_area.pack(side='left')

        tk.Label(title_area, text="  NETWORK COMMAND CENTER",
                 bg=COLORS['bg_card'], fg=COLORS['accent'],
                 font=('Segoe UI', 14, 'bold')).pack(anchor='w')

        # Typewriter label — starts empty, fills in on launch
        self.dev_label = tk.Label(title_area, text="",
                                   bg=COLORS['bg_card'], fg=COLORS['text_muted'],
                                   font=('Consolas', 9))
        self.dev_label.pack(anchor='w', padx=(6, 0))
        self._typewriter_text = " Developed by Chiranjib Kar"
        self._typewriter_idx = 0
        self.after(800, self._typewriter_tick)  # start after a brief delay

        info_frame = tk.Frame(top_frame, bg=COLORS['bg_card'])
        info_frame.pack(side='right')

        for label, value in [("HOST", get_hostname()), ("IP", get_ip()), ("MAC", get_mac())]:
            f = tk.Frame(info_frame, bg=COLORS['bg_card'])
            f.pack(side='left', padx=12)
            tk.Label(f, text=label, bg=COLORS['bg_card'],
                     fg=COLORS['text_dim'], font=('Segoe UI', 8, 'bold')).pack()
            # Copyable readonly entry instead of plain label
            val_entry = tk.Entry(f, bg=COLORS['bg_card'], fg=COLORS['text'],
                                  font=('Consolas', 10, 'bold'), relief='flat',
                                  readonlybackground=COLORS['bg_card'],
                                  highlightthickness=0, bd=0, width=len(value) + 1,
                                  justify='center', cursor='arrow')
            val_entry.insert(0, value)
            val_entry.config(state='readonly')
            val_entry.pack()
            # Right-click copy menu
            self._add_copy_menu(val_entry)

        admin_ok = is_admin()
        badge_bg = COLORS['green_dark'] if admin_ok else COLORS['red_dark']
        badge_text = " ADMIN " if admin_ok else " NO ADMIN "
        badge_border = COLORS['green'] if admin_ok else COLORS['red']
        bw = tk.Frame(info_frame, bg=badge_border, padx=1, pady=1)
        bw.pack(side='left', padx=(20, 0))
        tk.Label(bw, text=badge_text, bg=badge_bg, fg='white',
                 font=('Segoe UI', 8, 'bold'), padx=6, pady=1).pack()

        # ── TABS ──
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)

        tabs = [
            ("  IPv4 Config  ", self._build_ipv4_tab),
            ("  Adapters  ",    self._build_adapters_tab),
            ("  Ping  ",        self._build_ping_tab),
            ("  Connections  ", self._build_netstat_tab),
            ("  Port Lookup  ", self._build_port_lookup_tab),
        ]
        for text, builder in tabs:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=text)
            builder(frame)

    # ═══════════════════ TAB 1: IPv4 CONFIG ═══════════════════
    def _build_ipv4_tab(self, parent):
        container = tk.Frame(parent, bg=COLORS['bg_dark'])
        container.pack(fill='both', expand=True, padx=15, pady=10)

        top = tk.Frame(container, bg=COLORS['bg_dark'])
        top.pack(fill='x', pady=(0, 10))

        tk.Label(top, text="Select Network Adapter:", bg=COLORS['bg_dark'],
                 fg=COLORS['text'], font=('Segoe UI', 10)).pack(side='left')

        self.ipv4_adapter_var = tk.StringVar()
        self.ipv4_combo = ttk.Combobox(top, textvariable=self.ipv4_adapter_var,
                                        state='readonly', width=40, font=('Consolas', 10))
        self.ipv4_combo.pack(side='left', padx=10)
        self.ipv4_combo.bind('<<ComboboxSelected>>', self._on_adapter_select)

        tk.Button(top, text="  Refresh  ", bg=COLORS['accent'], fg='white',
                  font=('Segoe UI', 9, 'bold'), relief='flat', padx=12, pady=4,
                  cursor='hand2', command=self._refresh_ipv4).pack(side='left', padx=5)

        # Warning banner
        self.warning_frame = tk.Frame(container, bg=COLORS['bg_dark'])
        self.warning_inner = tk.Frame(self.warning_frame, bg=COLORS['bg_card'],
                                       highlightthickness=2,
                                       highlightbackground=COLORS['orange'], padx=12, pady=8)
        self.warning_inner.pack(fill='x')

        warn_header = tk.Frame(self.warning_inner, bg=COLORS['bg_card'])
        warn_header.pack(fill='x')
        self.warn_icon_cv = tk.Canvas(warn_header, width=28, height=28,
                                       bg=COLORS['bg_card'], highlightthickness=0)
        self.warn_icon_cv.pack(side='left', padx=(0, 8))
        self.warn_icon_cv.create_polygon(14, 2, 26, 26, 2, 26,
                                          fill=COLORS['orange'], outline=COLORS['yellow'])
        self.warn_icon_cv.create_text(14, 18, text="!", fill='black',
                                       font=('Segoe UI', 11, 'bold'))
        tk.Label(warn_header, text="IP CONFLICT DETECTED", bg=COLORS['bg_card'],
                 fg=COLORS['orange'], font=('Segoe UI', 10, 'bold')).pack(side='left')

        self.warning_label = tk.Label(self.warning_inner, text="", bg=COLORS['bg_card'],
                                       fg=COLORS['text'], font=('Consolas', 9),
                                       anchor='w', justify='left')
        self.warning_label.pack(fill='x', pady=(4, 0))

        # Current values
        cur = tk.Frame(container, bg=COLORS['bg_card'], padx=15, pady=12,
                        highlightthickness=1, highlightbackground=COLORS['border'])
        cur.pack(fill='x', pady=(0, 10))

        tk.Label(cur, text="CURRENT SETTINGS", bg=COLORS['bg_card'], fg=COLORS['accent'],
                 font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, columnspan=3,
                                                       sticky='w', pady=(0, 8))

        ro_cfg = dict(bg=COLORS['bg_card'], fg=COLORS['text'], font=('Consolas', 11),
                      relief='flat', readonlybackground=COLORS['bg_card'],
                      highlightthickness=0, bd=0, width=20, cursor='arrow')
        ro_dim_cfg = {**ro_cfg}
        ro_dim_cfg['fg'] = COLORS['text_dim']
        ro_dim_cfg['font'] = ('Consolas', 10)

        for i, (lbl_text, attr_name, cfg) in enumerate([
            ("IP:",      'cur_ip_entry',  ro_cfg),
            ("Subnet:",  'cur_sub_entry', ro_cfg),
            ("Gateway:", 'cur_gw_entry',  ro_cfg),
            ("MAC:",     'cur_mac_entry', ro_dim_cfg),
        ]):
            tk.Label(cur, text=lbl_text, bg=COLORS['bg_card'], fg=COLORS['text_dim'],
                     font=('Consolas', 10)).grid(row=i+1, column=0, sticky='w', pady=2, padx=(0, 6))
            entry = tk.Entry(cur, **cfg)
            entry.insert(0, "--")
            entry.config(state='readonly')
            entry.grid(row=i+1, column=1, sticky='w', pady=2)
            setattr(self, attr_name, entry)
            self._add_copy_menu(entry)

        # Edit fields
        edit = tk.Frame(container, bg=COLORS['bg_card'], padx=15, pady=12,
                         highlightthickness=1, highlightbackground=COLORS['border'])
        edit.pack(fill='x', pady=(0, 10))

        tk.Label(edit, text="CHANGE IPv4 SETTINGS", bg=COLORS['bg_card'], fg=COLORS['accent'],
                 font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, columnspan=2,
                                                       sticky='w', pady=(0, 8))

        entry_cfg = dict(bg=COLORS['bg_input'], fg=COLORS['text'], insertbackground=COLORS['text'],
                         font=('Consolas', 11), relief='flat', highlightthickness=1,
                         highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'], width=20)
        self.ip_entry     = tk.Entry(edit, **entry_cfg)
        self.subnet_entry = tk.Entry(edit, **entry_cfg)
        self.gw_entry     = tk.Entry(edit, **entry_cfg)

        for i, (lbl, ent) in enumerate(zip(
                ["New IP Address:", "New Subnet Mask:", "New Default Gateway:"],
                [self.ip_entry, self.subnet_entry, self.gw_entry])):
            tk.Label(edit, text=lbl, bg=COLORS['bg_card'], fg=COLORS['text'],
                     font=('Segoe UI', 10)).grid(row=i+1, column=0, sticky='w', pady=4, padx=(0, 10))
            ent.grid(row=i+1, column=1, sticky='w', pady=4)

        btn_frame = tk.Frame(edit, bg=COLORS['bg_card'])
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(12, 0), sticky='w')

        tk.Button(btn_frame, text="  Apply Static IP  ", bg=COLORS['green_dark'], fg='white',
                  font=('Segoe UI', 10, 'bold'), relief='flat', padx=16, pady=6,
                  cursor='hand2', command=self._apply_static_ip).pack(side='left', padx=(0, 10))

        tk.Button(btn_frame, text="  Switch to DHCP  ", bg=COLORS['accent'], fg='white',
                  font=('Segoe UI', 10, 'bold'), relief='flat', padx=16, pady=6,
                  cursor='hand2', command=self._switch_to_dhcp).pack(side='left')

        self.ipv4_status = tk.Label(container, text="", bg=COLORS['bg_dark'],
                                     fg=COLORS['green'], font=('Segoe UI', 9))
        self.ipv4_status.pack(fill='x')

        self._refresh_ipv4()

    def _refresh_ipv4(self):
        def do():
            self.adapters_cache = get_adapters()
            names = [a['name'] for a in self.adapters_cache]
            self.ipv4_combo['values'] = names
            if names:
                self.ipv4_combo.current(0)
                self._on_adapter_select(None)
            self._check_duplicate_ips()
        threading.Thread(target=do, daemon=True).start()

    def _on_adapter_select(self, event):
        name = self.ipv4_adapter_var.get()
        for a in self.adapters_cache:
            if a['name'] == name:
                mac = get_adapter_mac(name)
                for entry, val in [
                    (self.cur_ip_entry,  a.get('ip', 'N/A')),
                    (self.cur_sub_entry, a.get('subnet', 'N/A')),
                    (self.cur_gw_entry,  a.get('gateway', 'N/A')),
                    (self.cur_mac_entry, mac),
                ]:
                    entry.config(state='normal')
                    entry.delete(0, 'end')
                    entry.insert(0, val)
                    entry.config(state='readonly')
                break

    def _check_duplicate_ips(self):
        ip_map = {}
        for a in self.adapters_cache:
            ip = a.get('ip', '')
            if ip and ip != 'N/A':
                ip_map.setdefault(ip, []).append(a['name'])
        dups = {ip: names for ip, names in ip_map.items() if len(names) > 1}
        if dups:
            parts = [f"IP {ip}  -->  {', '.join(names)}" for ip, names in dups.items()]
            self.warning_label.config(text="\n".join(parts))
            self.warning_frame.pack(fill='x', pady=(0, 10))
        else:
            self.warning_frame.pack_forget()

    def _validate_ip(self, ip_str):
        parts = ip_str.strip().split('.')
        if len(parts) != 4:
            return False
        for part in parts:
            try:
                num = int(part)
                if num < 0 or num > 255:
                    return False
            except ValueError:
                return False
        return True

    def _apply_static_ip(self):
        adapter = self.ipv4_adapter_var.get()
        ip = self.ip_entry.get().strip()
        subnet = self.subnet_entry.get().strip()
        gateway = self.gw_entry.get().strip()

        if not adapter:
            self.show_warning("No Adapter Selected", "Please select an adapter from the dropdown first.")
            return

        errors = []
        if not self._validate_ip(ip):      errors.append("Invalid IP Address format (expected x.x.x.x)")
        if not self._validate_ip(subnet):  errors.append("Invalid Subnet Mask format (expected x.x.x.x)")
        if gateway and not self._validate_ip(gateway): errors.append("Invalid Gateway format (expected x.x.x.x)")
        if errors:
            self.show_error("Validation Failed", "\n".join(errors))
            return

        for a in self.adapters_cache:
            if a['name'] != adapter and a.get('ip') == ip:
                if not self.ask_confirm("IP Conflict Detected",
                    f"IP {ip} is already assigned to '{a['name']}'.\n"
                    "This may cause network issues.\n\nProceed anyway?"):
                    return

        if not self.ask_confirm("Apply Network Changes",
            f"Set these values on '{adapter}'?\n\n"
            f"  IP:        {ip}\n  Subnet:  {subnet}\n  Gateway: {gateway or '(none)'}"):
            return

        def do():
            cmd = f'netsh interface ipv4 set address name="{adapter}" static {ip} {subnet}'
            if gateway: cmd += f' {gateway}'
            result = run_cmd(cmd)
            self.after(0, lambda: self._ipv4_done(result))
        self.ipv4_status.config(text="Applying changes...", fg=COLORS['orange'])
        threading.Thread(target=do, daemon=True).start()

    def _ipv4_done(self, result):
        if 'error' in result.lower() or 'failed' in result.lower():
            self.ipv4_status.config(text="", fg=COLORS['red'])
            self.show_error("Configuration Failed", result.strip())
        else:
            self.ipv4_status.config(text="", fg=COLORS['green'])
            self.toast("IP settings applied successfully!", 'success')
            self._refresh_ipv4()

    def _switch_to_dhcp(self):
        adapter = self.ipv4_adapter_var.get()
        if not adapter:
            self.show_warning("No Adapter Selected", "Please select an adapter from the dropdown first.")
            return
        if not self.ask_confirm("Switch to DHCP",
            f"Switch '{adapter}' to automatic (DHCP) mode?\n\n"
            "The system will request an IP from your DHCP server."):
            return
        def do():
            result = run_cmd(f'netsh interface ipv4 set address name="{adapter}" dhcp')
            self.after(0, lambda: self._ipv4_done(result))
        self.ipv4_status.config(text="Switching to DHCP...", fg=COLORS['orange'])
        threading.Thread(target=do, daemon=True).start()

    # ═══════════════════ TAB 2: ADAPTERS ═══════════════════
    def _build_adapters_tab(self, parent):
        container = tk.Frame(parent, bg=COLORS['bg_dark'])
        container.pack(fill='both', expand=True, padx=15, pady=10)

        top = tk.Frame(container, bg=COLORS['bg_dark'])
        top.pack(fill='x', pady=(0, 10))
        tk.Label(top, text="Network Adapters", bg=COLORS['bg_dark'], fg=COLORS['accent'],
                 font=('Segoe UI', 12, 'bold')).pack(side='left')
        tk.Button(top, text="  Refresh  ", bg=COLORS['accent'], fg='white',
                  font=('Segoe UI', 9, 'bold'), relief='flat', padx=12, pady=4,
                  cursor='hand2', command=self._refresh_adapters).pack(side='right')

        canvas_frame = tk.Frame(container, bg=COLORS['bg_dark'])
        canvas_frame.pack(fill='both', expand=True)
        self.adapter_canvas = tk.Canvas(canvas_frame, bg=COLORS['bg_dark'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.adapter_canvas.yview)
        self.adapter_list_frame = tk.Frame(self.adapter_canvas, bg=COLORS['bg_dark'])
        self.adapter_list_frame.bind('<Configure>',
            lambda e: self.adapter_canvas.configure(scrollregion=self.adapter_canvas.bbox('all')))
        self.adapter_canvas.create_window((0, 0), window=self.adapter_list_frame, anchor='nw')
        self.adapter_canvas.configure(yscrollcommand=scrollbar.set)
        self.adapter_canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self._refresh_adapters()

    def _refresh_adapters(self):
        def do():
            output = run_cmd('netsh interface show interface')
            adapters = []
            for line in output.strip().split('\n')[3:]:
                parts = line.split()
                if len(parts) >= 4:
                    admin_state, state, itype = parts[0], parts[1], parts[2]
                    name = ' '.join(parts[3:])
                    ip_out = run_cmd(f'netsh interface ipv4 show addresses name="{name}"')
                    ip = 'N/A'
                    for il in ip_out.split('\n'):
                        if 'IP Address' in il or 'IP address' in il:
                            p = il.split(':')
                            if len(p) >= 2: ip = p[-1].strip()
                    adapters.append({'name': name, 'admin': admin_state, 'state': state,
                                     'type': itype, 'ip': ip,
                                     'enabled': admin_state.lower() == 'enabled'})
            self.after(0, lambda: self._draw_adapters(adapters))
        threading.Thread(target=do, daemon=True).start()

    def _draw_adapters(self, adapters):
        for w in self.adapter_list_frame.winfo_children(): w.destroy()

        ethernet_count = sum(1 for a in adapters if 'ethernet' in a['type'].lower() or 'dedicated' in a['type'].lower())
        connected = sum(1 for a in adapters if a['state'].lower() == 'connected')

        summary = tk.Frame(self.adapter_list_frame, bg=COLORS['bg_card'], padx=12, pady=8,
                            highlightthickness=1, highlightbackground=COLORS['border'])
        summary.pack(fill='x', pady=(0, 8))
        tk.Label(summary,
                 text=f"Total: {len(adapters)}    Connected: {connected}    Ethernet-type: {ethernet_count}",
                 bg=COLORS['bg_card'], fg=COLORS['text'], font=('Segoe UI', 10, 'bold')).pack(side='left')

        for adapter in adapters:
            is_connected = adapter['state'].lower() == 'connected'
            border_color = COLORS['green_dark'] if is_connected else COLORS['border']
            card = tk.Frame(self.adapter_list_frame, bg=COLORS['bg_card'], padx=12, pady=10,
                            highlightthickness=2, highlightbackground=border_color)
            card.pack(fill='x', pady=3)

            tk.Frame(card, bg=COLORS['green'] if is_connected else COLORS['red'],
                     width=4).pack(side='left', fill='y', padx=(0, 10))

            info = tk.Frame(card, bg=COLORS['bg_card'])
            info.pack(side='left', fill='x', expand=True)
            color = COLORS['green'] if is_connected else COLORS['red']
            tk.Label(info, text=adapter['name'], bg=COLORS['bg_card'], fg=color,
                     font=('Segoe UI', 11, 'bold')).pack(anchor='w')
            tk.Label(info, text=f"Type: {adapter['type']}   IP: {adapter['ip']}   State: {adapter['state']}",
                     bg=COLORS['bg_card'], fg=COLORS['text_dim'], font=('Consolas', 9)).pack(anchor='w')

            is_on = adapter['enabled']
            btn_bg = COLORS['green_dark'] if is_on else COLORS['red_dark']
            btn_border = COLORS['green'] if is_on else COLORS['red']
            bw = tk.Frame(card, bg=btn_border, padx=1, pady=1)
            bw.pack(side='right', padx=5)
            tk.Button(bw, text="  ON  " if is_on else "  OFF  ", bg=btn_bg, fg='white',
                      font=('Segoe UI', 10, 'bold'), relief='flat', padx=14, pady=4,
                      cursor='hand2', activebackground=btn_border, activeforeground='white',
                      command=lambda n=adapter['name'], on=is_on: self._toggle_adapter(n, on)).pack()

    def _toggle_adapter(self, name, currently_on):
        word = "Disable" if currently_on else "Enable"
        if not self.ask_confirm(f"{word} Adapter",
            f"{word} the network adapter '{name}'?\n\n"
            f"{'This will disconnect you from this network.' if currently_on else 'This will attempt to connect.'}"):
            return
        def do():
            run_cmd(f'netsh interface set interface name="{name}" admin={"disable" if currently_on else "enable"}')
            self.after(500, self._refresh_adapters)
            self.after(600, lambda: self.toast(
                f"Adapter '{name}' {'disabled' if currently_on else 'enabled'}.",
                'warning' if currently_on else 'success'))
        threading.Thread(target=do, daemon=True).start()

    # ═══════════════════ TAB 3: PING ═══════════════════
    def _build_ping_tab(self, parent):
        container = tk.Frame(parent, bg=COLORS['bg_dark'])
        container.pack(fill='both', expand=True, padx=15, pady=10)

        # Saved IPs storage
        self._saved_ips_file = os.path.join(os.path.expanduser("~"), '.ncc_saved_ips.json')
        self._saved_ips = self._load_saved_ips()
        self._ping_process = None  # Track running ping subprocess

        tk.Label(container, text="Ping Diagnostic Tool", bg=COLORS['bg_dark'], fg=COLORS['accent'],
                 font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(0, 8))

        # ── Input row ──
        inp = tk.Frame(container, bg=COLORS['bg_dark'])
        inp.pack(fill='x', pady=(0, 8))
        tk.Label(inp, text="Target IP / Hostname:", bg=COLORS['bg_dark'], fg=COLORS['text'],
                 font=('Segoe UI', 10)).pack(side='left')

        entry_cfg = dict(bg=COLORS['bg_input'], fg=COLORS['text'], insertbackground=COLORS['text'],
                         font=('Consolas', 12), relief='flat', highlightthickness=1,
                         highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'], width=25)
        self.ping_entry = tk.Entry(inp, **entry_cfg)
        self.ping_entry.pack(side='left', padx=10)
        self.ping_entry.bind('<Return>', lambda e: self._do_ping())

        self.ping_btn = tk.Button(inp, text="  Ping  ", bg=COLORS['green_dark'], fg='white',
                                   font=('Segoe UI', 10, 'bold'), relief='flat', padx=16, pady=4,
                                   cursor='hand2', command=self._do_ping)
        self.ping_btn.pack(side='left', padx=(0, 4))

        # Stop button — red, hidden by default
        self.ping_stop_btn = tk.Button(inp, text="  Stop  ", bg=COLORS['red_dark'], fg='white',
                                        font=('Segoe UI', 10, 'bold'), relief='flat', padx=14, pady=4,
                                        cursor='hand2', command=self._stop_ping)
        # Not packed yet — shown only when pinging

        # ── Saved IPs section ──
        saved_section = tk.Frame(container, bg=COLORS['bg_card'], padx=10, pady=8,
                                  highlightthickness=1, highlightbackground=COLORS['border'])
        saved_section.pack(fill='x', pady=(0, 8))

        saved_header = tk.Frame(saved_section, bg=COLORS['bg_card'])
        saved_header.pack(fill='x', pady=(0, 6))

        tk.Label(saved_header, text="Saved IPs", bg=COLORS['bg_card'], fg=COLORS['text'],
                 font=('Segoe UI', 10, 'bold')).pack(side='left')

        self.saved_count_lbl = tk.Label(saved_header, text="", bg=COLORS['bg_card'],
                                         fg=COLORS['text_muted'], font=('Segoe UI', 8))
        self.saved_count_lbl.pack(side='left', padx=(6, 0))

        # Add IP row
        add_row = tk.Frame(saved_header, bg=COLORS['bg_card'])
        add_row.pack(side='right')

        self.add_ip_label_entry = tk.Entry(add_row, bg=COLORS['bg_input'], fg=COLORS['text'],
                                            insertbackground=COLORS['text'], font=('Segoe UI', 9),
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=COLORS['border'],
                                            highlightcolor=COLORS['accent'], width=12)
        self.add_ip_label_entry.pack(side='left', padx=2)
        self._setup_placeholder(self.add_ip_label_entry, "Label")

        self.add_ip_addr_entry = tk.Entry(add_row, bg=COLORS['bg_input'], fg=COLORS['text'],
                                           insertbackground=COLORS['text'], font=('Consolas', 9),
                                           relief='flat', highlightthickness=1,
                                           highlightbackground=COLORS['border'],
                                           highlightcolor=COLORS['accent'], width=16)
        self.add_ip_addr_entry.pack(side='left', padx=2)
        self._setup_placeholder(self.add_ip_addr_entry, "IP / Hostname")
        self.add_ip_addr_entry.bind('<Return>', lambda e: self._add_saved_ip())

        tk.Button(add_row, text=" + Add ", bg=COLORS['green_dark'], fg='white',
                  font=('Segoe UI', 8, 'bold'), relief='flat', padx=8, pady=2,
                  cursor='hand2', command=self._add_saved_ip).pack(side='left', padx=2)

        # Saved IPs button container (scrollable area)
        self.saved_ips_frame = tk.Frame(saved_section, bg=COLORS['bg_card'])
        self.saved_ips_frame.pack(fill='x')

        self._draw_saved_ips()

        # ── Ping output ──
        self.ping_output = scrolledtext.ScrolledText(
            container, bg=COLORS['bg_card'], fg=COLORS['text'], font=('Consolas', 10),
            relief='flat', height=12, insertbackground=COLORS['text'],
            highlightthickness=1, highlightbackground=COLORS['border'])
        self.ping_output.pack(fill='both', expand=True)
        self.ping_output.config(state='disabled')

        # ── Diagnosis frame ──
        self.diag_frame = tk.Frame(container, bg=COLORS['bg_dark'])
        self.diag_inner = tk.Frame(self.diag_frame, bg=COLORS['bg_card'], padx=14, pady=10,
                                    highlightthickness=2, highlightbackground=COLORS['border'])
        self.diag_inner.pack(fill='x')
        self.diag_icon = tk.Canvas(self.diag_inner, width=32, height=32,
                                    bg=COLORS['bg_card'], highlightthickness=0)
        self.diag_icon.pack(anchor='w')
        self.diag_label = tk.Label(self.diag_inner, text="", bg=COLORS['bg_card'], fg=COLORS['text'],
                                    font=('Segoe UI', 10), wraplength=900, justify='left')
        self.diag_label.pack(fill='x', pady=(4, 0))

    def _setup_placeholder(self, entry, placeholder):
        """Set up proper placeholder behavior — dim text that clears on focus, restores on blur."""
        entry.insert(0, placeholder)
        entry.config(fg=COLORS['text_muted'])
        entry._placeholder = placeholder
        entry._has_placeholder = True

        def on_focus_in(e):
            if entry._has_placeholder:
                entry.delete(0, 'end')
                entry.config(fg=COLORS['text'])
                entry._has_placeholder = False

        def on_focus_out(e):
            if not entry.get().strip():
                entry.delete(0, 'end')
                entry.insert(0, placeholder)
                entry.config(fg=COLORS['text_muted'])
                entry._has_placeholder = True

        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)

    def _get_entry_value(self, entry):
        """Get entry value, returning empty string if it's just the placeholder."""
        if hasattr(entry, '_has_placeholder') and entry._has_placeholder:
            return ""
        return entry.get().strip()

    def _load_saved_ips(self):
        """Load saved IPs from JSON file."""
        try:
            import json
            if os.path.exists(self._saved_ips_file):
                with open(self._saved_ips_file, 'r') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data[:10]  # Max 10
        except:
            pass
        return []

    def _save_ips_to_file(self):
        """Persist saved IPs to JSON file."""
        try:
            import json
            with open(self._saved_ips_file, 'w') as f:
                json.dump(self._saved_ips, f, indent=2)
        except:
            pass

    def _add_saved_ip(self):
        label = self._get_entry_value(self.add_ip_label_entry)
        addr = self._get_entry_value(self.add_ip_addr_entry)

        if not addr:
            self.show_warning("Missing IP", "Enter an IP address or hostname to save.")
            return

        if not label:
            label = addr  # Use address as label if none given

        if len(self._saved_ips) >= 10:
            self.show_warning("Limit Reached",
                              "You can save up to 10 IPs.\nRemove one first to add a new one.")
            return

        # Check for duplicate
        for item in self._saved_ips:
            if item['addr'] == addr:
                self.toast(f"'{addr}' is already saved.", 'warning')
                return

        self._saved_ips.append({'label': label, 'addr': addr})
        self._save_ips_to_file()
        self._draw_saved_ips()

        # Reset entries — restore placeholders
        self.add_ip_label_entry.delete(0, 'end')
        self.add_ip_label_entry.insert(0, "Label")
        self.add_ip_label_entry.config(fg=COLORS['text_muted'])
        self.add_ip_label_entry._has_placeholder = True

        self.add_ip_addr_entry.delete(0, 'end')
        self.add_ip_addr_entry.insert(0, "IP / Hostname")
        self.add_ip_addr_entry.config(fg=COLORS['text_muted'])
        self.add_ip_addr_entry._has_placeholder = True

        # Move focus away so placeholder looks right
        self.focus_set()

        self.toast(f"Saved '{label}' ({addr})", 'success')

    def _remove_saved_ip(self, index):
        if 0 <= index < len(self._saved_ips):
            removed = self._saved_ips.pop(index)
            self._save_ips_to_file()
            self._draw_saved_ips()
            self.toast(f"Removed '{removed['label']}'", 'warning')

    def _draw_saved_ips(self):
        """Redraw the saved IPs buttons."""
        for w in self.saved_ips_frame.winfo_children():
            w.destroy()

        count = len(self._saved_ips)
        self.saved_count_lbl.config(text=f"({count}/10)")

        if not self._saved_ips:
            tk.Label(self.saved_ips_frame, text="No saved IPs yet — add your frequently used ones above",
                     bg=COLORS['bg_card'], fg=COLORS['text_muted'],
                     font=('Segoe UI', 9)).pack(anchor='w', pady=2)
            return

        for i, item in enumerate(self._saved_ips):
            row = tk.Frame(self.saved_ips_frame, bg=COLORS['bg_card'])
            row.pack(fill='x', pady=1)

            # Click-to-ping button
            btn = tk.Button(row, text=f"  {item['label']}  ",
                            bg=COLORS['bg_input'], fg=COLORS['text'],
                            font=('Segoe UI', 9), relief='flat', padx=8, pady=2,
                            cursor='hand2',
                            highlightthickness=1, highlightbackground=COLORS['border'],
                            activebackground=COLORS['accent'], activeforeground='white',
                            command=lambda a=item['addr']: self._saved_ip_ping(a))
            btn.pack(side='left', padx=(0, 4))

            # Show address dimmed
            tk.Label(row, text=item['addr'], bg=COLORS['bg_card'],
                     fg=COLORS['text_muted'], font=('Consolas', 9)).pack(side='left', padx=(0, 4))

            # Remove button (small X)
            rm_btn = tk.Button(row, text=" ✕ ", bg=COLORS['bg_card'], fg=COLORS['red'],
                               font=('Segoe UI', 8), relief='flat', padx=2, pady=0,
                               cursor='hand2', activebackground=COLORS['red_dark'],
                               activeforeground='white',
                               command=lambda idx=i: self._remove_saved_ip(idx))
            rm_btn.pack(side='right')

    def _saved_ip_ping(self, addr):
        """Populate entry with saved IP and ping."""
        self.ping_entry.delete(0, 'end')
        self.ping_entry.insert(0, addr)
        self._do_ping()

    def _do_ping(self):
        target = self.ping_entry.get().strip()
        if not target:
            self.show_warning("Input Needed", "Enter an IP address or hostname to ping.")
            return

        # Show stop button, hide ping button
        self.ping_btn.pack_forget()
        self.ping_stop_btn.pack(side='left', padx=(0, 4))

        self.ping_output.config(state='normal')
        self.ping_output.delete('1.0', 'end')
        self.ping_output.insert('end', f"Pinging {target}...\n\n")
        self.ping_output.config(state='disabled')
        self.diag_frame.pack_forget()

        def do():
            try:
                kwargs = dict(stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                              text=True, shell=True)
                if platform.system() == 'Windows':
                    kwargs['creationflags'] = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)

                self._ping_process = subprocess.Popen(f'ping {target} -n 4', **kwargs)
                result = self._ping_process.communicate(timeout=30)[0]
                self._ping_process = None
            except subprocess.TimeoutExpired:
                if self._ping_process:
                    self._ping_process.kill()
                    self._ping_process = None
                result = "[Timed out — ping took too long]"
            except Exception as e:
                self._ping_process = None
                result = f"[Stopped or error: {e}]"

            resolved = ""
            try:
                hn = socket.gethostbyaddr(target)
                resolved = f"Resolved hostname: {hn[0]}\n"
            except: pass
            self.after(0, lambda: self._ping_done(target, result, resolved))
        threading.Thread(target=do, daemon=True).start()

    def _stop_ping(self):
        """Kill the running ping process."""
        if self._ping_process:
            try:
                self._ping_process.kill()
            except:
                pass
            self._ping_process = None

        # Restore buttons
        self.ping_stop_btn.pack_forget()
        self.ping_btn.pack(side='left', padx=(0, 4))

        self.ping_output.config(state='normal')
        self.ping_output.insert('end', "\n[Ping stopped by user]\n")
        self.ping_output.config(state='disabled')

        self.toast("Ping stopped.", 'warning')

    def _ping_done(self, target, result, resolved):
        # Restore buttons
        self.ping_stop_btn.pack_forget()
        self.ping_btn.pack(side='left', padx=(0, 4))

        self.ping_output.config(state='normal')
        if resolved: self.ping_output.insert('end', resolved + "\n")
        self.ping_output.insert('end', result)
        self.ping_output.config(state='disabled')

        diag = self._diagnose_ping(target, result)
        if diag:
            self.diag_frame.pack(fill='x', pady=(8, 0))
            self.diag_icon.delete('all')
            color = diag['color']
            self.diag_inner.config(highlightbackground=color)
            self.diag_icon.create_oval(2, 2, 30, 30, fill=color, outline='')
            self.diag_icon.create_text(16, 16, text=diag.get('sym', 'i'), fill='white',
                                        font=('Segoe UI', 14, 'bold'))
            self.diag_label.config(text=diag['text'], fg=COLORS['text'])

    def _diagnose_ping(self, target, result):
        r = result.lower()
        if 'reply from' in r and 'ttl=' in r:
            m = re.search(r'average\s*=\s*(\d+)', r)
            ms = int(m.group(1)) if m else 0
            q = "Excellent" if ms < 20 else ("Good" if ms < 100 else "Slow — might experience lag")
            return {'text': f"SUCCESS  —  {q} connection to {target} (avg {ms}ms). Network is healthy!",
                    'color': COLORS['green'], 'sym': '✓'}
        elif 'request timed out' in r:
            return {'text': "TIMED OUT  —  Target didn't respond.\n\nLikely causes:\n  - Device is offline or powered off\n  - Firewall blocking ICMP packets\n  - Wrong IP address\n  - Cable unplugged\n\nSuggestion: Check cables, verify the IP, ping your gateway first.",
                    'color': COLORS['red'], 'sym': '✕'}
        elif 'could not find host' in r or 'ping request could not find' in r:
            return {'text': "HOST NOT FOUND  —  DNS couldn't resolve this hostname.\n\nLikely causes:\n  - Typo in hostname/URL\n  - DNS server is down\n  - No internet connection\n\nSuggestion: Use IP directly, or set DNS to 8.8.8.8.",
                    'color': COLORS['red'], 'sym': '✕'}
        elif 'destination host unreachable' in r:
            return {'text': "UNREACHABLE  —  No route to host.\n\nLikely causes:\n  - Different network/subnet\n  - Gateway not set or wrong\n  - Cable disconnected\n\nSuggestion: Check gateway settings in IPv4 Config tab.",
                    'color': COLORS['orange'], 'sym': '!'}
        elif 'general failure' in r:
            return {'text': "GENERAL FAILURE  —  Network stack error.\n\nLikely causes:\n  - Adapter is disabled\n  - IPv4/IPv6 mismatch\n  - Driver issue\n\nSuggestion: Check Adapters tab — is your adapter ON?",
                    'color': COLORS['red'], 'sym': '✕'}
        elif 'transmit failed' in r:
            return {'text': "TRANSMIT FAILED  —  Couldn't send the ping.\n\nLikely causes:\n  - Adapter is down\n  - IP configuration invalid\n\nSuggestion: Reset adapter or switch to DHCP.",
                    'color': COLORS['red'], 'sym': '✕'}
        return None

    # ═══════════════════ TAB 4: CONNECTIONS ═══════════════════
    def _build_netstat_tab(self, parent):
        container = tk.Frame(parent, bg=COLORS['bg_dark'])
        container.pack(fill='both', expand=True, padx=15, pady=10)

        top = tk.Frame(container, bg=COLORS['bg_dark'])
        top.pack(fill='x', pady=(0, 10))
        tk.Label(top, text="Active Network Connections", bg=COLORS['bg_dark'], fg=COLORS['accent'],
                 font=('Segoe UI', 12, 'bold')).pack(side='left')

        bf = tk.Frame(top, bg=COLORS['bg_dark'])
        bf.pack(side='right')
        tk.Button(bf, text="  Refresh  ", bg=COLORS['accent'], fg='white',
                  font=('Segoe UI', 9, 'bold'), relief='flat', padx=12, pady=4,
                  cursor='hand2', command=self._refresh_netstat).pack(side='left', padx=3)

        ff = tk.Frame(container, bg=COLORS['bg_dark'])
        ff.pack(fill='x', pady=(0, 5))
        tk.Label(ff, text="Filter:", bg=COLORS['bg_dark'], fg=COLORS['text_dim'],
                 font=('Segoe UI', 9)).pack(side='left')
        self.netstat_filter = tk.Entry(ff, bg=COLORS['bg_input'], fg=COLORS['text'],
                                        insertbackground=COLORS['text'], font=('Consolas', 10),
                                        relief='flat', highlightthickness=1,
                                        highlightbackground=COLORS['border'],
                                        highlightcolor=COLORS['accent'], width=30)
        self.netstat_filter.pack(side='left', padx=8)
        self.netstat_filter.bind('<Return>', lambda e: self._apply_netstat_filter())
        tk.Button(ff, text=" Apply ", bg=COLORS['bg_card'], fg=COLORS['text'], font=('Segoe UI', 8),
                  relief='flat', padx=8, highlightthickness=1, highlightbackground=COLORS['border'],
                  command=self._apply_netstat_filter).pack(side='left')

        self.netstat_summary = tk.Label(container, text="", bg=COLORS['bg_card'],
                                         fg=COLORS['text_dim'], font=('Segoe UI', 9), padx=10, pady=5)
        self.netstat_summary.pack(fill='x', pady=(0, 5))

        tf = tk.Frame(container, bg=COLORS['bg_dark'])
        tf.pack(fill='both', expand=True)
        columns = ('proto', 'local_ip', 'local_port', 'remote_ip', 'remote_port', 'state', 'process')
        self.net_tree = ttk.Treeview(tf, columns=columns, show='headings', height=18)
        for col, (h, w) in {'proto': ('Protocol', 70), 'local_ip': ('Local IP', 140),
                              'local_port': ('L.Port', 65), 'remote_ip': ('Remote IP', 150),
                              'remote_port': ('R.Port', 65), 'state': ('State', 110),
                              'process': ('Process (PID)', 180)}.items():
            self.net_tree.heading(col, text=h)
            self.net_tree.column(col, width=w, minwidth=50)
        sy = ttk.Scrollbar(tf, orient='vertical', command=self.net_tree.yview)
        self.net_tree.configure(yscrollcommand=sy.set)
        self.net_tree.pack(side='left', fill='both', expand=True)
        sy.pack(side='right', fill='y')

        self.netstat_data = []
        self._refresh_netstat()

    def _get_pid_map(self):
        pid_map = {}
        for line in run_cmd('tasklist /fo csv /nh').strip().split('\n'):
            parts = line.strip().split(',')
            if len(parts) >= 2:
                pid_map[parts[1].strip('"')] = parts[0].strip('"')
        return pid_map

    def _parse_netstat(self, output, pid_map):
        connections = []
        for line in output.strip().split('\n'):
            parts = line.strip().split()
            if len(parts) >= 4 and parts[0] in ('TCP', 'UDP'):
                proto = parts[0]
                local = parts[1]
                remote = parts[2] if proto == 'TCP' else '*:*'
                if proto == 'TCP' and len(parts) >= 5:
                    state, pid = parts[3], parts[4]
                else:
                    state = '-'
                    pid = parts[3] if len(parts) > 3 else '?'
                lp = local.rsplit(':', 1)
                rp = remote.rsplit(':', 1)
                connections.append({
                    'proto': proto,
                    'local_ip': lp[0] if len(lp) == 2 else local,
                    'local_port': lp[1] if len(lp) == 2 else '',
                    'remote_ip': rp[0] if len(rp) == 2 else remote,
                    'remote_port': rp[1] if len(rp) == 2 else '',
                    'state': state,
                    'process': f"{pid_map.get(pid, '?')} ({pid})",
                    'pid': pid,
                })
        return connections

    def _refresh_netstat(self):
        def do():
            pid_map = self._get_pid_map()
            output = run_cmd('netstat -ano')
            self.netstat_data = self._parse_netstat(output, pid_map)
            self.after(0, lambda: self._draw_netstat(self.netstat_data))
        threading.Thread(target=do, daemon=True).start()

    def _draw_netstat(self, connections):
        for item in self.net_tree.get_children(): self.net_tree.delete(item)
        tc = sum(1 for c in connections if c['proto'] == 'TCP')
        uc = sum(1 for c in connections if c['proto'] == 'UDP')
        est = sum(1 for c in connections if c['state'] == 'ESTABLISHED')
        lis = sum(1 for c in connections if c['state'] == 'LISTENING')
        self.netstat_summary.config(
            text=f"Total: {len(connections)}    TCP: {tc}    UDP: {uc}    ESTABLISHED: {est}    LISTENING: {lis}")
        for c in connections:
            self.net_tree.insert('', 'end', values=(
                c['proto'], c['local_ip'], c['local_port'],
                c['remote_ip'], c['remote_port'], c['state'], c['process']))

    def _apply_netstat_filter(self):
        q = self.netstat_filter.get().strip().lower()
        if not q: self._draw_netstat(self.netstat_data); return
        self._draw_netstat([c for c in self.netstat_data if q in str(c).lower()])

    def _suggest_ports(self):
        used = set()
        for c in self.netstat_data:
            try: used.add(int(c['local_port']))
            except: pass

        win = tk.Toplevel(self)
        win.title("Port Suggestions")
        win.configure(bg=COLORS['bg_dark'])
        win.geometry("420x520")
        win.transient(self)
        try:
            win._icon = _make_icon_photoimage(win, 32)
            win.wm_iconphoto(True, win._icon)
        except: pass

        tk.Label(win, text="Port Availability Report", bg=COLORS['bg_dark'], fg=COLORS['accent'],
                 font=('Segoe UI', 13, 'bold')).pack(pady=(15, 10))

        s1 = tk.Frame(win, bg=COLORS['bg_card'], padx=12, pady=10,
                       highlightthickness=1, highlightbackground=COLORS['border'])
        s1.pack(fill='x', padx=15, pady=(0, 10))
        tk.Label(s1, text="Common Development Ports", bg=COLORS['bg_card'], fg=COLORS['text'],
                 font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 6))
        for port in [3000, 3001, 4000, 5000, 5001, 8000, 8080, 8081, 8888, 9000, 9090]:
            row = tk.Frame(s1, bg=COLORS['bg_card']); row.pack(fill='x', pady=1)
            free = port not in used
            dc = COLORS['green'] if free else COLORS['red']
            tk.Label(row, text="●", bg=COLORS['bg_card'], fg=dc, font=('Segoe UI', 10)).pack(side='left')
            tk.Label(row, text=f"  {port}", bg=COLORS['bg_card'], fg=COLORS['text'], font=('Consolas', 10)).pack(side='left')
            tk.Label(row, text="FREE" if free else "IN USE", bg=COLORS['bg_card'], fg=dc,
                     font=('Segoe UI', 9, 'bold')).pack(side='right')

        s2 = tk.Frame(win, bg=COLORS['bg_card'], padx=12, pady=10,
                       highlightthickness=1, highlightbackground=COLORS['border'])
        s2.pack(fill='x', padx=15)
        tk.Label(s2, text="Recommended Free Ports (Dynamic Range)", bg=COLORS['bg_card'],
                 fg=COLORS['text'], font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 6))
        sugg = [str(p) for p in range(49152, 49200) if p not in used][:10]
        tk.Label(s2, text=", ".join(sugg), bg=COLORS['bg_card'], fg=COLORS['green'],
                 font=('Consolas', 10), wraplength=380).pack(anchor='w')
        tk.Label(s2, text="(Range 49152-65535 — safe for custom use)", bg=COLORS['bg_card'],
                 fg=COLORS['text_dim'], font=('Segoe UI', 8)).pack(anchor='w', pady=(4, 0))

        tk.Button(win, text="  Close  ", bg=COLORS['accent'], fg='white',
                  font=('Segoe UI', 10, 'bold'), relief='flat', padx=20, pady=6,
                  cursor='hand2', command=win.destroy).pack(pady=15)

    # ═══════════════════ TAB 5: PORT LOOKUP ═══════════════════
    def _build_port_lookup_tab(self, parent):
        container = tk.Frame(parent, bg=COLORS['bg_dark'])
        container.pack(fill='both', expand=True, padx=15, pady=10)

        tk.Label(container, text="Port Lookup Tool", bg=COLORS['bg_dark'], fg=COLORS['accent'],
                 font=('Segoe UI', 12, 'bold')).pack(anchor='w', pady=(0, 4))
        tk.Label(container,
                 text="Check if a port is in use and which app is occupying it  (like netstat -an | findstr <port>)",
                 bg=COLORS['bg_dark'], fg=COLORS['text_dim'], font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 10))

        inp = tk.Frame(container, bg=COLORS['bg_dark'])
        inp.pack(fill='x', pady=(0, 10))
        tk.Label(inp, text="Port Number:", bg=COLORS['bg_dark'], fg=COLORS['text'],
                 font=('Segoe UI', 10)).pack(side='left')
        self.port_entry = tk.Entry(inp, bg=COLORS['bg_input'], fg=COLORS['text'],
                                    insertbackground=COLORS['text'], font=('Consolas', 14),
                                    relief='flat', highlightthickness=1,
                                    highlightbackground=COLORS['border'],
                                    highlightcolor=COLORS['accent'], width=10)
        self.port_entry.pack(side='left', padx=10)
        self.port_entry.bind('<Return>', lambda e: self._do_port_lookup())

        self.port_btn = tk.Button(inp, text="  Scan Port  ", bg=COLORS['accent'], fg='white',
                                   font=('Segoe UI', 10, 'bold'), relief='flat', padx=16, pady=4,
                                   cursor='hand2', command=self._do_port_lookup)
        self.port_btn.pack(side='left', padx=(0, 6))

        tk.Button(inp, text="  Suggest Free Ports  ", bg=COLORS['bg_card'], fg=COLORS['text'],
                  font=('Segoe UI', 9), relief='flat', padx=12, pady=4, cursor='hand2',
                  highlightthickness=1, highlightbackground=COLORS['border'],
                  command=self._suggest_ports).pack(side='left', padx=(0, 4))

        quick = tk.Frame(container, bg=COLORS['bg_dark'])
        quick.pack(fill='x', pady=(0, 10))
        tk.Label(quick, text="Common ports:", bg=COLORS['bg_dark'], fg=COLORS['text_dim'],
                 font=('Segoe UI', 9)).pack(side='left')
        for label, port in [("HTTP 80", "80"), ("HTTPS 443", "443"), ("SSH 22", "22"),
                            ("FTP 21", "21"), ("DNS 53", "53"), ("RDP 3389", "3389"),
                            ("MySQL 3306", "3306"), ("PgSQL 5432", "5432")]:
            tk.Button(quick, text=f" {label} ", bg=COLORS['bg_card'], fg=COLORS['text_dim'],
                      font=('Segoe UI', 8), relief='flat', padx=6, pady=2, cursor='hand2',
                      highlightthickness=1, highlightbackground=COLORS['border'],
                      command=lambda p=port: self._quick_port(p)).pack(side='left', padx=2)

        # Result card area
        self.port_result_frame = tk.Frame(container, bg=COLORS['bg_dark'])
        self.port_result_frame.pack(fill='x', pady=(0, 10))

        tk.Label(container, text="Matching Connections:", bg=COLORS['bg_dark'], fg=COLORS['text'],
                 font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 4))

        tf = tk.Frame(container, bg=COLORS['bg_dark'])
        tf.pack(fill='both', expand=True)
        port_cols = ('proto', 'local_addr', 'remote_addr', 'state', 'process', 'pid')
        self.port_tree = ttk.Treeview(tf, columns=port_cols, show='headings', height=14)
        for col, (h, w) in {'proto': ('Protocol', 75), 'local_addr': ('Local Address', 180),
                              'remote_addr': ('Remote Address', 180), 'state': ('State', 120),
                              'process': ('Process', 160), 'pid': ('PID', 70)}.items():
            self.port_tree.heading(col, text=h)
            self.port_tree.column(col, width=w, minwidth=50)
        ps = ttk.Scrollbar(tf, orient='vertical', command=self.port_tree.yview)
        self.port_tree.configure(yscrollcommand=ps.set)
        self.port_tree.pack(side='left', fill='both', expand=True)
        ps.pack(side='right', fill='y')

    def _quick_port(self, port):
        self.port_entry.delete(0, 'end')
        self.port_entry.insert(0, port)
        self._do_port_lookup()

    def _do_port_lookup(self):
        port_str = self.port_entry.get().strip()
        if not port_str:
            self.show_warning("No Port Entered", "Please type a port number to look up.")
            return
        try:
            port_num = int(port_str)
            if port_num < 0 or port_num > 65535: raise ValueError
        except ValueError:
            self.show_error("Invalid Port", f"'{port_str}' is not a valid port number.\nEnter a number between 0 and 65535.")
            return

        self.port_btn.config(state='disabled', text="  Scanning...  ")

        def do():
            output = run_cmd(f'netstat -ano | findstr :{port_num}')
            pid_map = self._get_pid_map()

            matches = []
            for line in output.strip().split('\n'):
                parts = line.strip().split()
                if len(parts) >= 4 and parts[0] in ('TCP', 'UDP'):
                    proto = parts[0]
                    local = parts[1]
                    remote = parts[2] if proto == 'TCP' else '*:*'
                    if proto == 'TCP' and len(parts) >= 5:
                        state, pid = parts[3], parts[4]
                    else:
                        state, pid = '-', parts[3] if len(parts) > 3 else '?'

                    lp = local.rsplit(':', 1)[-1] if ':' in local else ''
                    rp = remote.rsplit(':', 1)[-1] if ':' in remote else ''
                    if lp == str(port_num) or rp == str(port_num):
                        matches.append({'proto': proto, 'local': local, 'remote': remote,
                                        'state': state, 'process': pid_map.get(pid, 'Unknown'), 'pid': pid})

            self.after(0, lambda: self._port_lookup_done(port_num, matches))
        threading.Thread(target=do, daemon=True).start()

    def _port_lookup_done(self, port, matches):
        self.port_btn.config(state='normal', text="  Scan Port  ")
        for w in self.port_result_frame.winfo_children(): w.destroy()
        for item in self.port_tree.get_children(): self.port_tree.delete(item)

        well_known = {20: "FTP Data", 21: "FTP Control", 22: "SSH", 23: "Telnet", 25: "SMTP",
                      53: "DNS", 67: "DHCP Server", 68: "DHCP Client", 80: "HTTP", 110: "POP3",
                      143: "IMAP", 443: "HTTPS", 445: "SMB", 993: "IMAPS", 995: "POP3S",
                      1433: "MSSQL", 1521: "Oracle DB", 3306: "MySQL", 3389: "RDP",
                      5432: "PostgreSQL", 5900: "VNC", 6379: "Redis", 8080: "HTTP Proxy",
                      8443: "HTTPS Alt", 27017: "MongoDB"}

        if matches:
            border = COLORS['red']
            card = tk.Frame(self.port_result_frame, bg=border, padx=2, pady=2)
            card.pack(fill='x')
            inner = tk.Frame(card, bg=COLORS['bg_card'], padx=14, pady=10)
            inner.pack(fill='both')

            row = tk.Frame(inner, bg=COLORS['bg_card'])
            row.pack(fill='x')
            ic = tk.Canvas(row, width=36, height=36, bg=COLORS['bg_card'], highlightthickness=0)
            ic.pack(side='left', padx=(0, 10))
            ic.create_oval(2, 2, 34, 34, fill=COLORS['red_dark'], outline=COLORS['red'])
            ic.create_text(18, 18, text="!", fill='white', font=('Segoe UI', 16, 'bold'))

            info = tk.Frame(row, bg=COLORS['bg_card'])
            info.pack(side='left', fill='x', expand=True)
            procs = set(f"{m['process']} (PID {m['pid']})" for m in matches)
            tk.Label(info, text=f"PORT {port} IS IN USE", bg=COLORS['bg_card'], fg=COLORS['red'],
                     font=('Segoe UI', 12, 'bold')).pack(anchor='w')
            tk.Label(info, text=f"Occupied by: {', '.join(procs)}", bg=COLORS['bg_card'],
                     fg=COLORS['text'], font=('Segoe UI', 10)).pack(anchor='w')
            tk.Label(info, text=f"{len(matches)} connection(s) found using this port",
                     bg=COLORS['bg_card'], fg=COLORS['text_dim'], font=('Segoe UI', 9)).pack(anchor='w')

            if port in well_known:
                tk.Label(inner, text=f"Standard service: {well_known[port]}", bg=COLORS['bg_card'],
                         fg=COLORS['text_dim'], font=('Segoe UI', 9)).pack(anchor='w', pady=(6, 0))

            tk.Label(inner, text=f"To free this port:  taskkill /PID {matches[0]['pid']} /F",
                     bg=COLORS['bg_card'], fg=COLORS['orange'], font=('Consolas', 9)).pack(anchor='w', pady=(6, 0))

            for m in matches:
                self.port_tree.insert('', 'end', values=(
                    m['proto'], m['local'], m['remote'], m['state'], m['process'], m['pid']))
        else:
            border = COLORS['green']
            card = tk.Frame(self.port_result_frame, bg=border, padx=2, pady=2)
            card.pack(fill='x')
            inner = tk.Frame(card, bg=COLORS['bg_card'], padx=14, pady=10)
            inner.pack(fill='both')

            row = tk.Frame(inner, bg=COLORS['bg_card'])
            row.pack(fill='x')
            ic = tk.Canvas(row, width=36, height=36, bg=COLORS['bg_card'], highlightthickness=0)
            ic.pack(side='left', padx=(0, 10))
            ic.create_oval(2, 2, 34, 34, fill=COLORS['green_dark'], outline=COLORS['green'])
            ic.create_text(18, 18, text="✓", fill='white', font=('Segoe UI', 16, 'bold'))

            info = tk.Frame(row, bg=COLORS['bg_card'])
            info.pack(side='left')
            tk.Label(info, text=f"PORT {port} IS FREE", bg=COLORS['bg_card'], fg=COLORS['green'],
                     font=('Segoe UI', 12, 'bold')).pack(anchor='w')
            tk.Label(info, text="No application is currently using this port. You're good to go!",
                     bg=COLORS['bg_card'], fg=COLORS['text'], font=('Segoe UI', 10)).pack(anchor='w')
            if port in well_known:
                tk.Label(info, text=f"Note: Port {port} is commonly used for {well_known[port]}",
                         bg=COLORS['bg_card'], fg=COLORS['text_dim'], font=('Segoe UI', 9)).pack(anchor='w', pady=(4, 0))


# ═══════════════════════════ ENTRY POINT ═══════════════════════════
if __name__ == "__main__":
    if platform.system() == 'Windows':
        run_as_admin()

    try:
        app = NetworkCommandCenter()
        app.mainloop()
    except Exception as e:
        # If the app crashes, show the error instead of vanishing silently
        import traceback
        error_details = traceback.format_exc()
        try:
            root = tk.Tk()
            root.withdraw()
            # Use basic messagebox as fallback since CustomDialog needs a visible parent
            from tkinter import messagebox as mb
            mb.showerror("Network Command Center — Crash Report",
                         f"The app encountered an error and needs to close.\n\n"
                         f"Error: {e}\n\n"
                         f"Details:\n{error_details}\n\n"
                         f"Please report this to the developer.")
            root.destroy()
        except:
            # Last resort: dump to a log file on the desktop
            log_path = os.path.join(os.path.expanduser("~"), "Desktop", "ncc_crash.log")
            with open(log_path, 'w') as f:
                f.write(f"NCC Crash Report\n{'='*50}\n{error_details}")
