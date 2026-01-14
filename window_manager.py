import win32api
import win32con
import win32gui
import win32process
import os
import ctypes
from ctypes import wintypes

def get_window_exe_path(hwnd):
    """
    Gets the executable path for the process that owns the window.
    """
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        path = win32process.GetModuleFileNameEx(handle, 0)
        win32api.CloseHandle(handle)
        return path
    except Exception:
        return None

def list_visible_windows():
    """
    Lists visible windows that are at least 1/10 of the screen size.
    Returns a list of dicts: {'hwnd': hwnd, 'title': title, 'rect': (l, t, r, b), 'path': exe_path}
    """
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    min_area = (screen_width * screen_height) / 10

    windows = []

    def enum_handler(hwnd, lParam):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            
            # Exclude some special windows
            if title in ["Program Manager", "Settings", "Microsoft Text Input Application", "Task View"]:
                return True

            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            area = w * h

            if area >= min_area:
                exe_path = get_window_exe_path(hwnd)
                windows.append({
                    'hwnd': hwnd,
                    'title': title,
                    'rect': rect,
                    'path': exe_path
                })
        return True

    win32gui.EnumWindows(enum_handler, None)
    return windows

def get_window_rect_actual(hwnd):
    """
    获取窗口的实际可见矩形（排除 Win10/11 的不可见阴影边框）
    """
    rect = win32gui.GetWindowRect(hwnd)
    try:
        # DWMWA_EXTENDED_FRAME_BOUNDS = 9
        actual_rect = wintypes.RECT()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, 9, ctypes.byref(actual_rect), ctypes.sizeof(actual_rect)
        )
        return (actual_rect.left, actual_rect.top, actual_rect.right, actual_rect.bottom)
    except Exception:
        return rect

def tile_windows(hwnds, rows, cols, compact=False):
    """
    将窗口排列成网格。如果 compact 为 True，则尝试消除系统阴影导致的间隙。
    """
    if not hwnds:
        return

    # Work area (excludes taskbar)
    # Using GetMonitorInfo for the primary monitor
    monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0), win32con.MONITOR_DEFAULTTOPRIMARY))
    work_area = monitor_info['Work'] # (left, top, right, bottom)
    
    wa_left, wa_top, wa_right, wa_bottom = work_area
    wa_width = wa_right - wa_left
    wa_height = wa_bottom - wa_top

    cell_width = wa_width // cols
    cell_height = wa_height // rows

    for i, hwnd in enumerate(hwnds):
        if i >= rows * cols:
            break
        
        row = i // cols
        col = i % cols

        x = wa_left + col * cell_width
        y = wa_top + row * cell_height
        w, h = cell_width, cell_height

        if compact:
            # 计算偏移逻辑：
            # 实际位置 = 目标位置 - (实际边界.left - 窗口边界.left)
            full_rect = win32gui.GetWindowRect(hwnd)
            actual_rect = get_window_rect_actual(hwnd)
            
            offset_l = actual_rect[0] - full_rect[0]
            offset_t = actual_rect[1] - full_rect[1]
            offset_r = full_rect[2] - actual_rect[2]
            offset_b = full_rect[3] - actual_rect[3]
            
            x -= offset_l
            y -= offset_t
            w += (offset_l + offset_r)
            h += (offset_t + offset_b)

        # Use SetWindowPos to move and bring to top
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        flags = win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, flags)
        
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
