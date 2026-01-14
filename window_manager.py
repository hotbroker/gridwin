import win32process
import os
import win32gui
import win32con
import win32api

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

def tile_windows(hwnds, rows, cols):
    """
    Tiles the given windows into a grid on the primary monitor.
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

        # Use SetWindowPos to move and bring to top
        # HWND_TOP: Brings the window to the top of the Z order.
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        flags = win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, cell_width, cell_height, flags)
        
        # Additionally force to foreground to ensure it's active if desired, 
        # but HWND_TOP should be enough for visibility.
        win32gui.SetForegroundWindow(hwnd)
