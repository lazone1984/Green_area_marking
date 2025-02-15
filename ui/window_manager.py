import win32gui
import win32con

class WindowManager:
    def init_window_handles(self):
        """初始化窗口句柄"""
        self.cad_hwnd = None
        self.ui_hwnd = None
        self.root.update()
        self.ui_hwnd = self.root.winfo_id()
        self.find_cad_window()

    def find_cad_window(self):
        """查找CAD窗口句柄"""
        # ... (原find_cad_window方法的代码)

    def switch_to_cad(self):
        """切换到CAD窗口"""
        # ... (原switch_to_cad方法的代码)

    def switch_to_ui(self):
        """切换回UI窗口"""
        # ... (原switch_to_ui方法的代码) 