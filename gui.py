import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QListWidget, QListWidgetItem, QPushButton, 
                             QLabel, QSpinBox, QMessageBox, QFileIconProvider)
from PySide6.QtCore import Qt, QFileInfo
from PySide6.QtGui import QIcon, QPixmap
import window_manager

def resource_path(relative_path):
    """ 获取资源的绝对路径，兼容开发环境和 PyInstaller 打包环境 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class GridWinApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GridWin - 窗口网格排列工具(surewin.chau@hotmail.com)")
        self.resize(700, 500)
        
        # Set Application Icon
        icon_path = resource_path("image.jpg")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.icon_provider = QFileIconProvider()
        self.init_ui()
        self.refresh_window_list()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Window List Section
        main_layout.addWidget(QLabel("选择要排列的窗口（当前可见且面积 > 1/10 屏幕）："))
        self.window_list = QListWidget()
        self.window_list.setSelectionMode(QListWidget.MultiSelection)
        self.window_list.setIconSize(self.window_list.iconSize() * 1.5)
        main_layout.addWidget(self.window_list)

        # Controls Section
        controls_layout = QHBoxLayout()
        
        # Grid settings
        controls_layout.addWidget(QLabel("行数:"))
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 10)
        self.rows_spin.setValue(2)
        controls_layout.addWidget(self.rows_spin)

        controls_layout.addWidget(QLabel("列数:"))
        self.cols_spin = QSpinBox()
        self.cols_spin.setRange(1, 10)
        self.cols_spin.setValue(2)
        controls_layout.addWidget(self.cols_spin)

        controls_layout.addStretch()

        # Action buttons
        self.refresh_btn = QPushButton("刷新列表")
        self.refresh_btn.clicked.connect(self.refresh_window_list)
        controls_layout.addWidget(self.refresh_btn)

        self.tile_btn = QPushButton("开始排列")
        self.tile_btn.setStyleSheet("background-color: #0078d4; color: white; font-weight: bold; padding: 5px 20px;")
        self.tile_btn.clicked.connect(self.tile_selected_windows)
        controls_layout.addWidget(self.tile_btn)

        main_layout.addLayout(controls_layout)

    def refresh_window_list(self):
        self.window_list.clear()
        self.windows = window_manager.list_visible_windows()
        
        # Sort by title
        self.windows.sort(key=lambda x: x['title'].lower())

        for win in self.windows:
            item = QListWidgetItem(win['title'])
            item.setData(Qt.UserRole, win['hwnd'])
            
            # Get and set icon using QFileIconProvider
            if win['path'] and os.path.exists(win['path']):
                try:
                    file_info = QFileInfo(win['path'])
                    icon = self.icon_provider.icon(file_info)
                    if not icon.isNull():
                        item.setIcon(icon)
                except Exception:
                    pass
            
            self.window_list.addItem(item)

    def tile_selected_windows(self):
        selected_items = self.window_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "未选择", "请至少选择一个窗口进行排列。")
            return

        hwnds = [item.data(Qt.UserRole) for item in selected_items]
        rows = self.rows_spin.value()
        cols = self.cols_spin.value()
        grid_capacity = rows * cols

        if len(hwnds) > grid_capacity:
            reply = QMessageBox.question(
                self, 
                "确认排列", 
                f"您选择了 {len(hwnds)} 个窗口，但网格容量仅为 {grid_capacity}。\n"
                f"排列将仅处理前 {grid_capacity} 个窗口。是否继续？",
                QMessageBox.Yes | QMessageBox.No, 
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        window_manager.tile_windows(hwnds, rows, cols)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GridWinApp()
    window.show()
    sys.exit(app.exec())
