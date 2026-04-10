import sys
import os
import datetime
import getpass
import ctypes
import webbrowser
import re
import json

try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    pass

try:
    from packaging import version
    HAS_PACKAGING = True
except ImportError:
    HAS_PACKAGING = False

import requests
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QCheckBox, 
                            QFrame, QMessageBox, QSystemTrayIcon,
                            QMenu, QAction, QGraphicsDropShadowEffect, QStyle)
from PyQt5.QtCore import Qt, QTimer, QSettings, QSize, QPoint, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QIcon, QPixmap, QImage, QColor, QFont, QPainter, QPainterPath

# ==========================================
# 0. 核心工具函数 - 修复图标路径问题
# ==========================================
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ==========================================
# 1. 业务逻辑
# ==========================================
class WallpaperUtils:
    @staticmethod
    def get_bing_url():
        url = "https://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
        response = requests.get(url, verify=False, timeout=10)
        response.raise_for_status()
        data = response.json()
        imgurl_base = data['images'][0]['urlbase']
        return f"https://cn.bing.com{imgurl_base}_UHD.jpg"

    @staticmethod
    def download_image(url, save_path):
        response = requests.get(url, verify=False, stream=True, timeout=30)
        response.raise_for_status()
        with open(save_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk: f.write(chunk)
        return save_path

    @staticmethod
    def set_wallpaper_api(image_path):
        abs_path = os.path.abspath(image_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError("壁纸文件未找到")
        SPI_SETDESKWALLPAPER = 20
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, abs_path, 3)

    @staticmethod
    def clean_old_wallpapers(save_dir):
        today = datetime.date.today()
        count = 0
        if not os.path.exists(save_dir): return 0
        for filename in os.listdir(save_dir):
            if filename.endswith("_UHD.jpg"):
                try:
                    date_str = filename.split("_")[0]
                    file_date = datetime.datetime.strptime(date_str, "%Y%m%d").date()
                    if file_date != today:
                        os.remove(os.path.join(save_dir, filename))
                        count += 1
                except: pass
        return count

    @staticmethod
    def check_update(current_ver):
        if not HAS_PACKAGING:
            return False, "库缺失", ""
        api_url = "https://api.github.com/repos/QsSama-W/wallpaper-win/releases/latest"
        response = requests.get(api_url, verify=True, timeout=5)
        response.raise_for_status()
        info = response.json()
        latest_tag = re.sub(r'^v', '', info.get('tag_name', ''))
        
        if version.parse(latest_tag) > version.parse(current_ver):
            return True, latest_tag, info.get('html_url')
        return False, latest_tag, ""

# ==========================================
# 2. 线程 Worker
# ==========================================
class WorkerSignals(QObject):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

class Worker(QThread):
    def __init__(self, task_func, *args, **kwargs):
        super().__init__()
        self.task_func = task_func
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.task_func(*self.args, **self.kwargs)
            self.signals.finished.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))

# ==========================================
# 3. UI 组件
# ==========================================
class ModernButton(QPushButton):
    def __init__(self, text, color="#007AFF", parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(42)
        self.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        self.update_style(color)

    def update_style(self, color):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border-radius: 21px;
                border: none;
                padding: 0 20px;
            }}
            QPushButton:hover {{ background-color: {self.adjust_color(color, 20)}; }}
            QPushButton:pressed {{ background-color: {self.adjust_color(color, -20)}; }}
            QPushButton:disabled {{ background-color: #cccccc; }}
        """)

    def adjust_color(self, hex_color, factor):
        c = QColor(hex_color)
        r = max(0, min(255, c.red() + factor))
        g = max(0, min(255, c.green() + factor))
        b = max(0, min(255, c.blue() + factor))
        return f"rgb({r}, {g}, {b})"

class ModernCheckBox(QCheckBox):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Microsoft YaHei", 10))
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("""
            QCheckBox { color: #555; spacing: 8px; }
            QCheckBox::indicator {
                width: 20px; height: 20px;
                border-radius: 6px;
                border: 2px solid #ddd;
                background: rgba(255,255,255,0.8);
            }
            QCheckBox::indicator:unchecked:hover { border-color: #007AFF; }
            QCheckBox::indicator:checked {
                background-color: #007AFF; border-color: #007AFF;
            }
        """)

# ==========================================
# 4. 主程序
# ==========================================
class BingWallpaperApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_version = "1.4.0" 
        self.username = getpass.getuser()
        self.save_dir = os.path.join(f"C:\\Users\\{self.username}", "Pictures", "bing_wallpaper")
        
        self.images_dir = resource_path("images")
        
        self.settings = QSettings("BingWallpaper", "Manager")
        
        self.preview_worker = None
        self.download_worker = None
        self.update_worker = None
        self.is_download_running = False
        self.is_preview_running = False
        
        self.shadow_margin = 25
        self.content_width = 680
        self.content_height = 780 
        self.resize(self.content_width + 2*self.shadow_margin, self.content_height + 2*self.shadow_margin)
        
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowSystemMenuHint | Qt.WindowMinimizeButtonHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        logo_path = os.path.join(self.images_dir, "logo.png")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        self.m_flag = False
        self.m_Position = QPoint()

        if not self.check_screen_resolution():
            QTimer.singleShot(0, QApplication.quit)
            return

        self.setup_ui()
        self.setup_tray()
        self.load_settings()
        
        os.makedirs(self.save_dir, exist_ok=True)
        
        if self.auto_start_chk.isChecked():
            self.set_startup_registry(True)
            
        self.hide()
        self.tray_icon.show()
        
        QTimer.singleShot(500, self.start_refresh_preview)
        QTimer.singleShot(2000, self.start_auto_download)
        
        if self.auto_update_chk.isChecked():
            QTimer.singleShot(5000, self.start_check_update)

    def setup_ui(self):
        base_widget = QWidget()
        self.setCentralWidget(base_widget)
        main_layout = QVBoxLayout(base_widget)
        main_layout.setContentsMargins(self.shadow_margin, self.shadow_margin, self.shadow_margin, self.shadow_margin)
        
        self.container = QFrame()
        self.container.setObjectName("Container")
        self.container.setStyleSheet("""
            #Container {
                background-color: rgba(255, 255, 255, 245); 
                border-radius: 24px;
                border: 1px solid rgba(255, 255, 255, 180);
            }
        """)
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(0, 0, 0, 40))
        shadow.setYOffset(10)
        self.container.setGraphicsEffect(shadow)
        main_layout.addWidget(self.container)
        
        self.content_layout = QVBoxLayout(self.container)
        self.content_layout.setContentsMargins(30, 25, 30, 30)
        self.content_layout.setSpacing(20)
        
        self.setup_header()
        self.setup_preview_area()
        self.setup_action_buttons()
        self.setup_settings_area()
        self.setup_footer()

    def setup_header(self):
        header = QHBoxLayout()
        titles = QVBoxLayout()
        titles.setSpacing(2)
        
        title = QLabel("Bing Wallpaper")
        title.setFont(QFont("Microsoft YaHei", 20, QFont.Bold))
        title.setStyleSheet("color: #1c1c1e;")
        
        sub_title = QLabel("每日高清壁纸自动同步")
        sub_title.setFont(QFont("Microsoft YaHei", 10))
        sub_title.setStyleSheet("color: #8E8E93;")
        
        titles.addWidget(title)
        titles.addWidget(sub_title)
        
        close_btn = QPushButton("×")
        close_btn.setFixedSize(32, 32)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background: rgba(0,0,0,0.05); color: #888; border-radius: 16px;
                font-family: Arial; font-size: 22px; border: none; padding-bottom: 2px;
            }
            QPushButton:hover { background: #FF3B30; color: white; }
        """)
        close_btn.clicked.connect(self.close)
        
        header.addLayout(titles)
        header.addStretch()
        header.addWidget(close_btn, 0, Qt.AlignTop)
        self.content_layout.addLayout(header)

    def setup_preview_area(self):
        self.preview_container = QFrame()
        
        self.preview_container.setFixedSize(620, 349)
        self.preview_container.setStyleSheet("background: rgba(0,0,0,0.04); border-radius: 16px;")
        
        p_layout = QVBoxLayout(self.preview_container)
        p_layout.setContentsMargins(0,0,0,0)
        
        self.preview_label = QLabel("等待获取...")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("color: #aeaeb2; border-radius: 16px;")
        
        p_layout.addWidget(self.preview_label)
        self.content_layout.addWidget(self.preview_container)

    def setup_action_buttons(self):
        layout = QHBoxLayout()
        layout.setSpacing(15)
        
        self.btn_apply = ModernButton("下载并应用今日壁纸", "#007AFF")
        self.btn_apply.clicked.connect(self.start_manual_download)
        
        self.btn_refresh = ModernButton("刷新预览", "#34C759")
        self.btn_refresh.clicked.connect(self.start_refresh_preview)
        
        layout.addWidget(self.btn_apply)
        layout.addWidget(self.btn_refresh)
        self.content_layout.addLayout(layout)

    def setup_settings_area(self):
        frame = QFrame()
        frame.setStyleSheet("background: rgba(255,255,255,0.6); border-radius: 16px;")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(12)
        
        self.auto_del_chk = ModernCheckBox("自动清理历史壁纸")
        self.auto_start_chk = ModernCheckBox("开机自动启动")
        self.auto_update_chk = ModernCheckBox("自动检查软件更新")
        self.silent_exit_chk = ModernCheckBox("壁纸无更新时不弹出通知") 
        
        self.auto_del_chk.stateChanged.connect(self.save_settings)
        self.auto_start_chk.stateChanged.connect(self.on_autostart_change)
        self.auto_update_chk.stateChanged.connect(self.save_settings)
        self.silent_exit_chk.stateChanged.connect(self.save_settings)
        
        layout.addWidget(self.auto_del_chk)
        layout.addWidget(self.auto_start_chk)
        layout.addWidget(self.auto_update_chk)
        layout.addWidget(self.silent_exit_chk)
        self.content_layout.addWidget(frame)

    def setup_footer(self):
        layout = QHBoxLayout()
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #999; font-size: 12px;")
        
        self.btn_update_chk = QPushButton("检查更新")
        self.btn_update_chk.setCursor(Qt.PointingHandCursor)
        self.btn_update_chk.setStyleSheet("border:none; color:#007AFF; font-weight:bold;")
        self.btn_update_chk.clicked.connect(self.start_check_update)
        
        icon_style = "border:none; padding:4px; border-radius:4px;"
        
        btn_home = QPushButton("") 
        btn_home.setFixedSize(28,28)
        btn_home.setStyleSheet(icon_style)
        btn_home.clicked.connect(lambda: webbrowser.open("https://ombk.xyz"))
        
        home_icon_path = os.path.join(self.images_dir, "home.png")
        if os.path.exists(home_icon_path):
            btn_home.setIcon(QIcon(home_icon_path))
        
        btn_git = QPushButton("") 
        btn_git.setFixedSize(28,28)
        btn_git.setStyleSheet(icon_style)
        btn_git.clicked.connect(lambda: webbrowser.open("https://github.com/QsSama-W/wallpaper-win"))
        
        git_icon_path = os.path.join(self.images_dir, "github.png")
        if os.path.exists(git_icon_path):
            btn_git.setIcon(QIcon(git_icon_path))

        layout.addWidget(self.status_label)
        layout.addStretch()
        layout.addWidget(self.btn_update_chk)
        layout.addWidget(btn_home)
        layout.addWidget(btn_git)
        self.content_layout.addLayout(layout)

    def set_ui_busy(self, busy, msg=""):
        self.is_download_running = busy
        self.btn_apply.setEnabled(not busy)
        self.btn_apply.setText(msg if busy else "下载并应用今日壁纸")
        if busy:
            self.status_label.setText(msg + "...")
        else:
            self.status_label.setText("就绪")

    def start_refresh_preview(self):
        if self.is_preview_running: return 
        self.is_preview_running = True
        self.btn_refresh.setEnabled(False)
        self.btn_refresh.setText("刷新中...")
        self.preview_worker = Worker(self.task_refresh_preview)
        self.preview_worker.signals.finished.connect(self.on_preview_ready)
        self.preview_worker.signals.error.connect(self.on_preview_error)
        self.preview_worker.start()

    def task_refresh_preview(self):
        url = WallpaperUtils.get_bing_url()
        preview_url = url.replace("_UHD.jpg", "_1366x768.jpg")
        resp = requests.get(preview_url, verify=False, timeout=10)
        resp.raise_for_status()
        return resp.content

    def on_preview_ready(self, img_data):
        self.is_preview_running = False
        self.btn_refresh.setEnabled(True)
        self.btn_refresh.setText("刷新预览")
        try:
            image = QImage()
            image.loadFromData(img_data)
            pixmap = QPixmap.fromImage(image)
            
            container_w = self.preview_container.width()
            container_h = self.preview_container.height()
            
            scaled_pix = pixmap.scaled(container_w, container_h, 
                                     Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            
            rounded = QPixmap(scaled_pix.size())
            rounded.fill(Qt.transparent)
            
            painter = QPainter(rounded)
            painter.setRenderHint(QPainter.Antialiasing)
            path = QPainterPath()
            path.addRoundedRect(0, 0, scaled_pix.width(), scaled_pix.height(), 16, 16)
            painter.setClipPath(path)
            painter.drawPixmap(0, 0, scaled_pix)
            painter.end()
            
            self.preview_label.setPixmap(rounded)
            
            if not self.is_download_running:
                self.status_label.setText("预览已更新")
        except Exception as e:
            self.on_preview_error(str(e))
            
    def on_preview_error(self, err_msg):
        self.is_preview_running = False
        self.btn_refresh.setEnabled(True)
        self.btn_refresh.setText("刷新预览")
        self.preview_label.setText("获取预览失败")

    def start_manual_download(self):
        if self.is_download_running: return
        self.set_ui_busy(True, "正在下载")
        self.download_worker = Worker(self.task_download_set, auto_exit=False)
        self.download_worker.signals.finished.connect(self.on_download_success)
        self.download_worker.signals.error.connect(self.on_worker_error)
        self.download_worker.start()

    def start_auto_download(self):
        if self.is_download_running: return 
        
        self.is_download_running = True
        self.status_label.setText("自动运行中...")
        self.download_worker = Worker(self.task_download_set, auto_exit=True)
        self.download_worker.signals.finished.connect(self.on_download_success)
        self.download_worker.signals.error.connect(self.on_auto_error)
        self.download_worker.start()

    def task_download_set(self, auto_exit=False):
        url = WallpaperUtils.get_bing_url()
        today = datetime.date.today().strftime("%Y%m%d")
        filename = f"{today}_UHD.jpg"
        save_path = os.path.join(self.save_dir, filename)
        
        is_new = False
        if not os.path.exists(save_path):
            WallpaperUtils.download_image(url, save_path)
            is_new = True
            
        WallpaperUtils.set_wallpaper_api(save_path)
        
        cleaned = 0
        if self.auto_del_chk.isChecked():
            cleaned = WallpaperUtils.clean_old_wallpapers(self.save_dir)
            
        return {"path": save_path, "cleaned": cleaned, "auto": auto_exit, "is_new": is_new}

    def on_download_success(self, result):
        self.set_ui_busy(False)
        self.status_label.setText("壁纸设置成功")
        if result['auto']:
            self.schedule_exit(is_new=result.get('is_new', True))
        else:
            QMessageBox.information(self, "成功", "今日壁纸已应用到桌面！")

    def on_worker_error(self, err_msg):
        self.set_ui_busy(False)
        self.status_label.setText("错误")
        if self.isVisible():
            QMessageBox.warning(self, "操作失败", str(err_msg))
            
    def on_auto_error(self, err_msg):
        self.is_download_running = False
        self.status_label.setText(f"自动任务失败: {err_msg}")

    def schedule_exit(self, is_new=True):
        self.is_download_running = False
        if is_new:
            self.status_label.setText("任务完成，1分钟后自动退出")
            self.tray_icon.showMessage("每日必应壁纸", "Bing壁纸已更新，程序即将退出", QSystemTrayIcon.Information, 2000)
        else:
            if self.silent_exit_chk.isChecked():
                self.status_label.setText("壁纸已是最新，静默等待自动退出")
            else:
                self.status_label.setText("壁纸已是最新，1分钟后自动退出")
                self.tray_icon.showMessage("每日必应壁纸", "今日壁纸已是最新，程序即将退出", QSystemTrayIcon.Information, 2000)
                
        QTimer.singleShot(60000, self.on_exit)

    def start_check_update(self):
        if not HAS_PACKAGING:
            self.status_label.setText("无法检查更新(缺失库)")
            return
            
        self.status_label.setText("检查更新...")
        self.update_worker = Worker(WallpaperUtils.check_update, self.current_version)
        self.update_worker.signals.finished.connect(self.on_update_checked)
        self.update_worker.start()

    def on_update_checked(self, result):
        has_update, ver, url = result
        if has_update:
            if self.isVisible():
                btn = QMessageBox.question(self, "发现新版本", f"版本 {ver} 可用，是否去下载？")
                if btn == QMessageBox.Yes:
                    webbrowser.open(url)
            else:
                self.tray_icon.showMessage("每日必应壁纸", f"发现新版本 {ver} 已发布，点击查看", QSystemTrayIcon.Information, 5000)
            self.status_label.setText(f"发现新版本: {ver}")
        else:
            self.status_label.setText("当前是最新版本")

    def check_screen_resolution(self):
        geo = QApplication.primaryScreen().geometry()
        return not (geo.width() < 1280 or geo.height() < 720)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and event.y() < 80:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()
            event.accept()
            
    def mouseMoveEvent(self, event):
        if self.m_flag:
            self.move(event.globalPos() - self.m_Position)
            event.accept()
            
    def mouseReleaseEvent(self, event):
        self.m_flag = False

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        
        logo_path = os.path.join(self.images_dir, "logo.png")
        if os.path.exists(logo_path):
            self.tray_icon.setIcon(QIcon(logo_path))
        else:
            self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        
        menu = QMenu()
        menu.setStyleSheet("QMenu{background:#fff; border:1px solid #eee;} QMenu::item{padding:5px 20px; color:#333;}")
        
        menu.addAction("显示主界面", self.showNormal)
        menu.addAction("立即更新壁纸", self.start_manual_download)
        menu.addSeparator()
        menu.addAction("退出", self.on_exit)
        
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.DoubleClick or r == QSystemTrayIcon.Trigger else None)
        self.tray_icon.show()

    def closeEvent(self, e):
        if self.tray_icon.isVisible():
            self.hide()
            self.tray_icon.showMessage("每日必应壁纸", "程序已最小化到托盘", QSystemTrayIcon.Information, 1000)
            e.ignore()
        else: e.accept()
        
    def _get_bool_setting(self, key, default=False):
        val = self.settings.value(key, default)
        if isinstance(val, bool):
            return val
        if isinstance(val, str):
            return val.lower() == 'true'
        return bool(val)

    def load_settings(self):
        self.auto_del_chk.blockSignals(True)
        self.auto_start_chk.blockSignals(True)
        self.auto_update_chk.blockSignals(True)
        self.silent_exit_chk.blockSignals(True)

        self.auto_del_chk.setChecked(self._get_bool_setting("auto_delete", False))
        self.auto_start_chk.setChecked(self._get_bool_setting("auto_start", False))
        self.auto_update_chk.setChecked(self._get_bool_setting("auto_check_update", True))
        self.silent_exit_chk.setChecked(self._get_bool_setting("silent_exit", False))
        
        self.auto_del_chk.blockSignals(False)
        self.auto_start_chk.blockSignals(False)
        self.auto_update_chk.blockSignals(False)
        self.silent_exit_chk.blockSignals(False)

    def save_settings(self):
        self.settings.setValue("auto_delete", self.auto_del_chk.isChecked())
        self.settings.setValue("auto_check_update", self.auto_update_chk.isChecked())
        self.settings.setValue("silent_exit", self.silent_exit_chk.isChecked())

    def on_autostart_change(self):
        self.settings.setValue("auto_start", self.auto_start_chk.isChecked())
        self.set_startup_registry(self.auto_start_chk.isChecked())

    def set_startup_registry(self, enable):
        try:
            import winreg as reg
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            with reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_SET_VALUE) as key:
                if enable:
                    reg.SetValueEx(key, "BingWallpaperManager", 0, reg.REG_SZ, os.path.abspath(sys.argv[0]))
                else:
                    try: reg.DeleteValue(key, "BingWallpaperManager")
                    except: pass
        except: pass

    def on_exit(self):
        self.tray_icon.hide()
        QApplication.quit()

if __name__ == "__main__":
    if sys.platform.startswith('win'):
        try:
            myappid = '每日必应壁纸'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except: pass

    QApplication.setApplicationName("Bing壁纸")
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    
    app = QApplication(sys.argv)
    
    if sys.platform.startswith('win'):
        try:
            import win32event, win32api
            from winerror import ERROR_ALREADY_EXISTS
            mutex = win32event.CreateMutex(None, False, "BingWallpaperMutex_V1.4")
            if win32api.GetLastError() == ERROR_ALREADY_EXISTS: sys.exit(0)
        except: pass
    
    w = BingWallpaperApp()
    sys.exit(app.exec_())
