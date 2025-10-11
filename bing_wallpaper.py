import sys
import os
import datetime
import getpass
import ctypes
import webbrowser
import re
import time
from packaging import version
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QCheckBox, 
                            QGroupBox, QFormLayout, QMessageBox, QSystemTrayIcon,
                            QMenu, QAction, QSpinBox, QComboBox)
from PyQt5.QtCore import Qt, QTimer, QSettings, QSize, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QIcon, QImage, QPixmap
import requests
from requests.exceptions import RequestException


class CronCheckerThread(QThread):
    trigger_update = pyqtSignal()
    status_updated = pyqtSignal(str)
    debug_message = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.running = False
        self.check_interval = 10
        self.debug = True
        
    def run(self):
        self.running = True
        self.debug_message.emit("定时检查线程已启动")
        
        while self.running:
            if self.parent.cron_enabled:
                try:
                    cron_dict = CronParser.parse(self.parent.cron_expression)
                    if cron_dict:
                        now = datetime.datetime.now()
                        next_run = CronParser.get_next_run_time(cron_dict, self.parent.last_triggered)
                        
                        current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")
                        next_run_str = next_run.strftime("%Y-%m-%d %H:%M:%S")
                        debug_info = f"[{current_time_str}] 定时检查 - 下次运行时间: {next_run_str}"
                        self.debug_message.emit(debug_info)
                        
                        time_diff = (now - (next_run - datetime.timedelta(seconds=10))).total_seconds()
                        if time_diff >= 0:
                            trigger_msg = f"[{current_time_str}] 到达定时时间，触发更新..."
                            self.status_updated.emit(trigger_msg)
                            self.debug_message.emit(trigger_msg)
                            
                            self.trigger_update.emit()
                            
                            self.debug_message.emit("等待更新完成...")
                            time.sleep(10)
                        else:
                            delta = next_run - now
                            minutes = delta.seconds // 60
                            hours = minutes // 60
                            minutes = minutes % 60
                            
                            if hours > 0:
                                status_msg = f"定时运行中，下次更新将在{hours}小时{minutes}分钟后"
                            else:
                                status_msg = f"定时运行中，下次更新将在{minutes}分钟后"
                                
                            self.status_updated.emit(status_msg)
                    else:
                        self.debug_message.emit("无效的Cron表达式")
                except Exception as e:
                    error_msg = f"定时检查出错: {str(e)}"
                    self.status_updated.emit(error_msg)
                    self.debug_message.emit(error_msg)
            
            for _ in range(self.check_interval):
                if not self.running:
                    break
                time.sleep(1)
    
    def stop(self):
        self.debug_message.emit("正在停止定时检查线程...")
        self.running = False
        self.wait()
        self.debug_message.emit("定时检查线程已停止")


class CronParser:
    @staticmethod
    def parse(cron_expr):
        try:
            parts = cron_expr.split()
            if len(parts) != 5:
                return None
                
            minute, hour, day, month, weekday = parts
            return {
                'minute': minute,
                'hour': hour,
                'day': day,
                'month': month,
                'weekday': weekday
            }
        except Exception as e:
            print(f"Cron解析错误: {e}")
            return None
    
    @staticmethod
    def get_next_run_time(cron_dict, last_triggered=None):
        try:
            now = datetime.datetime.now()
            next_time = now + datetime.timedelta(minutes=1)
            
            if cron_dict['minute'] == '*' and cron_dict['hour'] == '*' and cron_dict['day'] == '*':
                return now + datetime.timedelta(minutes=1)
            
            elif cron_dict['hour'] == '*' and cron_dict['day'] == '*':
                minute = int(cron_dict['minute'])
                candidate = now.replace(minute=minute, second=0, microsecond=0)
                if candidate > now:
                    return candidate
                return candidate + datetime.timedelta(hours=1)
            
            elif cron_dict['day'] == '*':
                hour = int(cron_dict['hour'])
                minute = int(cron_dict['minute'])
                
                candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                
                if candidate <= now:
                    candidate += datetime.timedelta(days=1)
                
                if last_triggered and candidate <= last_triggered:
                    candidate += datetime.timedelta(days=1)
                
                return candidate
                
            return now + datetime.timedelta(hours=1)
        except Exception as e:
            print(f"计算下次运行时间出错: {e}")
            return now + datetime.timedelta(hours=1)


class BingWallpaperManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_version = "1.1.0"
        self.username = getpass.getuser()
        self.save_dir = os.path.join(f"C:\\Users\\{self.username}", "Pictures", "bing_wallpaper")
        self.settings = QSettings("BingWallpaper", "Manager")
        
        self.target_width = 650
        self.target_height = 780
        
        self.cron_enabled = False
        self.cron_expression = "0 * * * *"
        self.last_triggered = None
        self.last_triggered_date = None
        
        self.cron_thread = CronCheckerThread(self)
        self.cron_thread.trigger_update.connect(self.auto_download_and_set_wallpaper)
        self.cron_thread.status_updated.connect(self.update_status)
        self.cron_thread.debug_message.connect(self.handle_debug_message)
        
        font = self.font()
        font.setFamily("SimHei")
        self.setFont(font)
        
        self.setWindowFlags(
            Qt.WindowTitleHint |
            Qt.WindowCloseButtonHint
        )
        
        if not self.check_dependencies():
            QApplication.quit()
            return
            
        if not self.check_screen_resolution():
            QApplication.quit()
            return
        
        self.check_windows_version()
        self.init_ui()
        self.init_tray()
        self.load_settings()
        os.makedirs(self.save_dir, exist_ok=True)
        
        if self.auto_start_checkbox.isChecked():
            self.set_startup_registry(True)
        
        self.setup_cron_timer()
        
        self.hide()
        self.tray_icon.show()
        
        QTimer.singleShot(3000, self.post_init_operations)

    def check_dependencies(self):
        try:
            from PyQt5 import QtWidgets
            import requests
            return True
        except ImportError as e:
            missing_package = str(e).split("'")[1]
            QMessageBox.critical(
                None, 
                "缺少依赖", 
                f"程序缺少必要的组件: {missing_package}\n"
                f"请运行命令安装: pip install {missing_package}"
            )
            return False

    def post_init_operations(self):
        self.auto_download_and_set_wallpaper()
        if self.auto_check_update_checkbox.isChecked():
            QTimer.singleShot(5000, self.check_for_updates)

    def handle_debug_message(self, message):
        print(f"[debug] {message}")

    def check_screen_resolution(self):
        screen = QApplication.primaryScreen()
        if not screen:
            QMessageBox.critical(None, "错误", "无法检测屏幕信息")
            return False
            
        geometry = screen.geometry()
        width = geometry.width()
        height = geometry.height()
        
        if width < 1920 or height < 1080:
            QMessageBox.critical(
                None, 
                "分辨率不足", 
                f"检测到当前屏幕分辨率为 {width}x{height}\n"
                "本程序需要最低1920x1080的屏幕分辨率才能正常运行。\n"
                "请调整分辨率后重新启动程序。"
            )
            return False
        return True

    def check_windows_version(self):
        required_major = 10
        required_minor = 0
        required_build = 19045
        
        try:
            version = sys.getwindowsversion()
            major = version.major
            minor = version.minor
            build = version.build
            
            if (major < required_major or 
                (major == required_major and minor < required_minor) or 
                (major == required_major and minor == required_minor and build < required_build)):
                
                QMessageBox.warning(
                    self, 
                    "兼容性警告", 
                    "您的Windows版本低于Win10 22H2，程序可能存在兼容性问题。\n"
                    "建议升级到最新版本的Windows以获得最佳体验。"
                )
        except Exception as e:
            QMessageBox.information(
                self, 
                "版本检测", 
                f"无法检测Windows版本: {str(e)}\n程序将继续运行。"
            )

    def init_ui(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.images_dir = os.path.join(base_dir, "images")
        os.makedirs(self.images_dir, exist_ok=True)
        
        try:
            pixmap2 = QPixmap(os.path.join(self.images_dir, "logo.png"))
            icon = QIcon(pixmap2)
            self.setWindowIcon(icon)
        except:
            pass

        self.setWindowTitle(f"Bing壁纸v{self.current_version}")
        self.resize(self.target_width, self.target_height)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        top_layout = QHBoxLayout()
        title_label = QLabel("自动设置今日Bing壁纸为桌面壁纸")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        title_label.setAlignment(Qt.AlignCenter)
        top_layout.addStretch()
        top_layout.addWidget(title_label)
        top_layout.addStretch()
        main_layout.addLayout(top_layout)
        
        preview_group = QGroupBox("今日壁纸预览")
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumHeight(200)
        self.preview_label.setStyleSheet("border: 1px solid #ddd; border-radius: 4px;")
        preview_layout.addWidget(self.preview_label)
        preview_group.setLayout(preview_layout)
        main_layout.addWidget(preview_group)
        
        button_layout = QHBoxLayout()
        self.download_btn = QPushButton("下载并设置壁纸")
        self.download_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.download_btn.clicked.connect(self.download_and_set_wallpaper)
        
        self.refresh_btn = QPushButton("刷新预览")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        self.refresh_btn.clicked.connect(self.refresh_preview)
        
        button_layout.addWidget(self.download_btn)
        button_layout.addWidget(self.refresh_btn)
        main_layout.addLayout(button_layout)
        
        cron_group = QGroupBox("定时更新壁纸设置")
        cron_layout = QFormLayout()
        
        self.cron_enable_checkbox = QCheckBox("启用定时更新壁纸")
        self.cron_enable_checkbox.stateChanged.connect(self.on_cron_enabled_changed)
        
        cron_time_layout = QHBoxLayout()
        
        self.minute_spin = QSpinBox()
        self.minute_spin.setRange(0, 59)
        self.minute_spin.setValue(0)
        
        minute_label = QLabel("分")
        
        self.hour_spin = QSpinBox()
        self.hour_spin.setRange(0, 23)
        self.hour_spin.setValue(0)
        
        hour_label = QLabel("时")
        
        self.interval_combo = QComboBox()
        self.interval_combo.addItems(["每天"])
        
        self.cron_confirm_btn = QPushButton("确定")
        self.cron_confirm_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                padding: 4px 12px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.cron_confirm_btn.clicked.connect(self.update_cron_expression)
        
        cron_time_layout.addWidget(self.minute_spin)
        cron_time_layout.addWidget(minute_label)
        cron_time_layout.addWidget(self.hour_spin)
        cron_time_layout.addWidget(hour_label)
        cron_time_layout.addWidget(self.interval_combo)
        cron_time_layout.addWidget(self.cron_confirm_btn)
        cron_time_layout.addStretch()
        
        self.cron_expression_label = QLabel("当前定时规则: 0 0 * * * (每天0点0分)")
        
        cron_layout.addRow(self.cron_enable_checkbox)
        cron_layout.addRow("定时时间:", cron_time_layout)
        cron_layout.addRow(self.cron_expression_label)
        cron_group.setLayout(cron_layout)
        main_layout.addWidget(cron_group)
        
        settings_group = QGroupBox("设置")
        settings_layout = QFormLayout()
        
        self.auto_delete_checkbox = QCheckBox("开启自动删除历史壁纸")
        self.auto_delete_checkbox.stateChanged.connect(self.save_settings)
        
        self.auto_start_checkbox = QCheckBox("开机自动启动")
        self.auto_start_checkbox.stateChanged.connect(self.on_auto_start_changed)
        
        self.auto_check_update_checkbox = QCheckBox("开启自动检查更新")
        self.auto_check_update_checkbox.stateChanged.connect(self.save_settings)
        
        settings_layout.addRow(self.auto_delete_checkbox)
        settings_layout.addRow(self.auto_start_checkbox)
        settings_layout.addRow(self.auto_check_update_checkbox)
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        bottom_layout = QHBoxLayout()
        self.status_label = QLabel("就绪")
        bottom_layout.addWidget(self.status_label)
        bottom_layout.addStretch()
        
        self.home_btn = QPushButton()
        self.home_btn.setFixedSize(27, 27)
        self.home_btn.setStyleSheet("""
            QPushButton {
                border-radius: 6px;
                background-color: white;
                border: 1px solid #ddd;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
                border-color: #bbb;
            }
            QPushButton:pressed {
                background-color: #e0e0e0;
            }
        """)
        
        home_icon = QIcon(os.path.join(self.images_dir, "home.png"))
        self.home_btn.setIcon(home_icon)
        self.home_btn.setIconSize(QSize(23, 23))
        self.home_btn.clicked.connect(lambda: webbrowser.open("https://ombk.xyz"))
        
        self.github_btn = QPushButton()
        self.github_btn.setFixedSize(27, 27)
        self.github_btn.setStyleSheet("""
            QPushButton {
                border-radius: 6px;
                background-color: white;
                border: 1px solid #ddd;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
                border-color: #bbb;
            }
            QPushButton:pressed {
                background-color: #e0e0e0;
            }
        """)
        
        github_icon = QIcon(os.path.join(self.images_dir, "github.png"))
        self.github_btn.setIcon(github_icon)
        self.github_btn.setIconSize(QSize(25, 25))
        self.github_btn.clicked.connect(lambda: webbrowser.open("https://github.com/QsSama-W/wallpaper-win"))
        
        self.check_update_btn = QPushButton("检查更新")
        self.check_update_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                border-radius: 4px;
                padding: 6px 12px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        self.check_update_btn.clicked.connect(self.check_for_updates)
        
        bottom_layout.addWidget(self.home_btn)
        bottom_layout.addWidget(self.github_btn)
        bottom_layout.addWidget(self.check_update_btn)
        
        main_layout.addLayout(bottom_layout)
        QTimer.singleShot(1000, self.refresh_preview)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        current_width = self.width()
        current_height = self.height()
        
        if current_width != self.target_width or current_height != self.target_height:
            QTimer.singleShot(100, self.resetSize)
            
    def resetSize(self):
        self.resize(self.target_width, self.target_height)

    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        icon_path = os.path.join(self.images_dir, "logo.png")
        if os.path.exists(icon_path):
            icon = QIcon(QPixmap(icon_path))
            self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("Bing壁纸")
        
        tray_menu = QMenu(self)
        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)
        
        update_action = QAction("更新壁纸", self)
        update_action.triggered.connect(self.download_and_set_wallpaper)
        tray_menu.addAction(update_action)
        
        self.cron_menu_action = QAction("启用定时更新", self, checkable=True)
        self.cron_menu_action.toggled.connect(self.on_cron_menu_toggled)
        tray_menu.addAction(self.cron_menu_action)
        
        check_update_action = QAction("检查更新", self)
        check_update_action.triggered.connect(self.check_for_updates)
        tray_menu.addAction(check_update_action)
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.on_exit)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()

    def load_settings(self):
        auto_delete = self.settings.value("auto_delete", False, type=bool)
        auto_start = self.settings.value("auto_start", False, type=bool)
        auto_check_update = self.settings.value("auto_check_update", True, type=bool)
        
        self.cron_enabled = self.settings.value("cron_enabled", False, type=bool)
        self.cron_expression = self.settings.value("cron_expression", "0 * * * *")
        
        cron_dict = CronParser.parse(self.cron_expression)
        if cron_dict:
            try:
                self.minute_spin.setValue(int(cron_dict['minute']))
            except:
                pass
            try:
                self.hour_spin.setValue(int(cron_dict['hour']))
            except:
                pass
                
            if cron_dict['hour'] == '*':
                self.interval_combo.setCurrentIndex(1)
            else:
                self.interval_combo.setCurrentIndex(0)
        
        self.auto_delete_checkbox.setChecked(auto_delete)
        self.auto_start_checkbox.setChecked(auto_start)
        self.auto_check_update_checkbox.setChecked(auto_check_update)
        self.cron_enable_checkbox.setChecked(self.cron_enabled)
        self.cron_menu_action.setChecked(self.cron_enabled)
        
        self.update_cron_expression_label()

    def save_settings(self):
        self.settings.setValue("auto_delete", self.auto_delete_checkbox.isChecked())
        self.settings.setValue("auto_check_update", self.auto_check_update_checkbox.isChecked())
        self.settings.setValue("cron_enabled", self.cron_enabled)
        self.settings.setValue("cron_expression", self.cron_expression)
        self.settings.sync()

    def update_status(self, message):
        self.status_label.setText(message)

    def on_auto_start_changed(self):
        self.settings.setValue("auto_start", self.auto_start_checkbox.isChecked())
        self.settings.sync()
        self.set_startup_registry(self.auto_start_checkbox.isChecked())

    def set_startup_registry(self, enable):
        try:
            import winreg as reg
            app_path = sys.argv[0]
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            
            with reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_SET_VALUE) as key:
                if enable:
                    reg.SetValueEx(key, "BingWallpaperManager", 0, reg.REG_SZ, app_path)
                else:
                    try:
                        reg.DeleteValue(key, "BingWallpaperManager")
                    except FileNotFoundError:
                        pass
            return True
        except Exception as e:
            QMessageBox.warning(self, "错误", f"设置开机启动失败: {str(e)}")
            return False

    def get_bing_wallpaper_url(self):
        try:
            url = "https://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
            response = requests.get(url, verify=False, timeout=10)
            response.raise_for_status()
            data = response.json()
            imgurl_base = data['images'][0]['urlbase']
            return f"https://cn.bing.com{imgurl_base}_UHD.jpg"
        except (RequestException, KeyError, IndexError) as e:
            self.update_status(f"获取壁纸URL失败: {str(e)}")
            return None

    def refresh_preview(self):
        self.update_status("正在加载预览图...")
        url = self.get_bing_wallpaper_url()
        if not url:
            return
            
        try:
            preview_url = url.replace("_UHD.jpg", "_1366x768.jpg")
            response = requests.get(preview_url, verify=False, stream=True, timeout=10)
            response.raise_for_status()
            
            image = QImage()
            image.loadFromData(response.content)
            pixmap = QPixmap.fromImage(image)
            
            scaled_pixmap = pixmap.scaled(
                self.preview_label.width(), 
                self.preview_label.height(), 
                Qt.KeepAspectRatio, 
                Qt.SmoothTransformation
            )
            
            self.preview_label.setPixmap(scaled_pixmap)
            self.update_status("预览图加载完成")
        except Exception as e:
            self.update_status(f"预览图加载失败: {str(e)}")

    def delete_old_wallpapers(self):
        try:
            today = datetime.date.today()
            for filename in os.listdir(self.save_dir):
                if filename.endswith("_UHD.jpg"):
                    try:
                        date_str = filename.split("_")[0]
                        file_date = datetime.datetime.strptime(date_str, "%Y%m%d").date()
                        
                        if file_date != today:
                            file_path = os.path.join(self.save_dir, filename)
                            os.remove(file_path)
                            print(f"已删除历史壁纸: {filename}")
                    except (ValueError, OSError) as e:
                        print(f"删除文件 {filename} 失败: {str(e)}")
        except Exception as e:
            print(f"自动删除历史壁纸出错: {str(e)}")

    def auto_download_and_set_wallpaper(self):
        self.update_status("自动获取今日壁纸...")
        
        QTimer.singleShot(0, self._perform_auto_download)

    def _perform_auto_download(self):
        try:
            wallpaper_url = self.get_bing_wallpaper_url()
            if not wallpaper_url:
                return
                
            today = datetime.date.today().strftime("%Y%m%d")
            filename = f"{today}_UHD.jpg"
            save_path = os.path.join(self.save_dir, filename)
            
            if os.path.exists(save_path):
                self.update_status("今日壁纸已存在，无需重复下载")
                if self.set_wallpaper(save_path):
                    self.update_status("已使用现有壁纸更新桌面")
                self.last_triggered = datetime.datetime.now()
                self.last_triggered_date = datetime.date.today()
                return
            
            self.update_status("正在下载今日壁纸...")
            response = requests.get(wallpaper_url, verify=False, stream=True, timeout=30)
            response.raise_for_status()
            
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
            
            self.update_status("壁纸下载完成")
            
            if self.auto_delete_checkbox.isChecked():
                self.delete_old_wallpapers()
            
            if self.set_wallpaper(save_path):
                self.update_status("壁纸已自动设置为桌面背景")
                self.refresh_preview()
                self.tray_icon.showMessage(
                    "成功", 
                    "Bing每日壁纸已更新",
                    QSystemTrayIcon.Information,
                    3000
                )
                self.last_triggered = datetime.datetime.now()
                self.last_triggered_date = datetime.date.today()
            else:
                self.update_status("自动设置壁纸失败")
                
        except Exception as e:
            error_msg = f"自动更新失败: {str(e)}"
            self.update_status(error_msg)
            self.tray_icon.showMessage(
                "错误", 
                error_msg,
                QSystemTrayIcon.Critical,
                3000
            )

    def closeEvent(self, event):
        if self.tray_icon.isVisible():
            self.hide()
            self.tray_icon.showMessage(
                "提示",
                "程序已最小化到系统托盘",
                QSystemTrayIcon.Information,
                2000
            )
            event.ignore()
        else:
            event.accept()

    def on_exit(self):
        self.cron_thread.stop()
        self.tray_icon.hide()
        QApplication.quit()

    def check_for_updates(self):
        if not self.auto_check_update_checkbox.isChecked() and self.sender() != self.check_update_btn:
            return
            
        self.update_status("正在检查更新...")
        
        QTimer.singleShot(0, self._perform_update_check)

    def _perform_update_check(self):
        api_url = "https://api.github.com/repos/QsSama-W/wallpaper-win/releases/latest"
        
        try:
            response = requests.get(api_url, verify=True, timeout=10)
            response.raise_for_status()
            release_info = response.json()
            
            latest_version_tag = release_info.get('tag_name', '')
            latest_version = re.sub(r'^v', '', latest_version_tag)
            
            if version.parse(latest_version) > version.parse(self.current_version):
                self.show_update_message(latest_version, release_info.get('html_url'))
            else:
                self.update_status("当前已是最新版本")
                
        except RequestException:
            self.update_status("更新检查失败: 网络错误")
        except (KeyError, ValueError):
            self.update_status("更新检查失败: 解析错误")
        except Exception as e:
            self.update_status(f"更新检查失败: {str(e)}")

    def show_update_message(self, latest_version, release_url):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("发现新版本")
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(f"发现新版本 {latest_version}！\n当前版本: {self.current_version}")
        msg_box.setInformativeText("是否前往GitHub查看更新？")
        
        update_btn = msg_box.addButton("前往更新", QMessageBox.AcceptRole)
        later_btn = msg_box.addButton("稍后再说", QMessageBox.RejectRole)
        
        msg_box.exec_()
        
        if msg_box.clickedButton() == update_btn:
            webbrowser.open(release_url)
        self.update_status(f"发现新版本 {latest_version}")

    def download_and_set_wallpaper(self):
        self.update_status("正在获取壁纸信息...")
        
        QTimer.singleShot(0, self._perform_manual_download)

    def _perform_manual_download(self):
        try:
            wallpaper_url = self.get_bing_wallpaper_url()
            if not wallpaper_url:
                return
                
            today = datetime.date.today().strftime("%Y%m%d")
            filename = f"{today}_UHD.jpg"
            save_path = os.path.join(self.save_dir, filename)
            
            if os.path.exists(save_path):
                self.update_status("壁纸已存在，直接设置为桌面背景...")
            else:
                self.update_status("正在下载壁纸...")
                response = requests.get(wallpaper_url, verify=False, stream=True, timeout=30)
                response.raise_for_status()
                
                with open(save_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                
                self.update_status("壁纸下载完成")
            
            if self.auto_delete_checkbox.isChecked():
                self.delete_old_wallpapers()
            
            if self.set_wallpaper(save_path):
                self.update_status("壁纸设置成功！")
                QMessageBox.information(self, "成功", "壁纸已成功设置为桌面背景")
                self.refresh_preview()
                self.last_triggered = datetime.datetime.now()
                self.last_triggered_date = datetime.date.today()
            else:
                self.update_status("壁纸设置失败")
                QMessageBox.warning(self, "失败", "无法设置桌面壁纸")
                
        except Exception as e:
            error_msg = f"操作失败: {str(e)}"
            self.update_status(error_msg)
            QMessageBox.critical(self, "错误", error_msg)

    def set_wallpaper(self, image_path):
        try:
            SPI_SETDESKWALLPAPER = 20
            SPIF_UPDATEINIFILE = 1
            result = ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETDESKWALLPAPER, 
                0, 
                image_path, 
                SPIF_UPDATEINIFILE
            )
            return result != 0
        except Exception as e:
            print(f"设置壁纸失败: {str(e)}")
            return False

    def on_cron_enabled_changed(self, state):
        self.cron_enabled = state == Qt.Checked
        self.cron_menu_action.setChecked(self.cron_enabled)
        self.save_settings()
        self.setup_cron_timer()
        
    def on_cron_menu_toggled(self, checked):
        self.cron_enabled = checked
        self.cron_enable_checkbox.setChecked(checked)
        self.save_settings()
        self.setup_cron_timer()
    
    def update_cron_expression(self):
        minute = self.minute_spin.value()
        hour = self.hour_spin.value()
        interval = self.interval_combo.currentIndex()
        
        if interval == 1:
            self.cron_expression = f"{minute} * * * *"
        else:
            self.cron_expression = f"{minute} {hour} * * *"
            
        self.update_cron_expression_label()
        self.save_settings()
        if self.cron_enabled:
            self.setup_cron_timer()
    
    def update_cron_expression_label(self):
        cron_dict = CronParser.parse(self.cron_expression)
        if not cron_dict:
            return
            
        description = ""
        if cron_dict['hour'] == '*':
            description = f"每小时{int(cron_dict['minute'])}分"
        else:
            description = f"每天{int(cron_dict['hour'])}点{int(cron_dict['minute'])}分"
            
        self.cron_expression_label.setText(f"当前定时规则: {self.cron_expression} ({description})")
    
    def setup_cron_timer(self):
        if self.cron_thread.isRunning():
            self.cron_thread.stop()
            
        if self.cron_enabled:
            self.cron_thread.start()
            self.update_status(f"定时更新已启用: {self.cron_expression}")
        else:
            self.update_status("定时更新已禁用")


if __name__ == "__main__":
    QApplication.setApplicationName("Bing壁纸")
    app = QApplication(sys.argv)
    
    font = app.font()
    font.setFamily("SimHei")
    app.setFont(font)
    
    if sys.platform.startswith('win'):
        import win32event
        import win32api
        from winerror import ERROR_ALREADY_EXISTS
        
        mutex = win32event.CreateMutex(None, False, "BingWallpaperManagerMutex")
        if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
            QMessageBox.information(None, "提示", "程序已在运行中！")
            sys.exit(0)
    
    window = BingWallpaperManager()
    sys.exit(app.exec_())
