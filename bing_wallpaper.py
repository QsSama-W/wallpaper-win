import sys
import os
import datetime
import getpass
import ctypes
import webbrowser
import re
from packaging import version
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QCheckBox, 
                            QGroupBox, QFormLayout, QMessageBox, QSystemTrayIcon,
                            QMenu, QAction)
from PyQt5.QtCore import Qt, QTimer, QSettings, QSize
from PyQt5.QtGui import QIcon, QPixmap, QImage
import requests
from requests.exceptions import RequestException

class BingWallpaperManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_version = "1.0.1"
        self.username = getpass.getuser()
        self.save_dir = os.path.join(f"C:\\Users\\{self.username}", "Pictures", "bing_wallpaper")
        self.settings = QSettings("BingWallpaper", "Manager")
        
        self.target_width = 650
        self.target_height = 675
        
        # 禁用最小化和最大化按钮
        self.setWindowFlags(
            Qt.WindowTitleHint |
            Qt.WindowCloseButtonHint
        )
        
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
        
        self.hide()
        self.tray_icon.show()
        
        QTimer.singleShot(3000, self.auto_download_and_set_wallpaper)
        if self.auto_check_update_checkbox.isChecked():
            QTimer.singleShot(5000, self.check_for_updates)

    def check_screen_resolution(self):
        """检查屏幕分辨率是否不低于1920*1080"""
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
        """检查Windows版本是否高于Win10 22H2"""
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
        """初始化用户界面"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.images_dir = os.path.join(base_dir, "images")
        os.makedirs(self.images_dir, exist_ok=True)
        
        try:
            pixmap2 = QPixmap(os.path.join(self.images_dir, "logo.png"))
            icon = QIcon(pixmap2)
            self.setWindowIcon(icon)
        except:
            pass

        self.setWindowTitle("Bing壁纸v1.0.1")
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
        """窗口大小改变事件处理"""
        super().resizeEvent(event)
        current_width = self.width()
        current_height = self.height()
        
        if current_width != self.target_width or current_height != self.target_height:
            QTimer.singleShot(100, self.resetSize)
            
    def resetSize(self):
        """重置窗口到目标尺寸"""
        self.resize(self.target_width, self.target_height)

    def init_tray(self):
        """初始化系统托盘"""
        self.tray_icon = QSystemTrayIcon(self)
        icon_path = os.path.join(self.images_dir, "logo.png")
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
        
        check_update_action = QAction("检查更新", self)
        check_update_action.triggered.connect(self.check_for_updates)
        tray_menu.addAction(check_update_action)
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.on_exit)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        """托盘图标激活事件处理"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()

    def load_settings(self):
        """加载保存的设置"""
        auto_delete = self.settings.value("auto_delete", False, type=bool)
        auto_start = self.settings.value("auto_start", False, type=bool)
        auto_check_update = self.settings.value("auto_check_update", True, type=bool)
        
        self.auto_delete_checkbox.setChecked(auto_delete)
        self.auto_start_checkbox.setChecked(auto_start)
        self.auto_check_update_checkbox.setChecked(auto_check_update)

    def save_settings(self):
        """保存设置"""
        self.settings.setValue("auto_delete", self.auto_delete_checkbox.isChecked())
        self.settings.setValue("auto_check_update", self.auto_check_update_checkbox.isChecked())
        self.settings.sync()

    def on_auto_start_changed(self):
        """处理开机启动设置变化"""
        self.settings.setValue("auto_start", self.auto_start_checkbox.isChecked())
        self.settings.sync()
        self.set_startup_registry(self.auto_start_checkbox.isChecked())

    def set_startup_registry(self, enable):
        """设置开机启动注册表项"""
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
        """获取Bing壁纸URL"""
        try:
            url = "https://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN"
            response = requests.get(url, verify=False)
            response.raise_for_status()
            data = response.json()
            imgurl_base = data['images'][0]['urlbase']
            return f"https://cn.bing.com{imgurl_base}_UHD.jpg"
        except (RequestException, KeyError, IndexError) as e:
            self.status_label.setText(f"获取壁纸URL失败: {str(e)}")
            return None

    def refresh_preview(self):
        """刷新预览图"""
        self.status_label.setText("正在加载预览图...")
        url = self.get_bing_wallpaper_url()
        if not url:
            return
            
        try:
            preview_url = url.replace("_UHD.jpg", "_1366x768.jpg")
            response = requests.get(preview_url, verify=False, stream=True)
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
            self.status_label.setText("预览图加载完成")
        except Exception as e:
            self.status_label.setText(f"预览图加载失败: {str(e)}")

    def delete_old_wallpapers(self):
        """删除旧壁纸"""
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
        """自动下载并设置壁纸，不显示弹窗"""
        self.status_label.setText("程序启动，自动获取今日壁纸...")
        
        wallpaper_url = self.get_bing_wallpaper_url()
        if not wallpaper_url:
            return
            
        today = datetime.date.today().strftime("%Y%m%d")
        filename = f"{today}_UHD.jpg"
        save_path = os.path.join(self.save_dir, filename)
        
        try:
            if os.path.exists(save_path):
                self.status_label.setText("今日壁纸已存在，无需重复下载")
                if self.set_wallpaper(save_path):
                    self.status_label.setText("已使用现有壁纸更新桌面")
                return
            
            self.status_label.setText("正在下载今日壁纸...")
            response = requests.get(wallpaper_url, verify=False, stream=True)
            response.raise_for_status()
            
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
            
            self.status_label.setText("壁纸下载完成")
            
            if self.auto_delete_checkbox.isChecked():
                self.delete_old_wallpapers()
            
            if self.set_wallpaper(save_path):
                self.status_label.setText("壁纸已自动设置为桌面背景")
                self.refresh_preview()
                self.tray_icon.showMessage(
                    "成功", 
                    "Bing每日壁纸已更新",
                    QSystemTrayIcon.Information,
                    3000
                )
            else:
                self.status_label.setText("自动设置壁纸失败")
                
        except Exception as e:
            error_msg = f"自动更新失败: {str(e)}"
            self.status_label.setText(error_msg)
            self.tray_icon.showMessage(
                "错误", 
                error_msg,
                QSystemTrayIcon.Critical,
                3000
            )

    def closeEvent(self, event):
        """重写关闭事件，将窗口隐藏到托盘而不是退出程序"""
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
        """退出程序"""
        self.tray_icon.hide()
        QApplication.quit()

    def check_for_updates(self):
        """检查是否有新版本发布"""
        if not self.auto_check_update_checkbox.isChecked() and self.sender() != self.check_update_btn:
            return
            
        self.status_label.setText("正在检查更新...")
        
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
                self.status_label.setText("当前已是最新版本")
                
        except RequestException:
            self.status_label.setText("更新检查失败: 网络错误")
        except (KeyError, ValueError):
            self.status_label.setText("更新检查失败: 解析错误")
        except Exception as e:
            self.status_label.setText(f"更新检查失败: {str(e)}")

    def show_update_message(self, latest_version, release_url):
        """显示更新提示对话框"""
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
        self.status_label.setText(f"发现新版本 {latest_version}")

    def download_and_set_wallpaper(self):
        """手动下载并设置壁纸，带弹窗提示"""
        self.status_label.setText("正在获取壁纸信息...")
        
        wallpaper_url = self.get_bing_wallpaper_url()
        if not wallpaper_url:
            return
            
        today = datetime.date.today().strftime("%Y%m%d")
        filename = f"{today}_UHD.jpg"
        save_path = os.path.join(self.save_dir, filename)
        
        try:
            if os.path.exists(save_path):
                self.status_label.setText("壁纸已存在，直接设置为桌面背景...")
            else:
                self.status_label.setText("正在下载壁纸...")
                response = requests.get(wallpaper_url, verify=False, stream=True)
                response.raise_for_status()
                
                with open(save_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                
                self.status_label.setText("壁纸下载完成")
            
            if self.auto_delete_checkbox.isChecked():
                self.delete_old_wallpapers()
            
            if self.set_wallpaper(save_path):
                self.status_label.setText("壁纸设置成功！")
                QMessageBox.information(self, "成功", "壁纸已成功设置为桌面背景")
                self.refresh_preview()
            else:
                self.status_label.setText("壁纸设置失败")
                QMessageBox.warning(self, "失败", "无法设置桌面壁纸")
                
        except Exception as e:
            error_msg = f"操作失败: {str(e)}"
            self.status_label.setText(error_msg)
            QMessageBox.critical(self, "错误", error_msg)

    def set_wallpaper(self, image_path):
        """设置桌面壁纸"""
        try:
            SPI_SETDESKWALLPAPER = 20
            SPIF_UPDATEINIFILE = 1
            ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETDESKWALLPAPER, 
                0, 
                image_path, 
                SPIF_UPDATEINIFILE
            )
            return True
        except Exception as e:
            print(f"设置壁纸失败: {str(e)}")
            return False


if __name__ == "__main__":
    QApplication.setApplicationName("Bing壁纸")
    app = QApplication(sys.argv)
    
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
    
