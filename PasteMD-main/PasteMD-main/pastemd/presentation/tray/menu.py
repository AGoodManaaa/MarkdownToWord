"""Tray menu construction and callbacks."""

import os
import pystray
import threading
import webbrowser

from ... import __version__
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.paths import get_log_path, get_config_path
from ...domains.notification.manager import NotificationManager
from ...utils.fs import ensure_dir
from ...utils.logging import log
from ...utils.version_checker import VersionChecker
from ...i18n import t, iter_languages, get_language, set_language, get_language_label
from .icon import create_status_icon
from ..hotkey.dialog import HotkeyDialog
from ..settings.dialog import SettingsDialog


class TrayMenuManager:
    """托盘菜单管理器"""
    
    def __init__(self, config_loader: ConfigLoader, notification_manager: NotificationManager):
        self.config_loader = config_loader
        self.notification_manager = notification_manager
        self.restart_hotkey_callback = None  # 将由外部设置
        self.pause_hotkey_callback = None  # 暂停热键监听
        self.resume_hotkey_callback = None  # 恢复热键监听
        self.version_checker = None  # 将由外部设置或按需创建
        self.latest_version = None  # 存储最新版本号
        self.latest_release_url = None  # 存储最新版本的下载链接
        self.hotkey_dialog = None
        self.settings_dialog = None
    
    def set_restart_hotkey_callback(self, callback):
        """设置重启热键的回调函数"""
        self.restart_hotkey_callback = callback
    
    def set_pause_hotkey_callback(self, callback):
        """设置暂停热键的回调函数"""
        self.pause_hotkey_callback = callback
    
    def set_resume_hotkey_callback(self, callback):
        """设置恢复热键的回调函数"""
        self.resume_hotkey_callback = callback
    
    def build_menu(self) -> pystray.Menu:
        """构建托盘菜单"""
        config = app_state.config

        # 构建版本菜单项
        version_menu_items = [
            pystray.MenuItem(
                t("tray.menu.current_version", version=__version__),
                lambda icon, item: None,
                enabled=False
            ),
        ]
        if self.latest_version:
            version_menu_items.append(
                pystray.MenuItem(
                    t("tray.menu.new_version", version=self.latest_version),
                    self._on_open_release_page,
                    enabled=True
                )
            )
        else:
            version_menu_items.append(
                pystray.MenuItem(
                    t("tray.menu.check_update"),
                    self._on_check_update
                )
            )

        html_formatting_menu = self._build_html_formatting_menu()

        return pystray.Menu(
            pystray.MenuItem(
                t("tray.menu.hotkey_display", hotkey=app_state.config['hotkey']),
                lambda icon, item: None,
                enabled=False
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                t("tray.menu.enable_hotkey"),
                self._on_toggle_enabled,
                checked=lambda item: app_state.enabled
            ),
            pystray.MenuItem(
                t("tray.menu.show_notifications"),
                self._on_toggle_notify,
                checked=lambda item: config.get("notify", True)
            ),
            pystray.MenuItem(
                t("tray.menu.auto_open"),
                self._on_toggle_auto_open,
                checked=lambda item: config.get("auto_open_on_no_app", True)
            ),
            pystray.MenuItem(
                t("tray.menu.move_cursor"),
                self._on_toggle_move_cursor,
                checked=lambda item: config.get("move_cursor_to_end", True)
            ),
            html_formatting_menu,
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("tray.menu.set_hotkey"), self._on_set_hotkey),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                t("tray.menu.keep_file"),
                self._on_toggle_keep,
                checked=lambda item: config.get("keep_file", False)
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("tray.menu.open_save_dir"), self._on_open_save_dir),
            pystray.MenuItem(t("tray.menu.open_log"), self._on_open_log),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("settings.dialog.title"), self._on_open_settings),
            pystray.Menu.SEPARATOR,
            *version_menu_items,
            pystray.MenuItem(
                t("tray.menu.about"),
                self._on_open_about_page
            ),
            pystray.MenuItem(t("tray.menu.quit"), self._on_quit)
        )

    # 菜单回调函数
    def _on_toggle_enabled(self, icon, item):
        """切换热键启用状态"""
        app_state.enabled = not app_state.enabled
        icon.icon = create_status_icon(ok=app_state.enabled)
        
        status = t("tray.status.hotkey_enabled") if app_state.enabled else t("tray.status.hotkey_paused")
        icon.menu = self.build_menu()
        self.notification_manager.notify("PasteMD", status, ok=app_state.enabled)
    
    def _on_set_hotkey(self, icon, item):
        """设置热键，确保 Tk 对话框在主线程 UI 队列中运行"""
        def save_hotkey(new_hotkey: str):
            """保存新热键并重启热键绑定"""
            try:
                # 更新配置
                app_state.config["hotkey"] = new_hotkey
                app_state.hotkey_str = new_hotkey
                self._save_config()
                
                # 重启热键绑定
                if self.restart_hotkey_callback:
                    self.restart_hotkey_callback()
                
                # 刷新菜单
                icon.menu = self.build_menu()
                
                log(f"Hotkey changed to: {new_hotkey}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.status.hotkey_saved", hotkey=new_hotkey),
                    ok=True)
            except Exception as e:
                log(f"Failed to save hotkey: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.error.hotkey_save_failed", error=str(e)),
                    ok=False)
                raise

        def show_dialog_on_main():
            """在主线程打开/销毁 Tk 对话框"""
            paused = False
            try:
                if self.hotkey_dialog and self.hotkey_dialog.is_alive():
                    self.hotkey_dialog.restore_and_focus()
                    return
                
                # 暂停全局热键监听（避免录制时触发）
                if self.pause_hotkey_callback:
                    self.pause_hotkey_callback()
                    paused = True
                
                dialog = HotkeyDialog(
                    current_hotkey=app_state.hotkey_str,
                    on_save=save_hotkey,
                    on_close=self.resume_hotkey_callback  # 关闭对话框时恢复监听
                )
                self.hotkey_dialog = dialog

                def _clear_dialog_ref(event=None):
                    if getattr(event, "widget", None) is dialog.root or event is None:
                        self.hotkey_dialog = None

                dialog.root.bind("<Destroy>", _clear_dialog_ref)
                dialog.show()
            except Exception as e:
                log(f"Failed to show hotkey dialog: {e}")
                self.notification_manager.notify("PasteMD", t("tray.error.open_hotkey_dialog", error=str(e)), ok=False)
            finally:
                # 确保恢复热键监听
                if paused and self.resume_hotkey_callback:
                    self.resume_hotkey_callback()

        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(show_dialog_on_main)
        else:
            # 兜底：未获取到 UI 队列时，仍在当前线程执行
            show_dialog_on_main()
    
    def _on_toggle_notify(self, icon, item):
        """切换通知状态"""
        current = app_state.config.get("notify", True)
        app_state.config["notify"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        if app_state.config["notify"]:
            self.notification_manager.notify("PasteMD", t("tray.status.notifications_enabled"), ok=True)
        else:
            log("Notifications disabled via tray toggle")
    
    def _on_toggle_auto_open(self, icon, item):
        """切换无应用时自动打开状态"""
        current = app_state.config.get("auto_open_on_no_app", True)
        app_state.config["auto_open_on_no_app"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.auto_open_on") if app_state.config["auto_open_on_no_app"] else t("tray.status.auto_open_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_toggle_move_cursor(self, icon, item):
        """切换插入后光标移动到末尾状态"""
        current = app_state.config.get("move_cursor_to_end", True)
        app_state.config["move_cursor_to_end"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.move_cursor_on") if app_state.config["move_cursor_to_end"] else t("tray.status.move_cursor_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
        
    def _on_toggle_excel(self, icon, item):
        """切换启用 Excel 插入"""
        current = app_state.config.get("enable_excel", True)
        app_state.config["enable_excel"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.excel_insert_on") if app_state.config["enable_excel"] else t("tray.status.excel_insert_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
        
    def _on_toggle_excel_format(self, icon, item):
        """切换 Excel 粘贴时是否保留格式"""
        current = app_state.config.get("excel_keep_format", True)
        app_state.config["excel_keep_format"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.excel_format_on") if app_state.config["excel_keep_format"] else t("tray.status.excel_format_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_toggle_keep(self, icon, item):
        """切换保留文件状态"""
        current = app_state.config.get("keep_file", False)
        app_state.config["keep_file"] = not current
        self._save_config()
        icon.menu = self.build_menu()
        status = t("tray.status.keep_file_on") if app_state.config["keep_file"] else t("tray.status.keep_file_off")
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_open_save_dir(self, icon, item):
        """打开保存目录"""
        save_dir = app_state.config.get("save_dir", "")
        save_dir = os.path.expandvars(save_dir)
        ensure_dir(save_dir)
        os.startfile(save_dir)
    
    def _on_open_log(self, icon, item):
        """打开日志文件"""
        log_path = get_log_path()
        if not os.path.exists(log_path):
            # 创建空日志文件
            open(log_path, "w", encoding="utf-8").close()
        os.startfile(log_path)
    
    def _on_open_settings(self, icon, item):
        """打开设置界面"""
        def on_settings_save():
            """设置保存后的回调"""
            # 刷新菜单以反映可能的配置更改（如语言）
            set_language(app_state.config.get("language", "zh"))
            icon.menu = self.build_menu()
            
            # 如果热键更改，可能需要重启热键监听
            if self.restart_hotkey_callback:
                self.restart_hotkey_callback()

        def show_dialog_on_main():
            """在主线程显示设置对话框"""
            paused = False
            try:
                if self.settings_dialog and self.settings_dialog.is_alive():
                    self.settings_dialog.restore_and_focus()
                    return

                # 暂停热键监听
                if self.pause_hotkey_callback:
                    self.pause_hotkey_callback()
                    paused = True
                    
                dialog = SettingsDialog(
                    on_save=on_settings_save,
                    on_close=self.resume_hotkey_callback
                )
                self.settings_dialog = dialog

                def _clear_settings_dialog(event=None):
                    if getattr(event, "widget", None) is dialog.root or event is None:
                        self.settings_dialog = None

                dialog.root.bind("<Destroy>", _clear_settings_dialog)
                dialog.show()
            except Exception as e:
                log(f"Failed to show settings dialog: {e}")
                self.notification_manager.notify("PasteMD", f"Error opening settings: {e}", ok=False)
            finally:
                # 恢复热键监听
                if paused and self.resume_hotkey_callback:
                    self.resume_hotkey_callback()

        ui_queue = getattr(app_state, "ui_queue", None)
        if ui_queue is not None:
            ui_queue.put(show_dialog_on_main)
        else:
            show_dialog_on_main()

    def _build_html_formatting_menu(self) -> pystray.MenuItem:
        """构建 HTML 格式化子菜单"""
        return pystray.MenuItem(
            t("tray.menu.html_formatting"),
            pystray.Menu(
                pystray.MenuItem(
                    t("tray.menu.strikethrough_to_del"),
                    self._on_toggle_html_strikethrough,
                    checked=lambda item: self._get_html_formatting_option("strikethrough_to_del", True),
                ),
            ),
        )

    def _get_html_formatting_option(self, key: str, default: bool) -> bool:
        options = app_state.config.get("html_formatting", {})
        if isinstance(options, dict):
            return bool(options.get(key, default))
        return default

    def _on_toggle_html_strikethrough(self, icon, item):
        """切换删除线转 <del> 的 HTML 格式化配置"""
        current = self._get_html_formatting_option("strikethrough_to_del", True)
        if not isinstance(app_state.config.get("html_formatting"), dict):
            app_state.config["html_formatting"] = {}
        app_state.config["html_formatting"]["strikethrough_to_del"] = not current
        self._save_config()
        icon.menu = self.build_menu()

        status = (
            t("tray.status.html_strike_on")
            if app_state.config["html_formatting"].get("strikethrough_to_del", True)
            else t("tray.status.html_strike_off")
        )
        self.notification_manager.notify("PasteMD", status, ok=True)
    
    def _on_check_update(self, icon, item):
        """检查更新"""
        # 在后台线程中检查更新，避免阻塞 UI
        def check_in_background():
            try:
                # 导入版本号
                from ... import __version__
                
                checker = VersionChecker(__version__)
                result = checker.check_update()
                
                if result is None:
                    # 网络错误或检查失败
                    log("Version check failed - network error")
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_failure')}",
                        t("tray.update.network_error"),
                        ok=False
                    )
                elif result.get("has_update"):
                    latest_version = result.get("latest_version")
                    release_url = result.get("release_url")
                    
                    # 使用 update_version_info 方法更新版本信息并重新绘制菜单
                    self.update_version_info(icon, latest_version, release_url)
                    
                    # 通知用户有新版本，并自动打开下载页面
                    message = t("tray.update.opening_release", version=latest_version)
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_new_version')}",
                        message,
                        ok=True
                    )
                    
                    # 自动打开下载页面
                    try:
                        webbrowser.open(release_url)
                    except Exception as e:
                        log(f"Failed to open browser: {e}")
                    
                    log(f"New version available: {latest_version}")
                    log(f"Download URL: {release_url}")
                else:
                    # 无需更新，通知用户已是最新版本
                    current_version = result.get("current_version")
                    log(f"Already on latest version: {current_version}")
                    self.notification_manager.notify(
                        f"PasteMD - {t('tray.update.title_latest')}",
                        t("tray.update.latest_version", version=current_version),
                        ok=True
                    )
            except Exception as e:
                error_text = str(e)
                short_error = error_text if len(error_text) <= 15 else error_text[:12] + "..."
                self.notification_manager.notify(
                    f"PasteMD - {t('tray.update.title_unexpected_error')}",
                    t("tray.update.error_with_message", error=short_error),
                    ok=False
                )
                log(f"Error checking update: {e}")
        
        # 启动后台线程
        thread = threading.Thread(target=check_in_background, daemon=True)
        thread.start()
    
    def _on_open_release_page(self, icon, item):
        """打开发布页面"""
        if self.latest_release_url:
            try:
                webbrowser.open(self.latest_release_url)
                log(f"Opening release page: {self.latest_release_url}")
            except Exception as e:
                log(f"Failed to open browser: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    t("tray.error.open_release_page"),
                    ok=False
                )

    def update_version_info(self, icon, latest_version: str, release_url: str):
        """更新最新版本信息"""
        self.latest_version = latest_version
        self.latest_release_url = release_url
        icon.menu = self.build_menu()
    
    def _on_open_about_page(self, icon, item):
        """打开关于页面"""
        about_url = "http://pastemd.richqaq.cn"
        try:
            webbrowser.open(about_url)
            log(f"Opening about page: {about_url}")
        except Exception as e:
            log(f"Failed to open browser: {e}")
            self.notification_manager.notify(
                "PasteMD",
                t("tray.error.open_about_page", url=about_url),
                ok=False
            )
    
    def _on_quit(self, icon, item):
        """退出应用程序"""
        icon.stop()
        
        # 设置退出事件（AppState 中已经声明了这个属性）
        if app_state.quit_event is None:
            import threading
            app_state.quit_event = threading.Event()
        
        app_state.quit_event.set()
        
        # 发送退出信号到主程序
        if getattr(app_state, "ui_queue", None):
            try:
                app_state.ui_queue.put(None)
            except Exception as e:
                log(f"Failed to send quit signal: {e}")
    
    def _save_config(self):
        """保存配置"""
        try:
            self.config_loader.save(app_state.config)
        except Exception as e:
            log(f"Failed to save config: {e}")
