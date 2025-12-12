# -*- coding: utf-8 -*-

from ui.theme import save_config


class WindowGeometryFeature:
    def __init__(self, app):
        self.app = app

    def save(self):
        """保存窗口位置和大小"""
        try:
            self.app.config['window_geometry'] = self.app.geometry()
            save_config(self.app.config)
        except Exception:
            pass

    def restore(self):
        """恢复窗口位置和大小"""
        geometry = self.app.config.get('window_geometry')
        if geometry:
            try:
                self.app.geometry(geometry)
            except Exception:
                pass
