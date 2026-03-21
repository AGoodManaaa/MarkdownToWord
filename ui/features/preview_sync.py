# -*- coding: utf-8 -*-

import tkinter as tk
import time

from ui.features.precise_scroll_sync import PreciseScrollSync, IncrementalPreviewUpdater
from utils import convert_latex_delimiters


class PreviewSyncFeature:
    def __init__(self, app):
        self.app = app

        self._last_preview_ts = 0.0
        self._last_outline_ts = 0.0
        self._last_counts_ts = 0.0
        self._throttle_ms_preview = 120
        self._throttle_ms_outline = 180
        self._throttle_ms_counts = 220
        self._pending_preview_id = None

        self._precise_sync = None
        self._incremental_updater = None
        self._last_content_for_sync = ""

        self._last_line_map_ts = 0.0
        self._throttle_ms_line_map = 500
        self._last_preview_content = None

        self._scroll_sync_enabled = True

        self._last_editor_scroll_ts = 0.0
        self._last_preview_scroll_ts = 0.0
        self._last_editor_scroll_pos = None
        self._last_preview_scroll_pos = None
        self._scroll_throttle_ms = 80
        self._scroll_delta_threshold = 0.004

        self._performance_mode = "normal"
        self._last_performance_hint = None

    def _detect_performance_mode(self, content: str) -> str:
        config = getattr(self.app, "config", {}) or {}
        requested = str(config.get("performance_mode", "auto")).lower()
        if requested in {"normal", "high"}:
            return requested

        line_threshold = int(config.get("large_doc_line_threshold", 800) or 800)
        char_threshold = int(config.get("large_doc_char_threshold", 30000) or 30000)
        image_threshold = int(config.get("large_doc_image_threshold", 20) or 20)

        line_count = content.count("\n") + 1 if content else 0
        char_count = len(content or "")
        image_count = (content or "").count("![")

        if line_count >= line_threshold or char_count >= char_threshold or image_count >= image_threshold:
            return "high"
        return "normal"

    def _apply_performance_mode(self, mode: str) -> None:
        mode = "high" if str(mode).lower() == "high" else "normal"
        if mode == self._performance_mode:
            return

        self._performance_mode = mode
        if mode == "high":
            self._throttle_ms_preview = 220
            self._throttle_ms_outline = 320
            self._throttle_ms_counts = 320
            self._throttle_ms_line_map = 900
            self._scroll_throttle_ms = 140
            self._scroll_delta_threshold = 0.015
            hint = "已启用性能模式"
        else:
            self._throttle_ms_preview = 120
            self._throttle_ms_outline = 180
            self._throttle_ms_counts = 220
            self._throttle_ms_line_map = 500
            self._scroll_throttle_ms = 80
            self._scroll_delta_threshold = 0.004
            hint = "已恢复标准模式"

        try:
            if hasattr(self.app, "preview") and hasattr(self.app.preview, "set_performance_mode"):
                self.app.preview.set_performance_mode(mode)
        except Exception:
            pass

        if hint != self._last_performance_hint:
            self._last_performance_hint = hint
            try:
                self.app.update_status(hint)
            except Exception:
                pass

    def set_scroll_sync_enabled(self, enabled: bool) -> None:
        try:
            self._scroll_sync_enabled = bool(enabled)
        except Exception:
            self._scroll_sync_enabled = True

        try:
            if hasattr(self.app, "preview") and hasattr(self.app.preview, "set_sync_scroll_enabled"):
                self.app.preview.set_sync_scroll_enabled(self._scroll_sync_enabled)
        except Exception:
            pass

    def toggle_scroll_sync(self) -> bool:
        self.set_scroll_sync_enabled(not bool(getattr(self, "_scroll_sync_enabled", True)))
        return bool(getattr(self, "_scroll_sync_enabled", True))

    def _init_precise_sync(self):
        if self._performance_mode == "high":
            return
        if self._precise_sync is not None:
            return

        try:
            if hasattr(self.app, "input_editor") and hasattr(self.app, "preview"):
                self._precise_sync = PreciseScrollSync(
                    self.app.input_editor,
                    self.app.preview,
                    self.app,
                )
                self._incremental_updater = IncrementalPreviewUpdater(self.app.preview)
        except Exception:
            pass

    def _update_line_map(self, content: str):
        if self._performance_mode == "high":
            return
        now = time.monotonic()
        min_interval = float(self._throttle_ms_line_map) / 1000.0

        if (now - self._last_line_map_ts) < min_interval:
            return
        if content == self._last_content_for_sync:
            return

        self._last_line_map_ts = now
        self._last_content_for_sync = content
        self._init_precise_sync()

        if self._precise_sync:
            try:
                self._precise_sync.build_line_map(content)
            except Exception:
                pass

    def on_text_change_debounced(self, event=None):
        try:
            if hasattr(self.app, "_debounce_id") and self.app._debounce_id:
                try:
                    self.app.after_cancel(self.app._debounce_id)
                except Exception:
                    pass

            delay = 120 if self._performance_mode == "high" else 50
            self.app._debounce_id = self.app.after(delay, lambda: self.on_text_change(event))
        except Exception as e:
            print(f"Debounce error: {e}")

    def on_text_change(self, event=None):
        try:
            content = self.app.input_text.get("1.0", "end-1c")
        except Exception:
            return

        self._apply_performance_mode(self._detect_performance_mode(content))
        preview_content = convert_latex_delimiters(content)
        self._update_line_map(content)

        now = time.monotonic()

        if hasattr(self.app, "preview"):
            try:
                min_interval = 0.12 if self._performance_mode == "high" else 0.05

                if preview_content == self._last_preview_content:
                    pass
                elif (now - self._last_preview_ts) >= min_interval:
                    self._last_preview_ts = now
                    if self._pending_preview_id:
                        self.app.after_cancel(self._pending_preview_id)
                        self._pending_preview_id = None

                    self._last_preview_content = preview_content
                    self.app.preview.set_updating(True)
                    self.app.preview.update_preview(preview_content)
                    self.app.preview.set_updating(False)
                else:
                    delay = int(max(1, (min_interval - (now - self._last_preview_ts)) * 1000))
                    if self._pending_preview_id:
                        self.app.after_cancel(self._pending_preview_id)
                    self._last_preview_content = preview_content
                    self._pending_preview_id = self.app.after(
                        delay,
                        lambda c=preview_content: self._render_preview_later(c),
                    )
            except Exception as e:
                print(f"Preview refresh error: {e}")
                if hasattr(self.app, "preview"):
                    self.app.preview.set_updating(False)

        try:
            if hasattr(self.app, "outline_view"):
                min_interval = float(self._throttle_ms_outline) / 1000.0
                if (now - self._last_outline_ts) >= min_interval:
                    self._last_outline_ts = now
                    self.app.outline_view.update_outline(content)
        except Exception:
            pass

        try:
            min_interval = float(self._throttle_ms_counts) / 1000.0
            if (now - self._last_counts_ts) >= min_interval:
                self._last_counts_ts = now
                if hasattr(self.app, "statistics_detail"):
                    self.app.statistics_detail.update_status_bar(content)
                else:
                    self.app.status_bar_feature.update_counts(content)
        except Exception:
            pass

        new_modified = content != getattr(self.app, "_last_saved_content", "")
        if new_modified != getattr(self.app, "_content_modified", False):
            self.app._content_modified = new_modified
            try:
                self.app._update_title()
            except Exception:
                pass
            try:
                if hasattr(self.app, "tab_manager") and new_modified:
                    self.app.tab_manager.mark_current_modified()
            except Exception:
                pass

    def on_preview_change(self, markdown_text: str):
        if hasattr(self.app, "_preview_updating") and self.app._preview_updating:
            return

        self.app._preview_updating = True
        try:
            try:
                cursor_pos = self.app.input_text.text.index(tk.INSERT)
            except Exception:
                cursor_pos = None

            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", markdown_text)

            if cursor_pos:
                try:
                    self.app.input_text.text.mark_set(tk.INSERT, cursor_pos)
                except Exception:
                    pass

            self.app.status_bar_feature.update_counts(markdown_text)

            try:
                if hasattr(self.app, "outline_view"):
                    self.app.outline_view.update_outline(markdown_text)
            except Exception:
                pass

            self.app._last_content_snapshot = markdown_text

            new_modified = markdown_text != getattr(self.app, "_last_saved_content", "")
            if new_modified != getattr(self.app, "_content_modified", False):
                self.app._content_modified = new_modified
                try:
                    self.app._update_title()
                except Exception:
                    pass

            try:
                self.app._update_cursor_position()
            except Exception:
                pass

            try:
                self.app.update_status("预览区已编辑")
            except Exception:
                pass
        finally:
            self.app._preview_updating = False

    def _render_preview_later(self, content: str) -> None:
        try:
            self._pending_preview_id = None
            if not hasattr(self.app, "preview"):
                return
            if not getattr(self.app, "preview_visible", True):
                return

            now = time.monotonic()
            self._last_preview_ts = now
            try:
                self.app.preview.set_updating(True)
                self.app.preview.update_preview(content)
            finally:
                try:
                    self.app.preview.set_updating(False)
                except Exception:
                    pass
        except Exception:
            pass

    def on_editor_scroll(self, position: float):
        try:
            if not bool(getattr(self, "_scroll_sync_enabled", True)):
                return
            if not hasattr(self.app, "preview") or not getattr(self.app, "preview_visible", True):
                return
            if getattr(self.app.preview, "_is_rendering", False):
                return

            now = time.monotonic()
            if (now - float(getattr(self, "_last_editor_scroll_ts", 0.0))) < (float(self._scroll_throttle_ms) / 1000.0):
                return
            last_pos = getattr(self, "_last_editor_scroll_pos", None)
            if isinstance(last_pos, (int, float)) and abs(float(position) - float(last_pos)) < float(self._scroll_delta_threshold):
                return
            self._last_editor_scroll_ts = now
            self._last_editor_scroll_pos = float(position)

            if self._performance_mode == "high":
                if hasattr(self.app.preview, "sync_scroll_to"):
                    self.app.preview.sync_scroll_to(position)
                else:
                    self.app.preview.text.yview_moveto(position)
                return

            self._init_precise_sync()

            if self._precise_sync:
                self._precise_sync.sync_editor_to_preview()
            elif hasattr(self.app.preview, "sync_scroll_to"):
                self.app.preview.sync_scroll_to(position)
            else:
                self.app.preview.text.yview_moveto(position)
        except Exception:
            pass

    def on_preview_scroll(self, position: float):
        try:
            if not bool(getattr(self, "_scroll_sync_enabled", True)):
                return
            if not hasattr(self.app, "input_editor"):
                return
            if hasattr(self.app, "preview") and getattr(self.app.preview, "_is_rendering", False):
                return

            now = time.monotonic()
            if (now - float(getattr(self, "_last_preview_scroll_ts", 0.0))) < (float(self._scroll_throttle_ms) / 1000.0):
                return
            last_pos = getattr(self, "_last_preview_scroll_pos", None)
            if isinstance(last_pos, (int, float)) and abs(float(position) - float(last_pos)) < float(self._scroll_delta_threshold):
                return
            self._last_preview_scroll_ts = now
            self._last_preview_scroll_pos = float(position)

            if self._performance_mode == "high":
                self.app.input_editor._textbox.yview_moveto(position)
                return

            self._init_precise_sync()

            if self._precise_sync:
                self._precise_sync.sync_preview_to_editor(position)
            elif hasattr(self.app.input_editor, "_textbox"):
                self.app.input_editor._textbox.yview_moveto(position)
        except Exception:
            pass

    def get_sync_accuracy(self) -> int:
        if not self._precise_sync:
            return -1

        try:
            editor_line = self._precise_sync._get_editor_first_visible_line()
            if editor_line is None:
                return -1

            if hasattr(self.app.preview, "text"):
                preview_pos = self.app.preview.text.yview()[0]
                preview_line = self._precise_sync._find_source_line_from_preview(preview_pos)
                if preview_line:
                    return self._precise_sync.get_sync_accuracy(editor_line, preview_line)

            return -1
        except Exception:
            return -1

    def is_sync_accurate(self, tolerance: int = 2) -> bool:
        accuracy = self.get_sync_accuracy()
        return accuracy >= 0 and accuracy <= tolerance
