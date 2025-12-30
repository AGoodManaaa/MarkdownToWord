# -*- coding: utf-8 -*-
"""
链接检查器功能
检测失效链接、验证图片路径、生成检查报告
"""

import re
import os
import threading
import urllib.request
import urllib.error
import customtkinter as ctk
from tkinter import END
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass
from enum import Enum
from ui.dialog_utils import set_dialog_icon


class LinkStatus(Enum):
    """链接状态"""
    PENDING = "pending"
    OK = "ok"
    BROKEN = "broken"
    TIMEOUT = "timeout"
    ERROR = "error"


@dataclass
class LinkInfo:
    """链接信息"""
    url: str
    text: str
    line_number: int
    link_type: str  # 'url', 'image', 'anchor'
    status: LinkStatus = LinkStatus.PENDING
    error_message: str = ""


class LinkCheckerFeature:
    """链接检查器功能"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.links: List[LinkInfo] = []
        self.checking = False
        self.check_thread = None
        self.timeout = 5  # 秒
    
    def show_dialog(self):
        """显示链接检查对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🔗 链接检查器")
        self.dialog.geometry("650x500")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 650) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 500) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 顶部按钮区
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 10))
        
        self.scan_btn = ctk.CTkButton(
            btn_frame,
            text="🔍 扫描链接",
            width=120,
            command=self.scan_links
        )
        self.scan_btn.pack(side="left", padx=(0, 10))
        
        self.check_btn = ctk.CTkButton(
            btn_frame,
            text="✔️ 检查链接",
            width=120,
            command=self.check_links,
            state="disabled"
        )
        self.check_btn.pack(side="left", padx=(0, 10))
        
        self.stop_btn = ctk.CTkButton(
            btn_frame,
            text="⏹ 停止",
            width=80,
            command=self.stop_check,
            state="disabled",
            fg_color=("red", "darkred")
        )
        self.stop_btn.pack(side="left")
        
        # 状态标签
        self.status_label = ctk.CTkLabel(
            btn_frame,
            text="",
            anchor="e"
        )
        self.status_label.pack(side="right")
        
        # 统计信息
        stats_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"))
        stats_frame.pack(fill="x", pady=(0, 10))
        
        self.stats_labels = {}
        stats = [("total", "总链接"), ("ok", "✓ 有效"), ("broken", "✗ 失效"), ("pending", "⏳ 待检查")]
        for i, (key, label) in enumerate(stats):
            frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
            frame.pack(side="left", expand=True, padx=10, pady=8)
            ctk.CTkLabel(frame, text=label, font=("", 11)).pack()
            self.stats_labels[key] = ctk.CTkLabel(frame, text="0", font=("", 16, "bold"))
            self.stats_labels[key].pack()
        
        # 结果列表
        self.results_text = ctk.CTkTextbox(main_frame, wrap="none")
        self.results_text.pack(fill="both", expand=True)
        self.results_text.configure(state="disabled")
        
        # 底部说明
        ctk.CTkLabel(
            main_frame,
            text="💡 点击扫描链接以提取文档中的所有链接，然后点击检查链接验证其有效性",
            font=("", 11),
            text_color=("gray50", "gray70")
        ).pack(fill="x", pady=(10, 0))
        
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _get_document_text(self) -> Tuple[str, str]:
        """获取当前文档文本和基础路径"""
        text = ""
        base_path = ""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                textbox = getattr(self.app.input_text, '_textbox', self.app.input_text)
                text = textbox.get("1.0", "end-1c")
            if hasattr(self.app, 'current_file') and self.app.current_file:
                base_path = os.path.dirname(self.app.current_file)
        except Exception:
            pass
        return text, base_path
    
    def scan_links(self):
        """扫描文档中的所有链接"""
        text, base_path = self._get_document_text()
        if not text:
            self.status_label.configure(text="文档为空")
            return
        
        self.links = []
        lines = text.split('\n')
        
        # 扫描链接
        for i, line in enumerate(lines):
            line_num = i + 1
            
            # 图片链接 ![alt](url)
            for match in re.finditer(r'!\[([^\]]*)\]\(([^)]+)\)', line):
                alt = match.group(1)
                url = match.group(2).strip()
                self.links.append(LinkInfo(
                    url=url,
                    text=alt or url,
                    line_number=line_num,
                    link_type='image'
                ))
            
            # 普通链接 [text](url)
            for match in re.finditer(r'(?<!!)\[([^\]]+)\]\(([^)]+)\)', line):
                text_content = match.group(1)
                url = match.group(2).strip()
                # 判断是锚点还是URL
                link_type = 'anchor' if url.startswith('#') else 'url'
                self.links.append(LinkInfo(
                    url=url,
                    text=text_content,
                    line_number=line_num,
                    link_type=link_type
                ))
            
            # 裸链接 <http://...> 或自动链接
            for match in re.finditer(r'<(https?://[^>]+)>', line):
                url = match.group(1)
                self.links.append(LinkInfo(
                    url=url,
                    text=url,
                    line_number=line_num,
                    link_type='url'
                ))
        
        self._update_display()
        self._update_stats()
        
        if self.links:
            self.check_btn.configure(state="normal")
            self.status_label.configure(text=f"扫描完成，找到 {len(self.links)} 个链接")
        else:
            self.check_btn.configure(state="disabled")
            self.status_label.configure(text="未找到链接")
    
    def check_links(self):
        """开始检查链接"""
        if not self.links or self.checking:
            return
        
        self.checking = True
        self.scan_btn.configure(state="disabled")
        self.check_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        
        # 在线程中检查
        self.check_thread = threading.Thread(target=self._check_links_thread, daemon=True)
        self.check_thread.start()
    
    def _check_links_thread(self):
        """检查链接的线程"""
        _, base_path = self._get_document_text()
        
        for i, link in enumerate(self.links):
            if not self.checking:
                break
            
            self._update_status(f"检查中 ({i+1}/{len(self.links)}): {link.url[:50]}...")
            
            try:
                if link.link_type == 'anchor':
                    # 锚点链接暂时标记为OK
                    link.status = LinkStatus.OK
                elif link.link_type == 'image':
                    # 图片链接
                    self._check_image(link, base_path)
                else:
                    # URL链接
                    self._check_url(link)
            except Exception as e:
                link.status = LinkStatus.ERROR
                link.error_message = str(e)
            
            # 更新UI
            self._schedule_update()
        
        self.checking = False
        self._schedule_finish()
    
    def _check_url(self, link: LinkInfo):
        """检查URL链接"""
        url = link.url
        
        # 跳过特殊协议
        if any(url.startswith(p) for p in ['mailto:', 'tel:', 'javascript:', 'file:']):
            link.status = LinkStatus.OK
            return
        
        # 添加协议
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        
        try:
            req = urllib.request.Request(
                url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            )
            with urllib.request.urlopen(req, timeout=self.timeout) as response:
                if response.status == 200:
                    link.status = LinkStatus.OK
                else:
                    link.status = LinkStatus.BROKEN
                    link.error_message = f"HTTP {response.status}"
        except urllib.error.HTTPError as e:
            link.status = LinkStatus.BROKEN
            link.error_message = f"HTTP {e.code}"
        except urllib.error.URLError as e:
            link.status = LinkStatus.BROKEN
            link.error_message = str(e.reason)[:50]
        except TimeoutError:
            link.status = LinkStatus.TIMEOUT
            link.error_message = "请求超时"
        except Exception as e:
            link.status = LinkStatus.ERROR
            link.error_message = str(e)[:50]
    
    def _check_image(self, link: LinkInfo, base_path: str):
        """检查图片路径"""
        url = link.url
        
        # 网络图片
        if url.startswith(('http://', 'https://')):
            self._check_url(link)
            return
        
        # 本地图片
        if base_path:
            full_path = os.path.join(base_path, url)
            if os.path.isfile(full_path):
                link.status = LinkStatus.OK
            else:
                link.status = LinkStatus.BROKEN
                link.error_message = "文件不存在"
        else:
            # 无基础路径，检查绝对路径
            if os.path.isfile(url):
                link.status = LinkStatus.OK
            else:
                link.status = LinkStatus.BROKEN
                link.error_message = "文件不存在"
    
    def _update_status(self, text: str):
        """更新状态（线程安全）"""
        try:
            self.dialog.after(0, lambda: self.status_label.configure(text=text))
        except Exception:
            pass
    
    def _schedule_update(self):
        """调度UI更新"""
        try:
            self.dialog.after(0, self._update_display)
            self.dialog.after(0, self._update_stats)
        except Exception:
            pass
    
    def _schedule_finish(self):
        """调度完成处理"""
        try:
            self.dialog.after(0, self._on_check_finished)
        except Exception:
            pass
    
    def _on_check_finished(self):
        """检查完成"""
        self.scan_btn.configure(state="normal")
        self.check_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        
        broken = sum(1 for l in self.links if l.status in [LinkStatus.BROKEN, LinkStatus.TIMEOUT, LinkStatus.ERROR])
        if broken > 0:
            self.status_label.configure(
                text=f"检查完成，发现 {broken} 个问题",
                text_color=("red", "lightcoral")
            )
        else:
            self.status_label.configure(
                text="检查完成，所有链接有效 ✓",
                text_color=("green", "lightgreen")
            )
    
    def stop_check(self):
        """停止检查"""
        self.checking = False
        self.status_label.configure(text="已停止")
        self.scan_btn.configure(state="normal")
        self.check_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
    
    def _update_display(self):
        """更新结果显示"""
        try:
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", END)
            
            for link in self.links:
                # 状态图标
                if link.status == LinkStatus.OK:
                    icon = "✓"
                    color = "green"
                elif link.status == LinkStatus.BROKEN:
                    icon = "✗"
                    color = "red"
                elif link.status == LinkStatus.TIMEOUT:
                    icon = "⏱"
                    color = "orange"
                elif link.status == LinkStatus.ERROR:
                    icon = "⚠"
                    color = "red"
                else:
                    icon = "○"
                    color = "gray"
                
                # 类型图标
                type_icon = {"image": "🖼", "url": "🔗", "anchor": "⚓"}.get(link.link_type, "🔗")
                
                # 显示链接信息
                line = f"{icon} L{link.line_number:4d} {type_icon} {link.url[:60]}"
                if len(link.url) > 60:
                    line += "..."
                if link.error_message:
                    line += f"  [{link.error_message}]"
                line += "\n"
                
                self.results_text.insert(END, line)
            
            self.results_text.configure(state="disabled")
        except Exception:
            pass
    
    def _update_stats(self):
        """更新统计信息"""
        try:
            total = len(self.links)
            ok = sum(1 for l in self.links if l.status == LinkStatus.OK)
            broken = sum(1 for l in self.links if l.status in [LinkStatus.BROKEN, LinkStatus.TIMEOUT, LinkStatus.ERROR])
            pending = sum(1 for l in self.links if l.status == LinkStatus.PENDING)
            
            self.stats_labels["total"].configure(text=str(total))
            self.stats_labels["ok"].configure(text=str(ok))
            self.stats_labels["broken"].configure(text=str(broken))
            self.stats_labels["pending"].configure(text=str(pending))
        except Exception:
            pass
    
    def _on_close(self):
        """关闭对话框"""
        self.checking = False
        self.dialog.destroy()
        self.dialog = None
