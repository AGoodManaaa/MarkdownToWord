# -*- coding: utf-8 -*-
"""
版本历史功能（Git集成）
支持本地版本管理、历史查看、版本对比和恢复
"""

import os
import subprocess
import customtkinter as ctk
from tkinter import messagebox, END
from typing import List, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime
from ui.dialog_utils import set_dialog_icon


@dataclass
class CommitInfo:
    """提交信息"""
    hash: str
    short_hash: str
    message: str
    author: str
    date: str
    relative_date: str


class VersionControlFeature:
    """版本历史功能（Git集成）"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.current_file: Optional[str] = None
        self.repo_root: Optional[str] = None
    
    def _has_git(self) -> bool:
        """检查是否安装了 Git"""
        try:
            subprocess.run(['git', '--version'], capture_output=True, check=True)
            return True
        except Exception:
            return False
    
    def _run_git(self, args: List[str], cwd: str = None) -> Tuple[bool, str]:
        """运行 Git 命令"""
        try:
            result = subprocess.run(
                ['git'] + args,
                cwd=cwd or self.repo_root,
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            return result.returncode == 0, result.stdout + result.stderr
        except Exception as e:
            return False, str(e)
    
    def _get_repo_root(self, path: str) -> Optional[str]:
        """获取 Git 仓库根目录"""
        try:
            result = subprocess.run(
                ['git', 'rev-parse', '--show-toplevel'],
                cwd=os.path.dirname(path),
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        return None
    
    def show_dialog(self):
        """显示版本历史对话框"""
        if not self._has_git():
            messagebox.showerror("错误", "未检测到 Git，请先安装 Git")
            return
        
        # 获取当前文件
        if hasattr(self.app, 'current_file') and self.app.current_file:
            self.current_file = self.app.current_file
        else:
            messagebox.showwarning("警告", "请先保存文件")
            return
        
        # 检查是否在 Git 仓库中
        self.repo_root = self._get_repo_root(self.current_file)
        
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🔄 版本历史")
        self.dialog.geometry("700x550")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 700) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 550) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 状态栏
        status_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray20"))
        status_frame.pack(fill="x", pady=(0, 10))
        
        file_name = os.path.basename(self.current_file)
        if self.repo_root:
            status_text = f"📁 {file_name} | ✓ Git 仓库已初始化"
            status_color = ("green", "lightgreen")
        else:
            status_text = f"📁 {file_name} | ⚠️ 未初始化 Git 仓库"
            status_color = ("orange", "yellow")
        
        self.status_label = ctk.CTkLabel(
            status_frame, text=status_text, 
            text_color=status_color,
            anchor="w"
        )
        self.status_label.pack(fill="x", padx=10, pady=8)
        
        # 工具栏
        toolbar = ctk.CTkFrame(main_frame, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 10))
        
        if not self.repo_root:
            ctk.CTkButton(
                toolbar, text="🚀 初始化 Git", width=120,
                fg_color=("green", "darkgreen"),
                command=self._init_repo
            ).pack(side="left", padx=(0, 5))
        else:
            ctk.CTkButton(
                toolbar, text="📝 保存版本", width=100,
                fg_color=("green", "darkgreen"),
                command=self._commit_version
            ).pack(side="left", padx=(0, 5))
            
            ctk.CTkButton(
                toolbar, text="🔄 刷新", width=80,
                command=self._refresh_history
            ).pack(side="left", padx=(0, 5))
        
        # 历史列表
        ctk.CTkLabel(main_frame, text="版本历史:", font=("", 13, "bold")).pack(anchor="w")
        
        self.history_container = ctk.CTkScrollableFrame(main_frame)
        self.history_container.pack(fill="both", expand=True, pady=5)
        
        if self.repo_root:
            self._refresh_history()
        else:
            ctk.CTkLabel(
                self.history_container,
                text="点击「初始化 Git」开始版本管理",
                text_color=("gray50", "gray70")
            ).pack(expand=True, pady=50)
    
    def _init_repo(self):
        """初始化 Git 仓库"""
        file_dir = os.path.dirname(self.current_file)
        
        # 初始化仓库
        success, output = self._run_git(['init'], cwd=file_dir)
        if not success:
            messagebox.showerror("错误", f"初始化失败: {output}")
            return
        
        self.repo_root = file_dir
        
        # 创建 .gitignore
        gitignore_path = os.path.join(file_dir, '.gitignore')
        if not os.path.exists(gitignore_path):
            with open(gitignore_path, 'w', encoding='utf-8') as f:
                f.write("# 忽略临时文件\n*.tmp\n*.bak\n~$*\n")
        
        # 首次提交
        self._run_git(['add', '.'], cwd=file_dir)
        self._run_git(['commit', '-m', '初始化版本库'], cwd=file_dir)
        
        messagebox.showinfo("成功", "Git 仓库已初始化，并创建了初始版本")
        
        # 刷新对话框
        self.dialog.destroy()
        self.dialog = None
        self.show_dialog()
    
    def _commit_version(self):
        """提交新版本"""
        if not self.repo_root:
            return
        
        # 弹出输入框
        commit_dialog = ctk.CTkInputDialog(
            text="请输入版本说明:",
            title="保存版本"
        )
        message = commit_dialog.get_input()
        
        if not message:
            return
        
        # 添加文件并提交
        rel_path = os.path.relpath(self.current_file, self.repo_root)
        self._run_git(['add', rel_path])
        
        success, output = self._run_git(['commit', '-m', message])
        
        if success:
            messagebox.showinfo("成功", "版本已保存")
            self._refresh_history()
        elif "nothing to commit" in output:
            messagebox.showinfo("提示", "文件没有修改，无需保存")
        else:
            messagebox.showerror("错误", f"保存失败: {output}")
    
    def _refresh_history(self):
        """刷新历史列表"""
        # 清空
        for widget in self.history_container.winfo_children():
            widget.destroy()
        
        if not self.repo_root:
            return
        
        # 获取当前文件的提交历史
        rel_path = os.path.relpath(self.current_file, self.repo_root)
        success, output = self._run_git([
            'log', '--follow', '--pretty=format:%H|%h|%s|%an|%ai|%ar',
            '-n', '50', '--', rel_path
        ])
        
        if not success or not output.strip():
            ctk.CTkLabel(
                self.history_container,
                text="暂无版本历史",
                text_color=("gray50", "gray70")
            ).pack(expand=True, pady=50)
            return
        
        commits = []
        for line in output.strip().split('\n'):
            if '|' in line:
                parts = line.split('|')
                if len(parts) >= 6:
                    commits.append(CommitInfo(
                        hash=parts[0],
                        short_hash=parts[1],
                        message=parts[2],
                        author=parts[3],
                        date=parts[4],
                        relative_date=parts[5]
                    ))
        
        for i, commit in enumerate(commits):
            self._create_commit_item(commit, is_current=(i == 0))
    
    def _create_commit_item(self, commit: CommitInfo, is_current: bool = False):
        """创建提交项"""
        bg_color = ("gray85", "gray25") if is_current else ("gray95", "gray18")
        frame = ctk.CTkFrame(self.history_container, fg_color=bg_color)
        frame.pack(fill="x", pady=2)
        
        # 信息区
        info_frame = ctk.CTkFrame(frame, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        
        # 标题行
        title_text = f"{'🔵 ' if is_current else ''}{commit.message}"
        ctk.CTkLabel(
            info_frame, text=title_text,
            font=("", 12, "bold" if is_current else "normal"),
            anchor="w"
        ).pack(fill="x")
        
        # 详情行
        detail_text = f"📅 {commit.relative_date} | 🔖 {commit.short_hash}"
        ctk.CTkLabel(
            info_frame, text=detail_text,
            font=("", 10),
            text_color=("gray50", "gray70"),
            anchor="w"
        ).pack(fill="x")
        
        # 操作按钮
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(side="right", padx=5)
        
        ctk.CTkButton(
            btn_frame, text="👁", width=35,
            command=lambda c=commit: self._preview_version(c)
        ).pack(side="left", padx=2)
        
        if not is_current:
            ctk.CTkButton(
                btn_frame, text="⏪", width=35,
                fg_color=("orange", "darkorange"),
                command=lambda c=commit: self._restore_version(c)
            ).pack(side="left", padx=2)
    
    def _preview_version(self, commit: CommitInfo):
        """预览历史版本"""
        rel_path = os.path.relpath(self.current_file, self.repo_root)
        success, content = self._run_git(['show', f'{commit.hash}:{rel_path}'])
        
        if not success:
            messagebox.showerror("错误", "无法获取该版本的内容")
            return
        
        # 显示预览窗口
        preview = ctk.CTkToplevel(self.dialog)
        preview.title(f"版本预览 - {commit.short_hash}: {commit.message}")
        preview.geometry("600x500")
        preview.transient(self.dialog)
        
        frame = ctk.CTkFrame(preview)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        text = ctk.CTkTextbox(frame, wrap="word")
        text.pack(fill="both", expand=True)
        text.insert("1.0", content)
        text.configure(state="disabled")
        
        ctk.CTkButton(
            frame, text="关闭",
            command=preview.destroy
        ).pack(pady=10)
    
    def _restore_version(self, commit: CommitInfo):
        """恢复到历史版本"""
        if not messagebox.askyesno(
            "确认恢复", 
            f"确定要恢复到版本 [{commit.short_hash}] 吗？\n\n"
            f"说明: {commit.message}\n"
            f"时间: {commit.relative_date}\n\n"
            "⚠️ 当前修改将被覆盖！"
        ):
            return
        
        rel_path = os.path.relpath(self.current_file, self.repo_root)
        success, content = self._run_git(['show', f'{commit.hash}:{rel_path}'])
        
        if not success:
            messagebox.showerror("错误", "无法获取该版本的内容")
            return
        
        # 写入文件
        with open(self.current_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # 更新编辑器
        textbox = self._get_textbox()
        if textbox:
            textbox.delete("1.0", END)
            textbox.insert("1.0", content)
        
        messagebox.showinfo("成功", f"已恢复到版本 [{commit.short_hash}]")
        self.dialog.destroy()
        self.dialog = None
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
