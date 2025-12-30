# -*- coding: utf-8 -*-
"""
文档加密和水印功能
"""

import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
from docx import Document


class DocumentSecurityFeature:
    """文档安全功能"""
    
    def __init__(self, app):
        self.app = app
        self.security_dialog = None
        
    def show_security_dialog(self):
        """显示安全设置对话框"""
        if self.security_dialog and self.security_dialog.winfo_exists():
            self.security_dialog.focus()
            return
            
        self.security_dialog = ctk.CTkToplevel(self.app)
        self.security_dialog.title("🔐 文档安全")
        self.security_dialog.geometry("600x650")
        self.security_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.security_dialog,
            text="文档安全设置",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 选项卡
        tabview = ctk.CTkTabview(self.security_dialog, width=560, height=480)
        tabview.pack(padx=20, pady=10)
        
        tabview.add("文档加密")
        tabview.add("添加水印")
        tabview.add("元数据清理")
        
        # === 文档加密标签页 ===
        self._create_encryption_tab(tabview.tab("文档加密"))
        
        # === 添加水印标签页 ===
        self._create_watermark_tab(tabview.tab("添加水印"))
        
        # === 元数据清理标签页 ===
        self._create_metadata_tab(tabview.tab("元数据清理"))
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            self.security_dialog,
            text="关闭",
            command=self.security_dialog.destroy,
            width=120
        )
        close_btn.pack(pady=15)
        
    def _create_encryption_tab(self, parent):
        """创建文档加密标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="为导出的Word文档添加密码保护",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=15)
        
        # 密码输入
        password_frame = ctk.CTkFrame(parent, fg_color="transparent")
        password_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(password_frame, text="设置密码:").pack(anchor="w", pady=5)
        self.password_entry = ctk.CTkEntry(
            password_frame,
            placeholder_text="请输入密码",
            show="*",
            width=400
        )
        self.password_entry.pack(anchor="w", pady=5)
        
        ctk.CTkLabel(password_frame, text="确认密码:").pack(anchor="w", pady=5)
        self.password_confirm_entry = ctk.CTkEntry(
            password_frame,
            placeholder_text="请再次输入密码",
            show="*",
            width=400
        )
        self.password_confirm_entry.pack(anchor="w", pady=5)
        
        # 加密选项
        options_frame = ctk.CTkFrame(parent, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=10)
        
        self.protect_structure_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame,
            text="保护文档结构（禁止编辑）",
            variable=self.protect_structure_var
        ).pack(anchor="w", pady=5)
        
        self.allow_read_only_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame,
            text="允许只读访问",
            variable=self.allow_read_only_var
        ).pack(anchor="w", pady=5)
        
        # 说明
        info_text = ctk.CTkTextbox(parent, height=150, width=520)
        info_text.pack(padx=20, pady=15)
        info_text.insert("1.0",
            "💡 提示:\n\n"
            "• 密码长度建议至少8位，包含字母、数字和符号\n"
            "• 请务必记住密码，忘记密码将无法打开文档\n"
            "• 加密后的文档在导出时生效\n"
            "• 注意：此功能需要 python-docx 支持，部分版本可能不支持加密\n\n"
            "⚠️ 警告:\n"
            "• 加密功能为基础安全保护，不适用于高度敏感文档\n"
            "• 建议同时使用Windows或Office自带的加密功能"
        )
        info_text.configure(state="disabled")
        
        # 应用按钮
        apply_btn = ctk.CTkButton(
            parent,
            text="✅ 应用加密设置",
            command=self._apply_encryption,
            fg_color="#10B981",
            hover_color="#059669",
            width=160
        )
        apply_btn.pack(pady=10)
        
    def _create_watermark_tab(self, parent):
        """创建水印标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="为导出的文档添加文字水印",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=15)
        
        # 水印文本
        text_frame = ctk.CTkFrame(parent, fg_color="transparent")
        text_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(text_frame, text="水印文字:").pack(anchor="w", pady=5)
        self.watermark_text_entry = ctk.CTkEntry(
            text_frame,
            placeholder_text="例如：机密文档、仅供内部使用",
            width=400
        )
        self.watermark_text_entry.pack(anchor="w", pady=5)
        
        # 水印样式
        style_frame = ctk.CTkFrame(parent, fg_color="transparent")
        style_frame.pack(fill="x", padx=20, pady=10)
        
        # 颜色选择
        color_row = ctk.CTkFrame(style_frame, fg_color="transparent")
        color_row.pack(fill="x", pady=5)
        
        ctk.CTkLabel(color_row, text="颜色:").pack(side="left", padx=5)
        self.watermark_color_var = ctk.StringVar(value="浅灰色")
        ctk.CTkOptionMenu(
            color_row,
            values=["浅灰色", "红色", "蓝色", "绿色", "自定义"],
            variable=self.watermark_color_var,
            width=120
        ).pack(side="left", padx=5)
        
        # 透明度
        opacity_row = ctk.CTkFrame(style_frame, fg_color="transparent")
        opacity_row.pack(fill="x", pady=5)
        
        ctk.CTkLabel(opacity_row, text="透明度:").pack(side="left", padx=5)
        self.watermark_opacity = ctk.CTkSlider(
            opacity_row,
            from_=0,
            to=100,
            number_of_steps=100,
            width=300
        )
        self.watermark_opacity.set(30)
        self.watermark_opacity.pack(side="left", padx=5)
        
        self.opacity_label = ctk.CTkLabel(opacity_row, text="30%", width=50)
        self.opacity_label.pack(side="left", padx=5)
        self.watermark_opacity.configure(
            command=lambda v: self.opacity_label.configure(text=f"{int(v)}%")
        )
        
        # 字体大小
        size_row = ctk.CTkFrame(style_frame, fg_color="transparent")
        size_row.pack(fill="x", pady=5)
        
        ctk.CTkLabel(size_row, text="字体大小:").pack(side="left", padx=5)
        self.watermark_size_var = ctk.StringVar(value="48")
        ctk.CTkOptionMenu(
            size_row,
            values=["24", "36", "48", "60", "72"],
            variable=self.watermark_size_var,
            width=100
        ).pack(side="left", padx=5)
        
        # 位置选择
        position_frame = ctk.CTkFrame(parent, fg_color="transparent")
        position_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(position_frame, text="水印位置:").pack(anchor="w", pady=5)
        
        self.watermark_position_var = ctk.StringVar(value="diagonal")
        positions = ctk.CTkSegmentedButton(
            position_frame,
            values=["对角线", "中心", "平铺"],
            variable=self.watermark_position_var
        )
        positions.pack(anchor="w", pady=5)
        
        # 预览
        preview_frame = ctk.CTkFrame(parent)
        preview_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.watermark_preview = ctk.CTkLabel(
            preview_frame,
            text="水印预览区域",
            fg_color="#F3F4F6",
            height=100
        )
        self.watermark_preview.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 应用按钮
        apply_btn = ctk.CTkButton(
            parent,
            text="✅ 应用水印设置",
            command=self._apply_watermark,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=160
        )
        apply_btn.pack(pady=10)
        
    def _create_metadata_tab(self, parent):
        """创建元数据清理标签页"""
        info_label = ctk.CTkLabel(
            parent,
            text="清理文档中的敏感元数据信息",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=15)
        
        # 清理选项
        options_frame = ctk.CTkFrame(parent, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(
            options_frame,
            text="选择要清理的信息:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=10)
        
        self.clear_author_var = ctk.BooleanVar(value=True)
        self.clear_company_var = ctk.BooleanVar(value=True)
        self.clear_comments_var = ctk.BooleanVar(value=True)
        self.clear_revision_var = ctk.BooleanVar(value=True)
        self.clear_created_time_var = ctk.BooleanVar(value=False)
        
        ctk.CTkCheckBox(
            options_frame,
            text="作者信息",
            variable=self.clear_author_var
        ).pack(anchor="w", pady=3)
        
        ctk.CTkCheckBox(
            options_frame,
            text="公司/组织信息",
            variable=self.clear_company_var
        ).pack(anchor="w", pady=3)
        
        ctk.CTkCheckBox(
            options_frame,
            text="注释和批注",
            variable=self.clear_comments_var
        ).pack(anchor="w", pady=3)
        
        ctk.CTkCheckBox(
            options_frame,
            text="修订历史",
            variable=self.clear_revision_var
        ).pack(anchor="w", pady=3)
        
        ctk.CTkCheckBox(
            options_frame,
            text="创建时间/修改时间",
            variable=self.clear_created_time_var
        ).pack(anchor="w", pady=3)
        
        # 说明
        info_text = ctk.CTkTextbox(parent, height=200, width=520)
        info_text.pack(padx=20, pady=15)
        info_text.insert("1.0",
            "💡 为什么要清理元数据?\n\n"
            "文档元数据可能包含:\n"
            "• 作者姓名、公司名称\n"
            "• 文档创建和修改时间\n"
            "• 修订历史和批注\n"
            "• 计算机名称和用户信息\n\n"
            "这些信息可能会泄露:\n"
            "• 个人身份信息\n"
            "• 组织内部信息\n"
            "• 文档编辑历史\n\n"
            "⚠️ 注意:\n"
            "清理元数据后无法恢复，建议先备份原文档。"
        )
        info_text.configure(state="disabled")
        
        # 应用按钮
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        ctk.CTkButton(
            btn_frame,
            text="🔍 检查元数据",
            command=self._check_metadata,
            width=140
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="🧹 清理元数据",
            command=self._clear_metadata,
            fg_color="#EF4444",
            hover_color="#DC2626",
            width=140
        ).pack(side="left", padx=5)
        
    def _apply_encryption(self):
        """应用加密设置"""
        password = self.password_entry.get()
        confirm = self.password_confirm_entry.get()
        
        if not password:
            messagebox.showwarning("提示", "请输入密码！")
            return
            
        if password != confirm:
            messagebox.showerror("错误", "两次输入的密码不一致！")
            return
            
        if len(password) < 6:
            messagebox.showwarning("提示", "密码长度至少6位！")
            return
            
        # 保存密码到配置（注意：实际应用中应该安全存储）
        from ui.theme import load_config, save_config
        config = load_config()
        config['document_password'] = password  # 实际应该加密存储
        config['protect_structure'] = self.protect_structure_var.get()
        save_config(config)
        
        messagebox.showinfo("成功", "加密设置已保存，将在下次导出时生效")
        self.app.update_status("✅ 文档加密已设置")
        
    def _apply_watermark(self):
        """应用水印设置"""
        text = self.watermark_text_entry.get()
        
        if not text:
            messagebox.showwarning("提示", "请输入水印文字！")
            return
            
        # 保存水印设置到配置
        from ui.theme import load_config, save_config
        config = load_config()
        config['watermark'] = {
            'text': text,
            'color': self.watermark_color_var.get(),
            'opacity': int(self.watermark_opacity.get()),
            'size': int(self.watermark_size_var.get()),
            'position': self.watermark_position_var.get()
        }
        save_config(config)
        
        messagebox.showinfo("成功", "水印设置已保存，将在下次导出时生效")
        self.app.update_status("✅ 水印已设置")
        
    def _check_metadata(self):
        """检查文档元数据"""
        # 需要先有导出的文档
        messagebox.showinfo(
            "元数据检查",
            "此功能需要检查已导出的Word文档。\n\n"
            "请先导出文档，然后选择文件进行检查。"
        )
        
        file_path = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            doc = Document(file_path)
            core_props = doc.core_properties
            
            metadata_info = f"文档元数据:\n\n"
            metadata_info += f"作者: {core_props.author or '(无)'}\n"
            metadata_info += f"标题: {core_props.title or '(无)'}\n"
            metadata_info += f"主题: {core_props.subject or '(无)'}\n"
            metadata_info += f"创建时间: {core_props.created or '(无)'}\n"
            metadata_info += f"修改时间: {core_props.modified or '(无)'}\n"
            metadata_info += f"最后修改人: {core_props.last_modified_by or '(无)'}\n"
            metadata_info += f"版本: {core_props.revision or '(无)'}\n"
            
            messagebox.showinfo("元数据信息", metadata_info)
            
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文档元数据:\n{e}")
            
    def _clear_metadata(self):
        """清理文档元数据"""
        result = messagebox.askyesno(
            "确认",
            "确定要清理文档元数据吗？\n此操作不可恢复！"
        )
        
        if not result:
            return
            
        # 保存清理选项到配置
        from ui.theme import load_config, save_config
        config = load_config()
        config['clear_metadata'] = {
            'author': self.clear_author_var.get(),
            'company': self.clear_company_var.get(),
            'comments': self.clear_comments_var.get(),
            'revision': self.clear_revision_var.get(),
            'time': self.clear_created_time_var.get()
        }
        save_config(config)
        
        messagebox.showinfo("成功", "元数据清理设置已保存，将在下次导出时生效")
        self.app.update_status("✅ 元数据清理已设置")
