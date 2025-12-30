import os
import shutil
import customtkinter as ctk
from tkinter import filedialog, messagebox

class TemplateSelectorFeature:
    """Word模板管理功能"""
    
    def __init__(self, app):
        self.app = app
        # templates 目录在项目根目录下
        self.project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.templates_dir = os.path.join(self.project_root, 'templates')
        
        if not os.path.exists(self.templates_dir):
            try:
                os.makedirs(self.templates_dir)
            except Exception:
                pass

    def get_templates(self):
        """获取可用模板列表"""
        templates = ["默认 (空白)"]
        if os.path.exists(self.templates_dir):
            try:
                for f in os.listdir(self.templates_dir):
                    if f.lower().endswith('.docx'):
                        templates.append(f)
            except Exception:
                pass
        return templates

    def get_template_path(self, template_name):
        """获取模板的完整路径"""
        if not template_name or template_name == "默认 (空白)":
            return None
        
        path = os.path.join(self.templates_dir, template_name)
        if os.path.exists(path):
            return path
        return None

    def import_template(self):
        """导入新模板"""
        file_path = filedialog.askopenfilename(
            title="导入Word模板",
            filetypes=[("Word模板", "*.docx")]
        )
        if file_path:
            try:
                dest = os.path.join(self.templates_dir, os.path.basename(file_path))
                shutil.copy2(file_path, dest)
                messagebox.showinfo("成功", f"模板已导入: {os.path.basename(file_path)}")
                return True
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
        return False
