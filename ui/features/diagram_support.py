# -*- coding: utf-8 -*-
"""
Mermaid/PlantUML 图表支持
"""

import os
import tempfile
import subprocess
from tkinter import messagebox
import customtkinter as ctk
import re


class DiagramRenderer:
    """图表渲染器基类"""
    
    def __init__(self):
        self.temp_dir = tempfile.gettempdir()
        
    def render(self, code: str, output_path: str) -> bool:
        """渲染图表为图片"""
        raise NotImplementedError


class MermaidRenderer(DiagramRenderer):
    """Mermaid图表渲染器"""
    
    def __init__(self):
        super().__init__()
        self.mermaid_cli = self._find_mermaid_cli()
        
    def _find_mermaid_cli(self) -> str:
        """查找mermaid-cli"""
        # 尝试常见的位置
        possible_paths = [
            "mmdc",  # 全局安装
            "node_modules/.bin/mmdc",  # 本地安装
            os.path.join(os.path.expanduser("~"), "node_modules", ".bin", "mmdc")
        ]
        
        for path in possible_paths:
            try:
                result = subprocess.run(
                    [path, "--version"],
                    capture_output=True,
                    timeout=5
                )
                if result.returncode == 0:
                    return path
            except:
                continue
                
        return None
        
    def render(self, code: str, output_path: str) -> bool:
        """使用mermaid-cli渲染图表"""
        if not self.mermaid_cli:
            raise Exception("未找到 mermaid-cli，请先安装: npm install -g @mermaid-js/mermaid-cli")
            
        # 创建临时mmd文件
        temp_mmd = os.path.join(self.temp_dir, "temp.mmd")
        with open(temp_mmd, 'w', encoding='utf-8') as f:
            f.write(code)
            
        try:
            # 执行渲染
            result = subprocess.run(
                [self.mermaid_cli, "-i", temp_mmd, "-o", output_path],
                capture_output=True,
                timeout=30
            )
            
            return result.returncode == 0
            
        finally:
            # 清理临时文件
            if os.path.exists(temp_mmd):
                os.remove(temp_mmd)


class PlantUMLRenderer(DiagramRenderer):
    """PlantUML图表渲染器"""
    
    def __init__(self):
        super().__init__()
        self.plantuml_jar = self._find_plantuml()
        
    def _find_plantuml(self) -> str:
        """查找plantuml.jar"""
        # 尝试常见位置
        possible_paths = [
            "plantuml.jar",
            os.path.join(os.getcwd(), "plantuml.jar"),
            os.path.join(os.path.expanduser("~"), "plantuml.jar")
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
                
        return None
        
    def render(self, code: str, output_path: str) -> bool:
        """使用PlantUML渲染图表"""
        if not self.plantuml_jar:
            raise Exception("未找到 plantuml.jar，请先下载并放置在项目目录")
            
        # 创建临时puml文件
        temp_puml = os.path.join(self.temp_dir, "temp.puml")
        with open(temp_puml, 'w', encoding='utf-8') as f:
            f.write(code)
            
        try:
            # 执行渲染
            result = subprocess.run(
                ["java", "-jar", self.plantuml_jar, "-tpng", temp_puml, "-o", os.path.dirname(output_path)],
                capture_output=True,
                timeout=30
            )
            
            # PlantUML会生成temp.png，需要重命名
            generated = os.path.join(os.path.dirname(output_path), "temp.png")
            if os.path.exists(generated):
                import shutil
                shutil.move(generated, output_path)
                return True
                
            return False
            
        finally:
            if os.path.exists(temp_puml):
                os.remove(temp_puml)


class DiagramFeature:
    """图表功能管理器"""
    
    def __init__(self, app):
        self.app = app
        self.mermaid_renderer = MermaidRenderer()
        self.plantuml_renderer = PlantUMLRenderer()
        self.diagram_dialog = None
        
    def show_diagram_editor(self):
        """显示图表编辑器"""
        if self.diagram_dialog and self.diagram_dialog.winfo_exists():
            self.diagram_dialog.focus()
            return
            
        self.diagram_dialog = ctk.CTkToplevel(self.app)
        self.diagram_dialog.title("📊 图表编辑器")
        self.diagram_dialog.geometry("1000x700")
        self.diagram_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.diagram_dialog,
            text="Mermaid / PlantUML 图表编辑器",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=15)
        
        # 类型选择
        type_frame = ctk.CTkFrame(self.diagram_dialog, fg_color="transparent")
        type_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(type_frame, text="图表类型:").pack(side="left", padx=5)
        
        self.diagram_type_var = ctk.StringVar(value="mermaid")
        type_selector = ctk.CTkSegmentedButton(
            type_frame,
            values=["Mermaid", "PlantUML"],
            variable=self.diagram_type_var,
            command=self._on_diagram_type_change
        )
        type_selector.pack(side="left", padx=10)
        
        # 模板选择
        ctk.CTkLabel(type_frame, text="快速模板:").pack(side="left", padx=(20, 5))
        
        self.template_var = ctk.StringVar(value="流程图")
        template_menu = ctk.CTkOptionMenu(
            type_frame,
            values=["流程图", "时序图", "类图", "状态图", "甘特图", "饼图"],
            variable=self.template_var,
            command=self._load_template
        )
        template_menu.pack(side="left", padx=5)
        
        # 主内容区域（分割为编辑器和预览）
        content_frame = ctk.CTkFrame(self.diagram_dialog)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 左侧：代码编辑器
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        editor_label = ctk.CTkLabel(
            left_frame,
            text="📝 图表代码",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        editor_label.pack(pady=5)
        
        self.diagram_code = ctk.CTkTextbox(
            left_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="none"
        )
        self.diagram_code.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 右侧：预览
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        preview_label = ctk.CTkLabel(
            right_frame,
            text="👁️ 预览",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        preview_label.pack(pady=5)
        
        self.diagram_preview = ctk.CTkLabel(
            right_frame,
            text="点击渲染按钮预览图表",
            fg_color="#F3F4F6"
        )
        self.diagram_preview.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(self.diagram_dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        render_btn = ctk.CTkButton(
            btn_frame,
            text="🎨 渲染预览",
            command=self._render_diagram,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=120
        )
        render_btn.pack(side="left", padx=5)
        
        insert_btn = ctk.CTkButton(
            btn_frame,
            text="➕ 插入到文档",
            command=self._insert_diagram,
            fg_color="#10B981",
            hover_color="#059669",
            width=140
        )
        insert_btn.pack(side="left", padx=5)
        
        save_btn = ctk.CTkButton(
            btn_frame,
            text="💾 保存图片",
            command=self._save_diagram,
            width=120
        )
        save_btn.pack(side="left", padx=5)
        
        close_btn = ctk.CTkButton(
            btn_frame,
            text="关闭",
            command=self.diagram_dialog.destroy,
            fg_color="#6B7280",
            hover_color="#4B5563",
            width=100
        )
        close_btn.pack(side="right", padx=5)
        
        # 加载默认模板
        self._load_template("流程图")
        
    def _on_diagram_type_change(self, value):
        """图表类型改变时"""
        # 根据类型更新模板列表
        if value.lower() == "mermaid":
            templates = ["流程图", "时序图", "类图", "状态图", "甘特图", "饼图"]
        else:  # PlantUML
            templates = ["时序图", "用例图", "类图", "活动图", "组件图", "状态图"]
            
        # TODO: 更新模板下拉菜单
        
    def _load_template(self, template_name):
        """加载模板"""
        templates = {
            "流程图": """graph TD
    A[开始] --> B{判断条件}
    B -->|是| C[执行操作1]
    B -->|否| D[执行操作2]
    C --> E[结束]
    D --> E""",
            
            "时序图": """sequenceDiagram
    participant A as 用户
    participant B as 服务器
    A->>B: 发送请求
    B->>B: 处理请求
    B-->>A: 返回结果""",
            
            "类图": """classDiagram
    class Animal {
        +String name
        +int age
        +eat()
        +sleep()
    }
    class Dog {
        +bark()
    }
    Animal <|-- Dog""",
            
            "状态图": """stateDiagram-v2
    [*] --> 待处理
    待处理 --> 处理中
    处理中 --> 已完成
    处理中 --> 失败
    已完成 --> [*]
    失败 --> 待处理""",
            
            "甘特图": """gantt
    title 项目进度
    dateFormat  YYYY-MM-DD
    section 阶段1
    任务1           :a1, 2024-01-01, 30d
    任务2           :after a1, 20d
    section 阶段2
    任务3           :2024-02-01, 25d""",
            
            "饼图": """pie title 数据分布
    "类别A" : 45
    "类别B" : 30
    "类别C" : 15
    "类别D" : 10"""
        }
        
        if template_name in templates:
            self.diagram_code.delete("1.0", "end")
            self.diagram_code.insert("1.0", templates[template_name])
            
    def _render_diagram(self):
        """渲染图表"""
        code = self.diagram_code.get("1.0", "end-1c")
        if not code.strip():
            messagebox.showwarning("提示", "请输入图表代码！")
            return
            
        try:
            diagram_type = self.diagram_type_var.get().lower()
            
            # 创建临时输出文件
            output_path = os.path.join(tempfile.gettempdir(), "diagram_preview.png")
            
            # 选择渲染器
            if diagram_type == "mermaid":
                success = self.mermaid_renderer.render(code, output_path)
            else:
                success = self.plantuml_renderer.render(code, output_path)
                
            if success and os.path.exists(output_path):
                # 加载并显示图片
                from PIL import Image
                import customtkinter
                
                img = Image.open(output_path)
                # 调整大小以适应预览区域
                img.thumbnail((400, 400), Image.Resampling.LANCZOS)
                
                ctk_image = customtkinter.CTkImage(
                    light_image=img,
                    dark_image=img,
                    size=(img.width, img.height)
                )
                
                self.diagram_preview.configure(image=ctk_image, text="")
                self.diagram_preview.image = ctk_image  # 保持引用
                
                self.app.update_status("✅ 图表渲染成功")
            else:
                raise Exception("渲染失败")
                
        except Exception as e:
            messagebox.showerror("渲染失败", f"无法渲染图表:\n{e}\n\n请确保已安装相应的渲染工具。")
            self.app.update_status(f"❌ 渲染失败: {e}")
            
    def _insert_diagram(self):
        """插入图表到文档"""
        code = self.diagram_code.get("1.0", "end-1c")
        if not code.strip():
            messagebox.showwarning("提示", "请输入图表代码！")
            return
            
        diagram_type = self.diagram_type_var.get().lower()
        
        # 插入代码块到编辑器
        markdown_code = f"```{diagram_type}\n{code}\n```\n"
        
        self.app.input_text.insert("insert", markdown_code)
        self.app.on_text_change(None)
        
        messagebox.showinfo("成功", "图表代码已插入到文档")
        self.app.update_status("✅ 图表已插入")
        
    def _save_diagram(self):
        """保存图表为图片"""
        code = self.diagram_code.get("1.0", "end-1c")
        if not code.strip():
            messagebox.showwarning("提示", "请输入图表代码！")
            return
            
        # 选择保存位置
        from tkinter import filedialog
        output_path = filedialog.asksaveasfilename(
            title="保存图表",
            defaultextension=".png",
            filetypes=[("PNG图片", "*.png"), ("所有文件", "*.*")]
        )
        
        if not output_path:
            return
            
        try:
            diagram_type = self.diagram_type_var.get().lower()
            
            if diagram_type == "mermaid":
                success = self.mermaid_renderer.render(code, output_path)
            else:
                success = self.plantuml_renderer.render(code, output_path)
                
            if success:
                messagebox.showinfo("成功", f"图表已保存到:\n{output_path}")
                self.app.update_status("✅ 图表已保存")
            else:
                raise Exception("保存失败")
                
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存图表:\n{e}")
            
    def parse_diagrams_in_markdown(self, markdown_text: str) -> list:
        """解析Markdown中的图表代码块"""
        diagrams = []
        
        # 匹配Mermaid和PlantUML代码块
        pattern = r'```(mermaid|plantuml)\n(.*?)\n```'
        matches = re.finditer(pattern, markdown_text, re.DOTALL)
        
        for match in matches:
            diagram_type = match.group(1)
            code = match.group(2)
            diagrams.append({
                'type': diagram_type,
                'code': code,
                'start': match.start(),
                'end': match.end()
            })
            
        return diagrams
