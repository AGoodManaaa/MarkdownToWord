# -*- coding: utf-8 -*-
"""
Word to Markdown 反向转换功能
"""

import os
from tkinter import filedialog, messagebox
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


class WordToMarkdownConverter:
    """Word转Markdown转换器"""
    
    def __init__(self, output_image_dir: str = None):
        """
        初始化转换器
        
        Args:
            output_image_dir: 图片输出目录（用于保存Word中的图片）
        """
        self.output_image_dir = output_image_dir
        self.image_counter = 0
        
    def convert_file(self, docx_path: str, md_path: str) -> bool:
        """
        转换Word文件为Markdown
        
        Args:
            docx_path: Word文件路径
            md_path: 输出Markdown文件路径
            
        Returns:
            是否成功
        """
        try:
            doc = Document(docx_path)
            markdown_text = self.convert_document(doc, md_path)
            
            # 保存Markdown文件
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(markdown_text)
                
            return True
            
        except Exception as e:
            print(f"转换失败: {e}")
            return False
            
    def convert_document(self, doc: Document, md_path: str = None) -> str:
        """
        转换Document对象为Markdown文本
        
        Args:
            doc: Document对象
            md_path: Markdown文件路径（用于确定图片相对路径）
            
        Returns:
            Markdown文本
        """
        # 设置图片输出目录
        if md_path and not self.output_image_dir:
            md_dir = os.path.dirname(md_path)
            md_basename = os.path.splitext(os.path.basename(md_path))[0]
            self.output_image_dir = os.path.join(md_dir, f"{md_basename}_images")
            
        markdown_lines = []
        
        for element in doc.element.body:
            if isinstance(element, CT_P):
                # 段落
                para = Paragraph(element, doc)
                md_line = self._convert_paragraph(para)
                if md_line:
                    markdown_lines.append(md_line)
                    
            elif isinstance(element, CT_Tbl):
                # 表格
                table = Table(element, doc)
                md_table = self._convert_table(table)
                if md_table:
                    markdown_lines.append(md_table)
                    markdown_lines.append("")  # 空行
                    
        return "\n".join(markdown_lines)
        
    def _convert_paragraph(self, para: Paragraph) -> str:
        """转换段落为Markdown"""
        # 检查样式
        style_name = para.style.name if para.style else ""
        
        # 标题
        if style_name.startswith("Heading"):
            level = 1
            try:
                level = int(style_name.split()[-1])
            except:
                level = 1
            return f"{'#' * level} {self._convert_runs(para)}"
            
        # 代码块
        if "Code" in style_name:
            return f"```\n{para.text}\n```"
            
        # 引用
        if "Quote" in style_name:
            text = self._convert_runs(para)
            return f"> {text}" if text else ""
            
        # 列表项
        if para._element.pPr is not None:
            numPr = para._element.pPr.numPr
            if numPr is not None:
                # 有序或无序列表
                ilvl = numPr.ilvl
                level = int(ilvl.val) if ilvl is not None else 0
                indent = "  " * level
                
                # 判断是有序还是无序
                # 简化处理：默认都用无序列表
                return f"{indent}- {self._convert_runs(para)}"
                
        # 普通段落
        text = self._convert_runs(para)
        return text if text else ""
        
    def _convert_runs(self, para: Paragraph) -> str:
        """转换run为Markdown（处理粗体、斜体等）"""
        result = []
        
        for run in para.runs:
            text = run.text
            if not text:
                continue
                
            # 粗体
            if run.bold:
                text = f"**{text}**"
                
            # 斜体
            if run.italic:
                text = f"*{text}*"
                
            # 删除线
            if run.font.strike:
                text = f"~~{text}~~"
                
            # 行内代码
            if run.font.name and "Consolas" in run.font.name or "Courier" in run.font.name:
                text = f"`{text}`"
                
            result.append(text)
            
        return "".join(result)
        
    def _convert_table(self, table: Table) -> str:
        """转换表格为Markdown"""
        if not table.rows:
            return ""
            
        markdown_rows = []
        
        # 表头
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        markdown_rows.append("| " + " | ".join(header_cells) + " |")
        
        # 分隔线
        markdown_rows.append("|" + "|".join(["---"] * len(header_cells)) + "|")
        
        # 数据行
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            markdown_rows.append("| " + " | ".join(cells) + " |")
            
        return "\n".join(markdown_rows)
        
    def _extract_images(self, doc: Document) -> list:
        """提取Word中的图片（TODO: 实现图片提取）"""
        # 这需要处理document的rels和图片元素
        # 暂时留作TODO
        return []


class WordToMarkdownFeature:
    """Word转Markdown功能管理器"""
    
    def __init__(self, app):
        self.app = app
        
    def convert_word_file(self):
        """转换Word文件为Markdown"""
        # 选择Word文件
        word_file = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[
                ("Word文档", "*.docx"),
                ("所有文件", "*.*")
            ]
        )
        
        if not word_file:
            return
            
        # 选择输出路径
        default_name = os.path.splitext(os.path.basename(word_file))[0] + ".md"
        md_file = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            initialfile=default_name,
            filetypes=[("Markdown文件", "*.md"), ("所有文件", "*.*")]
        )
        
        if not md_file:
            return
            
        try:
            self.app.update_status("正在转换Word文档...")
            
            # 执行转换
            converter = WordToMarkdownConverter()
            success = converter.convert_file(word_file, md_file)
            
            if success:
                # 询问是否打开
                result = messagebox.askyesno(
                    "转换成功",
                    f"Word文档已转换为Markdown！\n\n是否在编辑器中打开？"
                )
                
                if result:
                    # 在编辑器中打开
                    with open(md_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    self.app.input_text.delete("1.0", "end")
                    self.app.input_text.insert("1.0", content)
                    self.app.current_file = md_file
                    self.app._update_title()
                    self.app.on_text_change(None)
                    
                self.app.update_status("✅ Word转Markdown成功")
            else:
                messagebox.showerror("错误", "转换失败！")
                self.app.update_status("❌ 转换失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {e}")
            self.app.update_status(f"❌ 转换失败: {e}")
