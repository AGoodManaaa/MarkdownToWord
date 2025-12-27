# -*- coding: utf-8 -*-
"""PDF 导出功能模块

通过 Word COM 接口将 Markdown 内容导出为 PDF 格式。
"""

import os
import tempfile
from tkinter import filedialog, messagebox
from typing import Optional


class PDFExportFeature:
    """PDF 导出功能类
    
    使用 Word COM 接口将 Markdown 转换为 PDF：
    1. 先使用现有转换器生成临时 docx
    2. 使用 win32com 打开 Word 并另存为 PDF
    3. 删除临时 docx 文件
    """
    
    def __init__(self, app):
        """初始化 PDF 导出功能
        
        Args:
            app: 主应用实例，用于访问编辑器内容和状态栏等
        """
        self.app = app
    
    def export_to_pdf(self) -> None:
        """导出当前内容为 PDF
        
        完整导出流程：
        1. 获取编辑器内容
        2. 显示保存对话框
        3. 生成临时 Word 文档
        4. 转换为 PDF
        5. 清理临时文件
        """
        # 获取编辑器内容
        try:
            content = self.app.input_text.get("1.0", "end-1c")
        except Exception:
            content = ""
        
        if not content.strip():
            messagebox.showwarning("提示", "没有可导出的内容")
            return
        
        # 显示保存对话框
        pdf_path = self._show_export_dialog()
        if not pdf_path:
            return
        
        # 更新状态
        self.app.update_status("📄 正在导出 PDF...")
        
        try:
            # 创建临时 docx 文件
            temp_dir = tempfile.gettempdir()
            temp_docx = os.path.join(temp_dir, f"temp_export_{os.getpid()}.docx")
            
            # 使用现有转换器生成 Word 文档
            from converter import MarkdownToWordConverter
            converter = MarkdownToWordConverter()
            
            # 获取页面大小设置
            page_size = self.app.config.get('page_size', 'A4')
            converter.convert(content, temp_docx, page_size=page_size)
            
            # 转换为 PDF
            success = self._convert_docx_to_pdf(temp_docx, pdf_path)
            
            # 清理临时文件
            try:
                if os.path.exists(temp_docx):
                    os.remove(temp_docx)
            except Exception:
                pass
            
            if success:
                self.app.update_status(f"✅ PDF 导出成功: {os.path.basename(pdf_path)}")
                messagebox.showinfo("导出成功", f"PDF 已保存到:\n{pdf_path}")
            else:
                self.app.update_status("❌ PDF 导出失败")
                
        except Exception as e:
            self.app.update_status("❌ PDF 导出失败")
            messagebox.showerror("导出失败", f"导出 PDF 时出错:\n{str(e)}")
    
    def _convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """使用 Word COM 将 docx 转换为 PDF
        
        Args:
            docx_path: Word 文档路径
            pdf_path: 目标 PDF 路径
            
        Returns:
            bool: 转换是否成功
        """
        word = None
        doc = None
        
        try:
            import win32com.client
            
            # 启动 Word 应用
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # 打开文档
            doc = word.Documents.Open(os.path.abspath(docx_path))
            
            # 另存为 PDF (wdFormatPDF = 17)
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            
            return True
            
        except ImportError:
            messagebox.showerror(
                "缺少依赖",
                "PDF 导出需要 pywin32 库。\n请运行: pip install pywin32"
            )
            return False
            
        except Exception as e:
            error_msg = str(e)
            if "Word.Application" in error_msg or "无法创建" in error_msg:
                messagebox.showerror(
                    "Word 未安装",
                    "PDF 导出需要安装 Microsoft Word。\n请安装 Word 后重试。"
                )
            else:
                messagebox.showerror("转换失败", f"Word 转 PDF 失败:\n{error_msg}")
            return False
            
        finally:
            # 关闭文档和 Word
            try:
                if doc:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word:
                    word.Quit()
            except Exception:
                pass
    
    def _show_export_dialog(self) -> Optional[str]:
        """显示保存对话框，返回用户选择的路径
        
        Returns:
            Optional[str]: 用户选择的文件路径，取消则返回 None
        """
        # 默认文件名
        default_name = "untitled.pdf"
        if self.app.current_file:
            base_name = os.path.splitext(os.path.basename(self.app.current_file))[0]
            default_name = f"{base_name}.pdf"
        
        file_path = filedialog.asksaveasfilename(
            title="导出为 PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        return file_path if file_path else None
