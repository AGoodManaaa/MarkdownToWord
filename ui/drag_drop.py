# -*- coding: utf-8 -*-

import os
import shutil
import time
from typing import List
from tkinter import messagebox


def _clean_drop_paths(app, data: str) -> List[str]:
    try:
        parts = list(app.tk.splitlist(data or ''))
    except Exception:
        parts = [data or '']

    clean: List[str] = []
    for p in parts:
        fp = str(p or '').strip()
        if fp.startswith('{') and fp.endswith('}'):
            fp = fp[1:-1]
        if fp:
            clean.append(fp)
    return clean


def handle_drop_for_app(app, event) -> None:
    """处理拖拽导入事件：支持文件打开和图片智能插入。"""
    paths = _clean_drop_paths(app, getattr(event, 'data', '') or '')
    if not paths:
        return

    # 分类处理：Markdown文件 vs 图片文件
    md_extensions = ('.md', '.markdown', '.txt')
    img_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg')
    
    md_files = [p for p in paths if str(p).lower().endswith(md_extensions)]
    img_files = [p for p in paths if str(p).lower().endswith(img_extensions)]

    if img_files:
        _handle_image_drop(app, img_files)
        return

    if not md_files:
        messagebox.showwarning('提示', '请拖拽Markdown文件或图片文件')
        return

    # 打开第一个文件，其余仅加入最近文件
    try:
        app.file_ops.load_file(md_files[0])
    except Exception:
        pass

    for fp in md_files[1:]:
        try:
            app.file_ops.add_recent_file(fp)
        except Exception:
            pass

    if len(md_files) > 1:
        try:
            app.update_status(f'✅ 已导入 {len(md_files)} 个文件（已打开第 1 个）')
        except Exception:
            pass

def _handle_image_drop(app, img_paths: List[str]) -> None:
    """图片拖拽处理：自动保存到 assets 目录并插入 Markdown 代码"""
    if not hasattr(app, 'current_file') or not app.current_file:
        # 如果当前没有打开文件，则询问保存位置
        messagebox.showinfo('提示', '请先保存 Markdown 文件，以便确定图片 assets 目录的位置。')
        return

    # 确定 assets 目录
    doc_dir = os.path.dirname(app.current_file)
    assets_dir = os.path.join(doc_dir, 'assets')
    
    if not os.path.exists(assets_dir):
        try:
            os.makedirs(assets_dir)
        except Exception as e:
            messagebox.showerror('错误', f'无法创建 assets 目录: {e}')
            return

    inserted_count = 0
    for img_path in img_paths:
        try:
            # 生成新文件名，避免冲突
            ext = os.path.splitext(img_path)[1]
            timestamp = int(time.time() * 1000)
            new_filename = f"image_{timestamp}_{inserted_count}{ext}"
            dest_path = os.path.join(assets_dir, new_filename)
            
            # 复制图片
            shutil.copy2(img_path, dest_path)
            
            # 插入 Markdown 代码
            # 相对路径
            rel_path = f"assets/{new_filename}"
            md_code = f"![图片]({rel_path})\n"
            
            if hasattr(app, 'insert_text'):
                app.insert_text(md_code)
                inserted_count += 1
        except Exception as e:
            print(f"Error copying image {img_path}: {e}")

    if inserted_count > 0:
        app.update_status(f'✅ 已成功插入 {inserted_count} 张图片到 assets 目录')
