# -*- coding: utf-8 -*-
"""精确滚动同步模块 - 基于行映射的双向同步"""

import re
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class LineMapping:
    """行映射数据"""
    source_line: int
    preview_element_id: str
    preview_y_position: float
    block_type: str  # heading, paragraph, code_block, etc.
    block_end_line: int = 0  # 块结束行号


@dataclass
class BlockInfo:
    """块信息"""
    start_line: int
    end_line: int
    block_type: str
    content_hash: str = ""


class PreciseScrollSync:
    """精确滚动同步 - 基于行映射"""
    
    def __init__(self, editor, preview, app=None):
        """
        初始化精确滚动同步
        
        Args:
            editor: 编辑器组件 (LineNumberedText)
            preview: 预览组件 (MarkdownPreview)
            app: 主应用引用
        """
        self.editor = editor
        self.preview = preview
        self.app = app
        
        # 行映射表: source_line -> LineMapping
        self.line_map: Dict[int, LineMapping] = {}
        
        # 块信息列表（用于更精确的映射）
        self.blocks: List[BlockInfo] = []
        
        # 同步锁防止循环触发
        self._sync_lock = False
        self._last_sync_time = 0.0
        self._sync_cooldown = 0.05  # 50ms 冷却时间
        
        # 平滑滚动参数
        self._smooth_scroll_enabled = True
        self._smooth_scroll_duration = 150  # ms
        self._smooth_scroll_steps = 10
        
        # 缓存
        self._content_hash = ""
        self._total_lines = 0

    def build_line_map(self, content: str) -> None:
        """
        构建源码行到预览位置的映射表
        
        Args:
            content: Markdown 源码内容
        """
        # 检查内容是否变化（使用简单哈希）
        content_hash = str(hash(content))
        if content_hash == self._content_hash:
            return
        
        self._content_hash = content_hash
        self.line_map.clear()
        self.blocks.clear()
        
        if not content:
            self._total_lines = 0
            return
        
        lines = content.split('\n')
        self._total_lines = len(lines)
        current_line = 1
        
        # 解析 Markdown 块级元素
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # 空行
            if not stripped:
                i += 1
                current_line += 1
                continue
            
            # 标题
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', stripped)
            if heading_match:
                level = len(heading_match.group(1))
                self.line_map[current_line] = LineMapping(
                    source_line=current_line,
                    preview_element_id=f"heading_{current_line}",
                    preview_y_position=0.0,
                    block_type=f"h{level}",
                    block_end_line=current_line
                )
                self.blocks.append(BlockInfo(
                    start_line=current_line,
                    end_line=current_line,
                    block_type=f"h{level}"
                ))
                i += 1
                current_line += 1
                continue
            
            # 代码块
            if stripped.startswith('```'):
                block_start = current_line
                i += 1
                current_line += 1
                # 找到代码块结束
                while i < len(lines):
                    if lines[i].strip() == '```':
                        i += 1
                        current_line += 1
                        break
                    i += 1
                    current_line += 1
                
                block_end = current_line - 1
                self.line_map[block_start] = LineMapping(
                    source_line=block_start,
                    preview_element_id=f"code_block_{block_start}",
                    preview_y_position=0.0,
                    block_type="code_block",
                    block_end_line=block_end
                )
                self.blocks.append(BlockInfo(
                    start_line=block_start,
                    end_line=block_end,
                    block_type="code_block"
                ))
                continue
            
            # 数学公式块
            if stripped.startswith('$$'):
                block_start = current_line
                i += 1
                current_line += 1
                while i < len(lines):
                    if lines[i].strip().endswith('$$'):
                        i += 1
                        current_line += 1
                        break
                    i += 1
                    current_line += 1
                
                block_end = current_line - 1
                self.line_map[block_start] = LineMapping(
                    source_line=block_start,
                    preview_element_id=f"math_block_{block_start}",
                    preview_y_position=0.0,
                    block_type="math_block",
                    block_end_line=block_end
                )
                self.blocks.append(BlockInfo(
                    start_line=block_start,
                    end_line=block_end,
                    block_type="math_block"
                ))
                continue
            
            # 引用块
            if stripped.startswith('>'):
                block_start = current_line
                while i < len(lines) and lines[i].strip().startswith('>'):
                    i += 1
                    current_line += 1
                
                block_end = current_line - 1
                self.line_map[block_start] = LineMapping(
                    source_line=block_start,
                    preview_element_id=f"quote_{block_start}",
                    preview_y_position=0.0,
                    block_type="quote",
                    block_end_line=block_end
                )
                self.blocks.append(BlockInfo(
                    start_line=block_start,
                    end_line=block_end,
                    block_type="quote"
                ))
                continue
            
            # 表格
            if '|' in line and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1].strip()):
                block_start = current_line
                while i < len(lines) and '|' in lines[i]:
                    i += 1
                    current_line += 1
                
                block_end = current_line - 1
                self.line_map[block_start] = LineMapping(
                    source_line=block_start,
                    preview_element_id=f"table_{block_start}",
                    preview_y_position=0.0,
                    block_type="table",
                    block_end_line=block_end
                )
                self.blocks.append(BlockInfo(
                    start_line=block_start,
                    end_line=block_end,
                    block_type="table"
                ))
                continue
            
            # 列表项
            list_match = re.match(r'^(\s*)([-*+]|\d+\.)\s+', line)
            if list_match:
                self.line_map[current_line] = LineMapping(
                    source_line=current_line,
                    preview_element_id=f"list_{current_line}",
                    preview_y_position=0.0,
                    block_type="list_item",
                    block_end_line=current_line
                )
                self.blocks.append(BlockInfo(
                    start_line=current_line,
                    end_line=current_line,
                    block_type="list_item"
                ))
                i += 1
                current_line += 1
                continue
            
            # 水平线
            if re.match(r'^[-*_]{3,}\s*$', stripped):
                self.line_map[current_line] = LineMapping(
                    source_line=current_line,
                    preview_element_id=f"hr_{current_line}",
                    preview_y_position=0.0,
                    block_type="hr",
                    block_end_line=current_line
                )
                self.blocks.append(BlockInfo(
                    start_line=current_line,
                    end_line=current_line,
                    block_type="hr"
                ))
                i += 1
                current_line += 1
                continue
            
            # 普通段落
            para_start = current_line
            while i < len(lines) and lines[i].strip() and not self._is_block_start(lines[i]):
                i += 1
                current_line += 1
            
            para_end = current_line - 1 if current_line > para_start else para_start
            self.line_map[para_start] = LineMapping(
                source_line=para_start,
                preview_element_id=f"paragraph_{para_start}",
                preview_y_position=0.0,
                block_type="paragraph",
                block_end_line=para_end
            )
            self.blocks.append(BlockInfo(
                start_line=para_start,
                end_line=para_end,
                block_type="paragraph"
            ))
    
    def _is_block_start(self, line: str) -> bool:
        """检查是否是块级元素的开始"""
        stripped = line.strip()
        if not stripped:
            return True
        if stripped.startswith('#'):
            return True
        if stripped.startswith('```'):
            return True
        if stripped.startswith('$$'):
            return True
        if stripped.startswith('>'):
            return True
        if re.match(r'^[-*_]{3,}\s*$', stripped):
            return True
        if re.match(r'^(\s*)([-*+]|\d+\.)\s+', line):
            return True
        return False

    def sync_editor_to_preview(self, editor_line: int = None) -> None:
        """
        编辑器滚动时同步预览
        
        Args:
            editor_line: 编辑器当前可见的第一行（如果为 None，自动获取）
        """
        if self._sync_lock:
            return
        
        # 冷却时间检查
        now = time.monotonic()
        if now - self._last_sync_time < self._sync_cooldown:
            return
        
        self._sync_lock = True
        self._last_sync_time = now
        
        try:
            # 获取编辑器当前可见的第一行
            if editor_line is None:
                editor_line = self._get_editor_first_visible_line()
            
            if editor_line is None:
                return
            
            # 找到最近的映射行
            mapped_line = self._find_nearest_mapped_line(editor_line)
            if mapped_line is None:
                # 使用比例同步作为后备
                self._sync_by_ratio(editor_line)
                return
            
            # 计算预览区应该滚动到的位置
            preview_position = self._calculate_preview_position(mapped_line, editor_line)
            
            # 执行滚动
            if self._smooth_scroll_enabled:
                self._smooth_scroll_preview(preview_position)
            else:
                self._scroll_preview_to(preview_position)
                
        finally:
            self._sync_lock = False
    
    def sync_preview_to_editor(self, preview_pos: float) -> None:
        """
        预览滚动时同步编辑器
        
        Args:
            preview_pos: 预览区滚动位置 (0.0 - 1.0)
        """
        if self._sync_lock:
            return
        
        # 冷却时间检查
        now = time.monotonic()
        if now - self._last_sync_time < self._sync_cooldown:
            return
        
        self._sync_lock = True
        self._last_sync_time = now
        
        try:
            # 反向查找对应的源码行
            source_line = self._find_source_line_from_preview(preview_pos)
            
            if source_line:
                self._scroll_editor_to_line(source_line)
            else:
                # 使用比例同步作为后备
                self._sync_editor_by_ratio(preview_pos)
                
        finally:
            self._sync_lock = False

    def _get_editor_first_visible_line(self) -> Optional[int]:
        """获取编辑器当前可见的第一行"""
        try:
            # 获取底层 Text 组件
            text_widget = None
            if hasattr(self.editor, '_textbox'):
                text_widget = self.editor._textbox
            elif hasattr(self.editor, 'text'):
                text_widget = self.editor.text
            else:
                text_widget = self.editor
            
            # 获取可见区域的第一行
            first_visible = text_widget.index("@0,0")
            line_num = int(first_visible.split('.')[0])
            return line_num
        except Exception:
            return None
    
    def _find_nearest_mapped_line(self, target_line: int) -> Optional[LineMapping]:
        """找到最近的映射行"""
        if not self.line_map:
            return None
        
        # 精确匹配
        if target_line in self.line_map:
            return self.line_map[target_line]
        
        # 找最近的映射行（向上搜索）
        nearest_line = None
        min_distance = float('inf')
        
        for line_num in self.line_map:
            distance = abs(line_num - target_line)
            if distance < min_distance:
                min_distance = distance
                nearest_line = line_num
        
        if nearest_line is not None:
            return self.line_map[nearest_line]
        
        return None
    
    def _calculate_preview_position(self, mapping: LineMapping, editor_line: int) -> float:
        """计算预览区应该滚动到的位置"""
        try:
            if self._total_lines <= 0:
                return 0.0
            
            # 找到当前行所在的块
            current_block_idx = -1
            for idx, block in enumerate(self.blocks):
                if block.start_line <= editor_line <= block.end_line:
                    current_block_idx = idx
                    break
                elif block.start_line > editor_line:
                    # 在两个块之间，使用前一个块
                    current_block_idx = max(0, idx - 1)
                    break
            
            if current_block_idx < 0:
                current_block_idx = len(self.blocks) - 1 if self.blocks else 0
            
            # 基于块索引计算位置（更精确）
            if self.blocks:
                total_blocks = len(self.blocks)
                # 块内偏移
                block = self.blocks[current_block_idx]
                block_lines = block.end_line - block.start_line + 1
                if block_lines > 1 and block.start_line <= editor_line <= block.end_line:
                    intra_block_offset = (editor_line - block.start_line) / block_lines
                else:
                    intra_block_offset = 0.0
                
                # 计算总体位置
                position = (current_block_idx + intra_block_offset) / max(1, total_blocks)
            else:
                # 回退到简单比例计算
                position = (editor_line - 1) / max(1, self._total_lines - 1)
            
            return min(1.0, max(0.0, position))
            
        except Exception:
            return 0.0
    
    def _find_source_line_from_preview(self, preview_pos: float) -> Optional[int]:
        """从预览位置反向查找源码行"""
        if not self.blocks:
            return None
        
        try:
            # 基于块索引反向计算
            total_blocks = len(self.blocks)
            target_block_idx = int(preview_pos * total_blocks)
            target_block_idx = min(total_blocks - 1, max(0, target_block_idx))
            
            block = self.blocks[target_block_idx]
            
            # 计算块内偏移
            block_start_pos = target_block_idx / total_blocks
            block_end_pos = (target_block_idx + 1) / total_blocks
            
            if block_end_pos > block_start_pos:
                intra_block_ratio = (preview_pos - block_start_pos) / (block_end_pos - block_start_pos)
                intra_block_ratio = min(1.0, max(0.0, intra_block_ratio))
            else:
                intra_block_ratio = 0.0
            
            block_lines = block.end_line - block.start_line + 1
            target_line = block.start_line + int(intra_block_ratio * (block_lines - 1))
            
            return max(1, min(self._total_lines, target_line))
            
        except Exception:
            return None

    def _sync_by_ratio(self, editor_line: int) -> None:
        """使用比例同步（后备方案）"""
        try:
            text_widget = None
            if hasattr(self.editor, '_textbox'):
                text_widget = self.editor._textbox
            elif hasattr(self.editor, 'text'):
                text_widget = self.editor.text
            else:
                text_widget = self.editor
            
            total_lines = int(text_widget.index('end').split('.')[0]) - 1
            
            if total_lines <= 0:
                return
            
            position = (editor_line - 1) / max(1, total_lines - 1)
            position = min(1.0, max(0.0, position))
            
            self._scroll_preview_to(position)
            
        except Exception:
            pass
    
    def _sync_editor_by_ratio(self, preview_pos: float) -> None:
        """使用比例同步编辑器（后备方案）"""
        try:
            text_widget = None
            if hasattr(self.editor, '_textbox'):
                text_widget = self.editor._textbox
            elif hasattr(self.editor, 'text'):
                text_widget = self.editor.text
            else:
                text_widget = self.editor
            
            text_widget.yview_moveto(preview_pos)
            
        except Exception:
            pass
    
    def _scroll_preview_to(self, position: float) -> None:
        """滚动预览区到指定位置"""
        try:
            if hasattr(self.preview, 'text'):
                self.preview.text.yview_moveto(position)
            elif hasattr(self.preview, 'sync_scroll_to'):
                self.preview.sync_scroll_to(position)
        except Exception:
            pass
    
    def _scroll_editor_to_line(self, line: int) -> None:
        """滚动编辑器到指定行"""
        try:
            text_widget = None
            if hasattr(self.editor, '_textbox'):
                text_widget = self.editor._textbox
            elif hasattr(self.editor, 'text'):
                text_widget = self.editor.text
            else:
                text_widget = self.editor
            
            text_widget.see(f"{line}.0")
            
        except Exception:
            pass
    
    def _smooth_scroll_preview(self, target_position: float) -> None:
        """平滑滚动预览区"""
        try:
            if not hasattr(self.preview, 'text'):
                self._scroll_preview_to(target_position)
                return
            
            # 获取当前位置
            current_pos = self.preview.text.yview()[0]
            
            # 如果差距很小，直接跳转
            if abs(target_position - current_pos) < 0.01:
                self._scroll_preview_to(target_position)
                return
            
            # 计算每步的增量
            step_count = self._smooth_scroll_steps
            step_delay = self._smooth_scroll_duration // step_count
            delta = (target_position - current_pos) / step_count
            
            def animate_step(step: int, pos: float):
                if step >= step_count:
                    self._scroll_preview_to(target_position)
                    return
                
                new_pos = pos + delta
                self._scroll_preview_to(new_pos)
                
                if self.app:
                    self.app.after(step_delay, lambda: animate_step(step + 1, new_pos))
            
            if self.app:
                animate_step(0, current_pos)
            else:
                self._scroll_preview_to(target_position)
                
        except Exception:
            self._scroll_preview_to(target_position)

    def set_smooth_scroll(self, enabled: bool, duration_ms: int = 150) -> None:
        """设置平滑滚动"""
        self._smooth_scroll_enabled = enabled
        self._smooth_scroll_duration = duration_ms
    
    def get_sync_accuracy(self, editor_line: int, preview_line: int) -> int:
        """
        获取同步精确度（行数误差）
        
        Args:
            editor_line: 编辑器当前行
            preview_line: 预览区对应行
            
        Returns:
            行数误差（绝对值）
        """
        return abs(editor_line - preview_line)
    
    def is_sync_accurate(self, editor_line: int, preview_line: int, tolerance: int = 2) -> bool:
        """
        检查同步是否精确
        
        Args:
            editor_line: 编辑器当前行
            preview_line: 预览区对应行
            tolerance: 允许的误差行数
            
        Returns:
            是否在允许误差范围内
        """
        return self.get_sync_accuracy(editor_line, preview_line) <= tolerance


class IncrementalPreviewUpdater:
    """增量预览更新器 - 只更新变化的块"""
    
    def __init__(self, preview):
        """
        初始化增量预览更新器
        
        Args:
            preview: 预览组件
        """
        self.preview = preview
        self.block_cache: Dict[str, str] = {}
        self._last_content = ""
        self._last_blocks: Dict[int, Tuple[str, str]] = {}
        
        # 块类型权重（用于优化更新顺序）
        self._block_weights = {
            'h1': 10, 'h2': 9, 'h3': 8, 'h4': 7, 'h5': 6, 'h6': 5,
            'code_block': 8,
            'table': 7,
            'quote': 5,
            'list': 4,
            'paragraph': 3,
        }
    
    def update(self, old_content: str, new_content: str) -> List[Tuple[str, str]]:
        """
        增量更新预览
        
        Args:
            old_content: 旧内容
            new_content: 新内容
            
        Returns:
            变化的块列表 [(block_id, block_content), ...]
        """
        if old_content == new_content:
            return []
        
        old_blocks = self._parse_blocks(old_content)
        new_blocks = self._parse_blocks(new_content)
        
        # 找出变化的块
        changed_blocks = self._diff_blocks(old_blocks, new_blocks)
        
        # 缓存新的块信息
        self._last_blocks = new_blocks
        self._last_content = new_content
        
        return changed_blocks
    
    def should_full_render(self, old_content: str, new_content: str) -> bool:
        """
        判断是否需要全量渲染
        
        当变化太大时，全量渲染可能比增量更新更高效
        
        Args:
            old_content: 旧内容
            new_content: 新内容
            
        Returns:
            是否需要全量渲染
        """
        if not old_content or not new_content:
            return True
        
        # 如果内容长度变化超过 50%，使用全量渲染
        len_diff = abs(len(new_content) - len(old_content))
        if len_diff > len(old_content) * 0.5:
            return True
        
        # 如果行数变化超过 30%，使用全量渲染
        old_lines = old_content.count('\n')
        new_lines = new_content.count('\n')
        if old_lines > 0:
            line_diff = abs(new_lines - old_lines)
            if line_diff > old_lines * 0.3:
                return True
        
        return False
    
    def get_changed_line_range(self, old_content: str, new_content: str) -> Tuple[int, int]:
        """
        获取变化的行范围
        
        Args:
            old_content: 旧内容
            new_content: 新内容
            
        Returns:
            (start_line, end_line) 变化的行范围
        """
        old_lines = old_content.split('\n')
        new_lines = new_content.split('\n')
        
        # 从头部找第一个不同的行
        start_line = 0
        min_len = min(len(old_lines), len(new_lines))
        for i in range(min_len):
            if old_lines[i] != new_lines[i]:
                start_line = i + 1
                break
        else:
            start_line = min_len + 1
        
        # 从尾部找最后一个不同的行
        end_line = len(new_lines)
        for i in range(1, min_len + 1):
            if old_lines[-i] != new_lines[-i]:
                end_line = len(new_lines) - i + 1
                break
        
        return (start_line, end_line)
    
    def _parse_blocks(self, content: str) -> Dict[int, Tuple[str, str]]:
        """
        解析内容为块
        
        Returns:
            {line_num: (block_type, block_content)}
        """
        blocks = {}
        if not content:
            return blocks
        
        lines = content.split('\n')
        i = 0
        line_num = 1
        
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # 空行
            if not stripped:
                i += 1
                line_num += 1
                continue
            
            # 代码块
            if stripped.startswith('```'):
                block_start = line_num
                block_lines = [line]
                i += 1
                line_num += 1
                
                while i < len(lines):
                    block_lines.append(lines[i])
                    if lines[i].strip() == '```':
                        i += 1
                        line_num += 1
                        break
                    i += 1
                    line_num += 1
                
                blocks[block_start] = ('code_block', '\n'.join(block_lines))
                continue
            
            # 数学公式块
            if stripped.startswith('$$'):
                block_start = line_num
                block_lines = [line]
                i += 1
                line_num += 1
                
                while i < len(lines):
                    block_lines.append(lines[i])
                    if lines[i].strip().endswith('$$'):
                        i += 1
                        line_num += 1
                        break
                    i += 1
                    line_num += 1
                
                blocks[block_start] = ('math_block', '\n'.join(block_lines))
                continue
            
            # 表格
            if '|' in line and i + 1 < len(lines) and re.match(r'^[\s\|\:\-]+$', lines[i + 1].strip()):
                block_start = line_num
                block_lines = [line]
                i += 1
                line_num += 1
                
                while i < len(lines) and '|' in lines[i]:
                    block_lines.append(lines[i])
                    i += 1
                    line_num += 1
                
                blocks[block_start] = ('table', '\n'.join(block_lines))
                continue
            
            # 引用块
            if stripped.startswith('>'):
                block_start = line_num
                block_lines = [line]
                i += 1
                line_num += 1
                
                while i < len(lines) and lines[i].strip().startswith('>'):
                    block_lines.append(lines[i])
                    i += 1
                    line_num += 1
                
                blocks[block_start] = ('quote', '\n'.join(block_lines))
                continue
            
            # 标题
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', stripped)
            if heading_match:
                level = len(heading_match.group(1))
                blocks[line_num] = (f'h{level}', line)
                i += 1
                line_num += 1
                continue
            
            # 列表项
            list_match = re.match(r'^(\s*)([-*+]|\d+\.)\s+', line)
            if list_match:
                block_start = line_num
                block_lines = [line]
                i += 1
                line_num += 1
                
                # 收集连续的列表项
                while i < len(lines):
                    next_line = lines[i]
                    if re.match(r'^(\s*)([-*+]|\d+\.)\s+', next_line) or (next_line.startswith('  ') and next_line.strip()):
                        block_lines.append(next_line)
                        i += 1
                        line_num += 1
                    elif not next_line.strip():
                        # 空行可能是列表项之间的分隔
                        if i + 1 < len(lines) and re.match(r'^(\s*)([-*+]|\d+\.)\s+', lines[i + 1]):
                            block_lines.append(next_line)
                            i += 1
                            line_num += 1
                        else:
                            break
                    else:
                        break
                
                blocks[block_start] = ('list', '\n'.join(block_lines))
                continue
            
            # 水平线
            if re.match(r'^[-*_]{3,}\s*$', stripped):
                blocks[line_num] = ('hr', line)
                i += 1
                line_num += 1
                continue
            
            # 普通段落
            blocks[line_num] = ('paragraph', line)
            i += 1
            line_num += 1
        
        return blocks
    
    def _diff_blocks(self, old_blocks: Dict, new_blocks: Dict) -> List[Tuple[str, str]]:
        """比较块差异"""
        changed = []
        
        # 检查新增和修改的块
        for line_num, (block_type, content) in new_blocks.items():
            block_id = f"{block_type}_{line_num}"
            
            if line_num not in old_blocks:
                # 新增块
                changed.append((block_id, content))
            elif old_blocks[line_num] != (block_type, content):
                # 修改的块
                changed.append((block_id, content))
        
        # 按块类型权重排序（重要的块优先更新）
        changed.sort(key=lambda x: -self._block_weights.get(x[0].split('_')[0], 0))
        
        return changed
    
    def get_cached_block(self, block_id: str) -> Optional[str]:
        """获取缓存的块内容"""
        return self.block_cache.get(block_id)
    
    def cache_block(self, block_id: str, content: str) -> None:
        """缓存块内容"""
        self.block_cache[block_id] = content
    
    def clear_cache(self) -> None:
        """清空缓存"""
        self.block_cache.clear()
        self._last_content = ""
        self._last_blocks.clear()
