# -*- coding: utf-8 -*-
"""
增强的撤销/重做功能
支持多次撤销，每个小操作（输入字符、删除字符）都可以单独撤销
"""

import tkinter as tk
from typing import List, Optional


class TextOperation:
    """文本操作记录"""
    
    def __init__(self, op_type: str, position: str, content: str, tags: tuple = ()):
        """
        Args:
            op_type: 操作类型 'insert' 或 'delete'
            position: 操作位置
            content: 操作的内容
            tags: 文本标签
        """
        self.op_type = op_type
        self.position = position
        self.content = content
        self.tags = tags


class BatchOperation:
    """批量操作记录 - 将多个操作合并为一个撤销点"""
    
    def __init__(self, operations: List['TextOperation'] = None):
        """
        Args:
            operations: 操作列表
        """
        self.operations = operations or []
        self.op_type = 'batch'  # 标识为批量操作
    
    def add(self, operation: 'TextOperation'):
        """添加操作到批量"""
        self.operations.append(operation)
    
    def is_empty(self) -> bool:
        """是否为空"""
        return len(self.operations) == 0


class UndoRedoManager:
    """撤销/重做管理器"""
    
    def __init__(self, text_widget, max_undo: int = 100):
        """
        Args:
            text_widget: Text widget实例
            max_undo: 最大撤销次数
        """
        self.text_widget = text_widget
        self.max_undo = max_undo
        
        # 撤销栈和重做栈
        self.undo_stack: List[TextOperation] = []
        self.redo_stack: List[TextOperation] = []
        
        # 是否正在执行撤销/重做（避免死循环）
        self.is_undoing = False
        self.is_redoing = False
        
        # 默认启用撤销记录
        self.enabled = True
        
        # 操作记录缓冲
        self.last_insert_pos = None
        self.last_delete_pos = None
        
        # 批量操作支持
        self._batch_mode = False
        self._current_batch: Optional[BatchOperation] = None
    
    def begin_batch(self):
        """
        开始批量操作模式
        在此模式下，所有操作将被收集到一个批量中，
        撤销时作为一个整体撤销
        """
        if not self._batch_mode:
            self._batch_mode = True
            self._current_batch = BatchOperation()
            print("[BATCH] 开始批量操作模式")
    
    def end_batch(self):
        """
        结束批量操作模式
        将收集的操作作为一个整体添加到撤销栈
        """
        if self._batch_mode and self._current_batch:
            if not self._current_batch.is_empty():
                # 清空重做栈
                self.redo_stack.clear()
                # 添加批量操作到撤销栈
                self.undo_stack.append(self._current_batch)
                # 限制栈大小
                if len(self.undo_stack) > self.max_undo:
                    self.undo_stack.pop(0)
                print(f"[BATCH] 结束批量操作，包含 {len(self._current_batch.operations)} 个操作")
            else:
                print("[BATCH] 结束批量操作，无操作记录")
            self._batch_mode = False
            self._current_batch = None
    
    def cancel_batch(self):
        """取消当前批量操作，不保存"""
        if self._batch_mode:
            print("[BATCH] 取消批量操作")
            self._batch_mode = False
            self._current_batch = None
    
    @property
    def in_batch_mode(self) -> bool:
        """是否处于批量操作模式"""
        return self._batch_mode
        
    def enable(self):
        """启用撤销记录"""
        self.enabled = True
        
    def disable(self):
        """禁用撤销记录"""
        self.enabled = False
        
    def record_insert(self, position: str, content: str):
        """
        记录插入操作
        
        Args:
            position: 插入位置
            content: 插入内容
        """
        print(f"[DEBUG] record_insert 被调用: pos={position}, content='{content}', enabled={self.enabled}, undoing={self.is_undoing}, redoing={self.is_redoing}")
        
        if not self.enabled or self.is_undoing or self.is_redoing:
            print(f"[DEBUG] record_insert 跳过记录")
            return
            
        if not content:  # 忽略空内容
            print(f"[DEBUG] record_insert 内容为空，跳过")
            return
            
        op = TextOperation('insert', position, content)
        self._add_operation(op)
        print(f"[DEBUG] record_insert 已记录，当前栈大小: {len(self.undo_stack)}")
        
    def record_delete(self, position: str, content: str):
        """
        记录删除操作
        
        Args:
            position: 删除位置
            content: 删除的内容
        """
        print(f"[DEBUG] record_delete 被调用: pos={position}, content='{content}', enabled={self.enabled}")
        
        if not self.enabled or self.is_undoing or self.is_redoing:
            print(f"[DEBUG] record_delete 跳过记录")
            return
            
        if not content:  # 忽略空内容
            print(f"[DEBUG] record_delete 内容为空，跳过")
            return
            
        op = TextOperation('delete', position, content)
        self._add_operation(op)
        print(f"[DEBUG] record_delete 已记录，当前栈大小: {len(self.undo_stack)}")
        
    def _add_operation(self, operation: TextOperation):
        """添加操作到撤销栈"""
        # 如果在批量模式，添加到当前批量
        if self._batch_mode and self._current_batch:
            self._current_batch.add(operation)
            return
        
        # 添加新操作时，清空重做栈
        self.redo_stack.clear()
        
        # 添加到撤销栈
        self.undo_stack.append(operation)
        
        # 限制栈大小
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
            
    def _undo_single_operation(self, operation: TextOperation) -> Optional[str]:
        """
        撤销单个操作
        
        Args:
            operation: 要撤销的操作
            
        Returns:
            操作位置（用于设置光标）
        """
        if operation.op_type == 'insert':
            # 撤销插入 = 删除
            end_pos = f"{operation.position} + {len(operation.content)} chars"
            self.text_widget.delete(operation.position, end_pos)
        else:  # delete
            # 撤销删除 = 插入
            self.text_widget.insert(operation.position, operation.content)
        return operation.position
    
    def _redo_single_operation(self, operation: TextOperation) -> Optional[str]:
        """
        重做单个操作
        
        Args:
            operation: 要重做的操作
            
        Returns:
            操作位置（用于设置光标）
        """
        if operation.op_type == 'insert':
            self.text_widget.insert(operation.position, operation.content)
        else:  # delete
            end_pos = f"{operation.position} + {len(operation.content)} chars"
            self.text_widget.delete(operation.position, end_pos)
        return operation.position

    def undo(self) -> bool:
        """
        执行撤销
        
        Returns:
            是否成功撤销
        """
        if not self.undo_stack:
            return False
            
        self.is_undoing = True
        
        try:
            # 从撤销栈弹出操作
            operation = self.undo_stack.pop()
            
            # 临时禁用Text widget的undo
            old_undo_state = None
            try:
                old_undo_state = self.text_widget.cget('undo')
                self.text_widget.configure(undo=False)
            except:
                pass
            
            cursor_pos = None
            try:
                # 检查是否为批量操作
                if isinstance(operation, BatchOperation):
                    # 批量操作：逆序撤销所有子操作
                    for sub_op in reversed(operation.operations):
                        cursor_pos = self._undo_single_operation(sub_op)
                    print(f"[UNDO] 批量撤销 {len(operation.operations)} 个操作")
                else:
                    # 单个操作
                    cursor_pos = self._undo_single_operation(operation)
                    
                # 添加到重做栈
                self.redo_stack.append(operation)
                
                # 移动光标到操作位置
                if cursor_pos:
                    try:
                        self.text_widget.mark_set('insert', cursor_pos)
                        self.text_widget.see('insert')
                    except:
                        pass
            finally:
                # 恢复Text widget的undo状态
                if old_undo_state is not None:
                    try:
                        self.text_widget.configure(undo=old_undo_state)
                    except:
                        pass
            
            return True
            
        finally:
            self.is_undoing = False
            
    def redo(self) -> bool:
        """
        执行重做
        
        Returns:
            是否成功重做
        """
        if not self.redo_stack:
            return False
            
        self.is_redoing = True
        
        try:
            # 从重做栈弹出操作
            operation = self.redo_stack.pop()
            
            # 临时禁用Text widget的undo
            old_undo_state = None
            try:
                old_undo_state = self.text_widget.cget('undo')
                self.text_widget.configure(undo=False)
            except:
                pass
            
            cursor_pos = None
            try:
                # 检查是否为批量操作
                if isinstance(operation, BatchOperation):
                    # 批量操作：正序重做所有子操作
                    for sub_op in operation.operations:
                        cursor_pos = self._redo_single_operation(sub_op)
                    print(f"[REDO] 批量重做 {len(operation.operations)} 个操作")
                else:
                    # 单个操作
                    cursor_pos = self._redo_single_operation(operation)
                    
                # 添加回撤销栈
                self.undo_stack.append(operation)
                
                # 移动光标到操作位置
                if cursor_pos:
                    try:
                        self.text_widget.mark_set('insert', cursor_pos)
                        self.text_widget.see('insert')
                    except:
                        pass
            finally:
                # 恢复Text widget的undo状态
                if old_undo_state is not None:
                    try:
                        self.text_widget.configure(undo=old_undo_state)
                    except:
                        pass
            
            return True
            
        finally:
            self.is_redoing = False
            
    def can_undo(self) -> bool:
        """是否可以撤销"""
        return len(self.undo_stack) > 0
        
    def can_redo(self) -> bool:
        """是否可以重做"""
        return len(self.redo_stack) > 0
        
    def clear(self):
        """清空撤销/重做历史"""
        self.undo_stack.clear()
        self.redo_stack.clear()
        
    def get_undo_count(self) -> int:
        """获取可撤销次数"""
        return len(self.undo_stack)
        
    def get_redo_count(self) -> int:
        """获取可重做次数"""
        return len(self.redo_stack)


class EnhancedTextWidget(tk.Text):
    """增强的Text组件，自动跟踪撤销/重做"""
    
    def __init__(self, master=None, **kwargs):
        # 移除undo参数，我们自己管理
        kwargs.pop('undo', None)
        super().__init__(master, **kwargs)
        
        # 创建撤销管理器
        self.undo_manager = UndoRedoManager(self)
        
        # 拦截插入和删除操作
        self._original_insert = super().insert
        self._original_delete = super().delete
        
    def insert(self, index, chars, *args):
        """重写insert方法，记录操作"""
        # 记录插入操作
        self.undo_manager.record_insert(index, chars)
        
        # 执行实际插入
        return self._original_insert(index, chars, *args)
        
    def delete(self, index1, index2=None):
        """重写delete方法，记录操作"""
        # 获取要删除的内容
        if index2 is None:
            index2 = f"{index1}+1c"
        content = self.get(index1, index2)
        
        # 记录删除操作
        self.undo_manager.record_delete(index1, content)
        
        # 执行实际删除
        return self._original_delete(index1, index2)
        
    def undo(self):
        """执行撤销"""
        return self.undo_manager.undo()
        
    def redo(self):
        """执行重做"""
        return self.undo_manager.redo()
        
    def can_undo(self):
        """是否可以撤销"""
        return self.undo_manager.can_undo()
        
    def can_redo(self):
        """是否可以重做"""
        return self.undo_manager.can_redo()


class UndoRedoFeature:
    """撤销/重做功能（用于集成到主应用）"""
    
    def __init__(self, app):
        self.app = app
        self.undo_manager = None
        
    def setup(self, text_widget):
        """
        设置撤销/重做功能
        
        Args:
            text_widget: Text widget实例
        """
        print("[SETUP] 开始设置撤销系统...")
        
        # 禁用Text widget默认的撤销功能，使用我们自己的管理器
        try:
            text_widget.configure(undo=False)
            print("[SETUP] 已禁用原生undo")
        except:
            pass
        
        # 创建撤销管理器（默认已启用）
        self.undo_manager = UndoRedoManager(text_widget)
        
        # 确保启用状态
        if not self.undo_manager.enabled:
            self.undo_manager.enable()
        print(f"[SETUP] 撤销管理器已启用: {self.undo_manager.enabled}")
        
        # 保存widget引用
        self.text_widget = text_widget
        
        # 绑定键盘事件来记录操作
        self._bind_events(text_widget)
        
        # 也尝试包装方法（双保险）
        self._wrap_text_widget(text_widget)
        
        print("[SETUP] 撤销系统设置完成")
        
    def _bind_events(self, text_widget):
        """绑定事件来监听文本变化"""
        print("[SETUP] 绑定键盘事件...")
        
        # 保存上一次的内容用于比较
        self._last_content = text_widget.get("1.0", "end-1c")
        
        def on_key_release(event):
            """键盘释放事件 - 记录变化"""
            if self.undo_manager.is_undoing or self.undo_manager.is_redoing:
                return
                
            try:
                current_content = text_widget.get("1.0", "end-1c")
                
                # 使用简单的diff检测
                if current_content != self._last_content:
                    # 获取光标位置
                    cursor_pos = text_widget.index("insert")
                    
                    # 比较内容
                    if len(current_content) > len(self._last_content):
                        # 插入操作
                        diff = current_content[len(self._last_content):]
                        self.undo_manager.record_insert(cursor_pos, diff)
                        print(f"[EVENT] 记录插入: '{diff}'")
                    elif len(current_content) < len(self._last_content):
                        # 删除操作
                        diff = self._last_content[len(current_content):]
                        self.undo_manager.record_delete(cursor_pos, diff)
                        print(f"[EVENT] 记录删除: '{diff}'")
                    
                    self._last_content = current_content
            except Exception as e:
                print(f"[EVENT] 记录操作失败: {e}")
        
        # 绑定各种可能导致文本变化的事件
        text_widget.bind('<KeyRelease>', on_key_release, add=True)
        text_widget.bind('<<Modified>>', lambda e: self._on_modified(), add=True)
        
        print("[SETUP] 事件绑定完成")
        
    def _on_modified(self):
        """文本修改事件"""
        try:
            if not hasattr(self, 'text_widget'):
                return
                
            if self.undo_manager.is_undoing or self.undo_manager.is_redoing:
                return
                
            # 更新内容快照
            current_content = self.text_widget.get("1.0", "end-1c")
            if hasattr(self, '_last_content'):
                if current_content != self._last_content:
                    self._last_content = current_content
        except:
            pass
        
    def _wrap_text_widget(self, text_widget):
        """包装文本组件的插入和删除方法（备用方案）"""
        print("[SETUP] 包装insert/delete方法...")
        
        # 保存原始方法
        original_insert = text_widget.insert
        original_delete = text_widget.delete
        
        def wrapped_insert(index, chars, *args):
            print(f"[WRAP] insert被调用: {index}, '{chars}'")
            self.undo_manager.record_insert(str(index), chars)
            return original_insert(index, chars, *args)
            
        def wrapped_delete(index1, index2=None):
            print(f"[WRAP] delete被调用: {index1}, {index2}")
            if index2 is None:
                index2 = f"{index1}+1c"
            try:
                content = text_widget.get(index1, index2)
                self.undo_manager.record_delete(str(index1), content)
            except:
                pass
            return original_delete(index1, index2)
            
        # 替换方法
        text_widget.insert = wrapped_insert
        text_widget.delete = wrapped_delete
        
        print("[SETUP] 方法包装完成")
        
        # 绑定粘贴事件以支持批量操作
        self._bind_paste_events(text_widget)
    
    def _bind_paste_events(self, text_widget):
        """绑定粘贴事件以支持批量撤销"""
        print("[SETUP] 绑定粘贴事件...")
        
        def on_paste_start(event):
            """粘贴开始时启用批量模式"""
            if self.undo_manager and not self.undo_manager.in_batch_mode:
                self.undo_manager.begin_batch()
                print("[PASTE] 开始批量模式")
            # 不阻止默认粘贴行为
            return None
        
        def on_paste_end(event=None):
            """粘贴结束时关闭批量模式"""
            if self.undo_manager and self.undo_manager.in_batch_mode:
                # 延迟结束批量模式，确保所有粘贴操作都被记录
                text_widget.after(10, self._end_paste_batch)
        
        # 绑定粘贴相关事件
        text_widget.bind('<<Paste>>', on_paste_start, add='+')
        text_widget.bind('<Control-v>', on_paste_start, add='+')
        text_widget.bind('<Control-V>', on_paste_start, add='+')
        
        # 使用 KeyRelease 来检测粘贴完成
        text_widget.bind('<KeyRelease-v>', on_paste_end, add='+')
        text_widget.bind('<KeyRelease-V>', on_paste_end, add='+')
        
        print("[SETUP] 粘贴事件绑定完成")
    
    def _end_paste_batch(self):
        """结束粘贴批量操作"""
        if self.undo_manager and self.undo_manager.in_batch_mode:
            self.undo_manager.end_batch()
            print("[PASTE] 结束批量模式")
        
    def undo(self):
        """执行撤销"""
        if self.undo_manager:
            success = self.undo_manager.undo()
            if success:
                # 触发文本变化事件
                if hasattr(self.app, 'on_text_change'):
                    self.app.on_text_change(None)
                self.app.update_status(f"↶ 撤销 ({self.undo_manager.get_undo_count()} 可撤销)")
            else:
                self.app.update_status("无法撤销")
            return success
        return False
        
    def redo(self):
        """执行重做"""
        if self.undo_manager:
            success = self.undo_manager.redo()
            if success:
                # 触发文本变化事件
                if hasattr(self.app, 'on_text_change'):
                    self.app.on_text_change(None)
                self.app.update_status(f"↷ 重做 ({self.undo_manager.get_redo_count()} 可重做)")
            else:
                self.app.update_status("无法重做")
            return success
        return False
        
    def can_undo(self):
        """是否可以撤销"""
        return self.undo_manager and self.undo_manager.can_undo()
        
    def can_redo(self):
        """是否可以重做"""
        return self.undo_manager and self.undo_manager.can_redo()
        
    def clear_history(self):
        """清空历史"""
        if self.undo_manager:
            self.undo_manager.clear()
            self.app.update_status("撤销历史已清空")
    
    def begin_batch(self):
        """开始批量操作模式"""
        if self.undo_manager:
            self.undo_manager.begin_batch()
    
    def end_batch(self):
        """结束批量操作模式"""
        if self.undo_manager:
            self.undo_manager.end_batch()
    
    def cancel_batch(self):
        """取消批量操作"""
        if self.undo_manager:
            self.undo_manager.cancel_batch()
    
    @property
    def in_batch_mode(self) -> bool:
        """是否处于批量操作模式"""
        return self.undo_manager.in_batch_mode if self.undo_manager else False
            
    def get_status(self) -> str:
        """获取状态信息"""
        if self.undo_manager:
            undo_count = self.undo_manager.get_undo_count()
            redo_count = self.undo_manager.get_redo_count()
            return f"可撤销: {undo_count} | 可重做: {redo_count}"
        return "撤销/重做未初始化"
