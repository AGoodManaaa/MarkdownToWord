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
        
        # 是否启用（可以临时禁用）
        self.enabled = False
        
        # 操作记录缓冲
        self.last_insert_pos = None
        self.last_delete_pos = None
        
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
        # 添加新操作时，清空重做栈
        self.redo_stack.clear()
        
        # 添加到撤销栈
        self.undo_stack.append(operation)
        
        # 限制栈大小
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
            
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
            
            try:
                # 执行反向操作
                if operation.op_type == 'insert':
                    # 撤销插入 = 删除
                    end_pos = f"{operation.position} + {len(operation.content)} chars"
                    self.text_widget.delete(operation.position, end_pos)
                else:  # delete
                    # 撤销删除 = 插入
                    self.text_widget.insert(operation.position, operation.content)
                    
                # 添加到重做栈
                self.redo_stack.append(operation)
                
                # 移动光标到操作位置
                try:
                    self.text_widget.mark_set('insert', operation.position)
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
            
            try:
                # 重新执行操作
                if operation.op_type == 'insert':
                    self.text_widget.insert(operation.position, operation.content)
                else:  # delete
                    end_pos = f"{operation.position} + {len(operation.content)} chars"
                    self.text_widget.delete(operation.position, end_pos)
                    
                # 添加回撤销栈
                self.undo_stack.append(operation)
                
                # 移动光标到操作位置
                try:
                    self.text_widget.mark_set('insert', operation.position)
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
        
        # 禁用Text widget默认的撤销功能
        try:
            text_widget.configure(undo=False)
            print("[SETUP] 已禁用原生undo")
        except:
            pass
        
        # 创建撤销管理器
        self.undo_manager = UndoRedoManager(text_widget)
        
        # 启用撤销记录
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
            
    def get_status(self) -> str:
        """获取状态信息"""
        if self.undo_manager:
            undo_count = self.undo_manager.get_undo_count()
            redo_count = self.undo_manager.get_redo_count()
            return f"可撤销: {undo_count} | 可重做: {redo_count}"
        return "撤销/重做未初始化"
