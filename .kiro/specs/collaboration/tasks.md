# Implementation Plan: 实时协作编辑

- [x] 1. 项目依赖和基础设施
  - [x] 1.1 添加协作相关依赖
    - websockets, asyncio
    - _Requirements: 全部_
  - [x] 1.2 创建 ui/collaboration/ 目录结构
    - __init__.py, server.py, client.py, crdt.py, cursor.py, comments.py, history.py, mentions.py
    - _Requirements: 全部_

- [x] 2. CRDT 引擎模块
  - [x] 2.1 实现 CRDTOperation 数据类
    - id, type, position, content, author, timestamp, vector_clock
    - _Requirements: 2.3_
  - [x] 2.2 实现 CRDTEngine 类
    - local_insert(), local_delete(), apply_remote()
    - _Requirements: 2.1, 2.2, 2.3_
  - [x]* 2.3 编写 CRDT 属性测试
    - **Property 1: CRDT 收敛性**
    - **Property 2: 操作顺序保持**
    - **Validates: Requirements 2.3, 2.5**
  - [x] 2.4 实现状态序列化
    - get_state(), load_state()
    - _Requirements: 1.5_

- [x] 3. WebSocket 服务器模块
  - [x] 3.1 实现 CollaborationServer 类
    - start(), stop(), create_session()
    - _Requirements: 1.1, 1.3_
  - [x] 3.2 实现连接处理
    - handle_connection(), broadcast()
    - _Requirements: 1.2, 1.4, 1.5_
  - [x]* 3.3 编写服务器属性测试
    - **Property 6: 会话隔离性**
    - **Validates: Requirements 1.1**
  - [x] 3.4 实现会话管理
    - 会话码生成, 密码验证, 参与者管理
    - _Requirements: 1.7, 7.1, 7.2_

- [x] 4. WebSocket 客户端模块
  - [x] 4.1 实现 CollaborationClient 类
    - connect(), disconnect(), send_operation()
    - _Requirements: 1.2, 1.4, 2.1_
  - [x] 4.2 实现消息处理
    - _receive_loop(), 回调注册
    - _Requirements: 2.2_
  - [x] 4.3 实现光标同步
    - send_cursor_update(), on_cursor_update()
    - _Requirements: 3.2_

- [x] 5. Checkpoint - 确保基础同步功能正常
  - All tests pass.

- [x] 6. 光标管理模块
  - [x] 6.1 实现 CursorManager 类
    - add_cursor(), remove_cursor(), update_cursor()
    - _Requirements: 3.1, 3.2, 3.3_
  - [x]* 6.2 编写光标属性测试
    - **Property 3: 光标位置一致性**
    - **Validates: Requirements 2.4**
  - [x] 6.3 实现光标渲染
    - render_cursors(), _draw_cursor(), _draw_selection()
    - _Requirements: 3.4, 3.5_
  - [x] 6.4 实现在线状态显示
    - 用户列表, 状态指示器
    - _Requirements: 3.5, 3.6_

- [x] 7. 评论管理模块
  - [x] 7.1 实现 CommentManager 类
    - create_thread(), add_reply(), resolve_thread()
    - _Requirements: 4.1, 4.4, 4.5_
  - [x]* 7.2 编写评论属性测试
    - **Property 4: 评论范围有效性**
    - **Validates: Requirements 4.2**
  - [x] 7.3 实现评论显示
    - 侧边栏面板, 文档内指示器
    - _Requirements: 4.3, 4.6, 4.7_
  - [x] 7.4 实现评论导出
    - export_comments()
    - _Requirements: 4.8_

- [x] 8. 历史管理模块
  - [x] 8.1 实现 HistoryManager 类
    - record(), get_history(), restore()
    - _Requirements: 6.1, 6.2, 6.5_
  - [x]* 8.2 编写历史属性测试
    - **Property 5: 历史可恢复性**
    - **Validates: Requirements 6.5**
  - [x] 8.3 实现版本对比
    - diff(), 高亮显示差异
    - _Requirements: 6.4_
  - [x] 8.4 实现历史压缩和持久化
    - compress_old_entries(), export_history(), import_history()
    - _Requirements: 6.7, 6.8_

- [x] 9. @提及和任务模块
  - [x] 9.1 实现 MentionManager 类
    - parse_mentions(), parse_tasks()
    - _Requirements: 5.1, 5.3_
  - [x] 9.2 实现提及通知
    - notify_mention(), 视觉/声音提醒
    - _Requirements: 5.2, 5.7_
  - [x] 9.3 实现任务管理
    - complete_task(), get_tasks_for_user()
    - _Requirements: 5.4, 5.5, 5.6_
  - [x] 9.4 实现@自动完成
    - get_user_suggestions()
    - _Requirements: 5.1_

- [x] 10. Checkpoint - 确保协作功能模块正常
  - All tests pass.

- [x] 11. 用户界面
  - [x] 11.1 创建协作控制面板
    - 开始协作, 加入会话, 会话码显示
    - _Requirements: 1.1, 1.2_
  - [x] 11.2 创建参与者列表面板
    - 在线用户, 颜色标识, 权限显示
    - _Requirements: 3.5, 7.3_
  - [x] 11.3 创建评论侧边栏
    - 评论线程列表, 回复输入, 解决按钮
    - _Requirements: 4.3_
  - [x] 11.4 创建历史时间线视图
    - 修改记录, 版本对比, 恢复按钮
    - _Requirements: 6.2, 6.3_
  - [x] 11.5 创建任务面板
    - 任务列表, 分配筛选, 完成状态
    - _Requirements: 5.5_

- [x] 12. 权限和安全
  - [x] 12.1 实现密码保护
    - 会话密码设置和验证
    - _Requirements: 7.1, 7.2_
  - [x] 12.2 实现权限控制
    - 查看/编辑/评论权限
    - _Requirements: 7.3, 7.4_
  - [x] 12.3 实现踢出功能
    - 主持人踢出参与者
    - _Requirements: 7.5_

- [-] 13. 集成到主应用



  - [ ] 13.1 在工具栏添加协作按钮
    - 开始/加入协作
    - _Requirements: 1.1_
  - [ ] 13.2 集成远程光标到编辑器
    - 光标渲染, 选区高亮
    - _Requirements: 3.3, 3.4_
  - [ ] 13.3 集成评论指示器到编辑器
    - 边栏标记, 点击跳转
    - _Requirements: 4.6, 4.7_

- [x] 14. Final Checkpoint - 确保协作功能完整可用
  - All tests pass (76 passed, 1 skipped).
