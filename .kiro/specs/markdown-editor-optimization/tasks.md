# Implementation Plan

## Phase 1: 关键Bug修复

- [x] 1. 修复撤销/重做系统






  - [x] 1.1 修复UndoRedoManager默认启用状态




    - 将`enabled`默认值从`False`改为`True`


    - 确保`setup()`方法正确初始化




    - _Requirements: 6.1_
  - [x] 1.2 编写撤销/重做属性测试




    - **Property 1: 撤销/重做一致性**
    - **Validates: Requirements 2.1, 6.1**
  - [x] 1.3 添加批量操作支持
    - 实现`begin_batch()`和`end_batch()`方法
    - 粘贴多行时合并为单个撤销点
    - _Requirements: 2.1_

- [x] 2. 修复快捷键冲突
  - [x] 2.1 解决Ctrl+Shift+C冲突
    - 将"协作"功能改为`Ctrl+Alt+C`
    - 保留`Ctrl+Shift+C`用于复制
    - _Requirements: 8.3_
  - [x] 2.2 编写快捷键唯一性属性测试
    - **Property 9: 快捷键唯一性**
    - **Validates: Requirements 8.1, 8.3**

- [x] 3. Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

## Phase 2: 性能优化

- [ ] 4. 实现虚拟化语法高亮
  - [ ] 4.1 修改SyntaxHighlighter只渲染可视区域
    - 添加`buffer_lines`参数控制缓冲区大小
    - 实现`_highlight_visible()`方法优化
    - _Requirements: 1.4_
  - [ ] 4.2 编写增量渲染属性测试
    - **Property 3: 增量渲染正确性**
    - **Validates: Requirements 1.3, 3.4**

- [ ] 5. 优化预览渲染性能
  - [ ] 5.1 实现增量预览更新
    - 添加块级元素缓存
    - 只重新渲染变化的块
    - _Requirements: 1.3, 3.4_
  - [ ] 5.2 添加渲染超时处理
    - 超过500ms显示加载指示器
    - 允许用户取消长时间渲染
    - _Requirements: 1.5_

- [ ] 6. Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

## Phase 3: 滚动同步精确化



- [-] 7. 实现精确滚动同步

  - [ ] 7.1 创建PreciseScrollSync类
    - 构建源码行到预览位置的映射表
    - 实现双向同步方法
    - _Requirements: 3.1_
  - [ ] 7.2 编写滚动同步属性测试
    - **Property 2: 滚动同步精确性**
    - **Validates: Requirements 3.1**
  - [ ] 7.3 集成到PreviewSyncFeature
    - 替换现有的简单比例同步
    - 添加平滑滚动动画
    - _Requirements: 3.1, 3.2_

- [ ] 8. Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

## Phase 4: UI/UX改进

- [ ] 9. 优化工具栏布局
  - [ ] 9.1 重新组织工具栏按钮
    - 将18个按钮分组为：文件、编辑、插入、视图、工具
    - 使用下拉菜单减少可见按钮数量
    - _Requirements: 5.1_
  - [ ] 9.2 添加工具栏自定义功能
    - 允许用户隐藏不常用按钮
    - 保存用户偏好到配置
    - _Requirements: 5.4_

- [ ] 10. 实现主题过渡动画
  - [ ] 10.1 创建ThemeTransition类
    - 实现颜色插值算法
    - 300ms平滑过渡
    - _Requirements: 5.3_
  - [ ] 10.2 编写主题颜色对比度属性测试
    - **Property 8: 主题颜色对比度**
    - **Validates: Requirements 5.5**

- [ ] 11. Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

## Phase 5: 编辑器增强

- [ ] 12. 实现智能格式化快捷键
  - [ ] 12.1 添加Ctrl+B/I/U格式化支持
    - 选中文本时添加/移除格式标记
    - 无选中时在光标处插入标记对
    - _Requirements: 2.1_
  - [ ] 12.2 编写格式化操作属性测试
    - **Property 4: 格式化操作可逆性**
    - **Validates: Requirements 2.1**

- [ ] 13. 改进搜索替换功能
  - [ ] 13.1 添加实时高亮和计数显示
    - 输入时实时高亮所有匹配
    - 显示"当前/总数"格式
    - _Requirements: 7.2, 7.4_
  - [ ] 13.2 编写搜索结果属性测试
    - **Property 6: 搜索结果完整性**
    - **Validates: Requirements 7.2, 7.4**

- [ ] 14. Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

## Phase 6: 稳定性增强

- [ ] 15. 改进自动保存和恢复
  - [ ] 15.1 优化自动保存机制
    - 使用增量保存减少IO
    - 添加崩溃恢复提示
    - _Requirements: 6.1, 6.5_
  - [ ] 15.2 编写自动保存属性测试
    - **Property 10: 自动保存完整性**
    - **Validates: Requirements 6.1, 6.5**

- [ ] 16. 改进标签页状态管理
  - [ ] 16.1 修复标签页修改状态显示
    - 确保修改标记与实际状态一致
    - 添加右键菜单功能
    - _Requirements: 9.1, 9.2, 9.5_
  - [ ] 16.2 编写标签页状态属性测试
    - **Property 7: 标签页状态一致性**
    - **Validates: Requirements 9.5**

- [ ] 17. Final Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.
