# Implementation Plan

## 1. PDF 导出功能

- [x] 1.1 创建 PDFExportFeature 类
  - 在 `ui/features/pdf_export.py` 创建基础类结构
  - 实现 `__init__` 方法，接收 app 引用
  - _Requirements: 1.1, 1.2_

- [x] 1.2 实现 Word 转 PDF 核心逻辑
  - 使用 `win32com.client` 打开 Word 应用
  - 实现 `_convert_docx_to_pdf` 方法
  - 处理 Word 未安装的情况
  - _Requirements: 1.2, 1.5_

- [x] 1.3 实现导出对话框和流程
  - 实现 `_show_export_dialog` 方法显示保存对话框
  - 实现 `export_to_pdf` 主方法整合完整流程
  - 添加成功/失败通知
  - _Requirements: 1.1, 1.3, 1.4_

- [x] 1.4 集成到主应用
  - 在 `gui.py` 中初始化 PDFExportFeature
  - 在工具栏添加 PDF 导出按钮
  - 添加快捷键绑定 (Ctrl+Shift+P)
  - 更新 `ui/features/__init__.py` 导出
  - _Requirements: 1.1, 1.6_

- [ ]* 1.5 编写 PDF 导出属性测试
  - **Property 1: PDF export produces valid file**
  - **Property 2: PDF export preserves page size**
  - **Validates: Requirements 1.2, 1.6**

## 2. 预览区缩放功能

- [x] 2.1 创建 PreviewZoomFeature 类
  - 在 `ui/features/preview_zoom.py` 创建基础类结构
  - 定义缩放范围常量 (0.5-1.5)
  - 实现配置保存/恢复方法
  - 注意：`MarkdownPreview` 已有 `set_scale()` 方法可复用
  - _Requirements: 2.6_

- [x] 2.2 实现缩放控件 UI
  - 实现 `create_controls` 方法创建缩放按钮组
  - 添加放大、缩小、重置按钮
  - 添加当前缩放比例显示标签
  - _Requirements: 2.1_

- [x] 2.3 实现缩放逻辑
  - 实现 `zoom_in` 方法（+10%，最大 150%）
  - 实现 `zoom_out` 方法（-10%，最小 50%）
  - 实现 `reset_zoom` 方法（重置为 100%）
  - 调用 `MarkdownPreview.set_scale()` 应用缩放
  - _Requirements: 2.2, 2.3, 2.4, 2.5_

- [ ]* 2.4 编写缩放属性测试
  - **Property 3: Zoom in increases scale correctly**
  - **Property 4: Zoom out decreases scale correctly**
  - **Property 5: Zoom scale persistence round-trip**
  - **Validates: Requirements 2.2, 2.3, 2.6**

- [x] 2.5 集成到主应用
  - 在 `gui.py` 中初始化 PreviewZoomFeature
  - 在预览面板顶部添加缩放控件
  - 应用启动时恢复上次缩放比例
  - 更新 `ui/features/__init__.py` 导出
  - _Requirements: 2.1, 2.6_

## 3. 多标签页编辑功能

- [x] 3.1 创建 TabData 数据类和 TabManagerFeature 类
  - 在 `ui/features/tab_manager.py` 创建 TabData 数据类
  - 创建 TabManagerFeature 基础类结构
  - 实现标签页列表管理方法
  - _Requirements: 3.1_

- [x] 3.2 实现标签栏 UI
  - 实现 `create_tab_bar` 方法创建标签栏框架
  - 实现 `_create_tab_button` 方法创建单个标签按钮
  - 添加新建标签按钮 (+)
  - 实现标签关闭按钮 (×)
  - _Requirements: 3.1, 3.4_

- [x] 3.3 实现标签页核心操作
  - 实现 `new_tab` 方法创建新标签页
  - 实现 `close_tab` 方法关闭标签页
  - 实现 `switch_tab` 方法切换标签页
  - 实现内容保存/恢复逻辑
  - _Requirements: 3.2, 3.3, 3.4_

- [x] 3.4 实现未保存状态管理
  - 实现 `update_tab_title` 方法更新标签标题
  - 添加修改状态标记 (*)
  - 实现关闭前未保存检查对话框
  - _Requirements: 3.5, 3.6_

- [x] 3.5 实现标签页右键菜单
  - 实现 `_show_context_menu` 方法
  - 添加"关闭"、"关闭其他"、"关闭全部"选项
  - _Requirements: 3.9_

- [x] 3.6 实现空标签页自动创建
  - 当所有标签关闭时自动创建新空白标签
  - _Requirements: 3.7_

- [ ]* 3.7 编写标签页属性测试
  - **Property 6: Tab creation for opened files**
  - **Property 7: Tab switch displays correct content**
  - **Property 8: Modified indicator accuracy**
  - **Validates: Requirements 3.2, 3.3, 3.5**

- [x] 3.8 集成到主应用
  - 在 `gui.py` 中初始化 TabManagerFeature
  - 修改 `_create_main_content` 添加标签栏
  - 修改 `FileOpsFeature` 与标签页集成
  - 更新 `ui/features/__init__.py` 导出
  - _Requirements: 3.1, 3.2_

## 4. 字数统计详情功能

- [x] 4.1 创建 DocumentStatistics 数据类和 StatisticsDetailFeature 类
  - 在 `ui/features/statistics_detail.py` 创建数据类
  - 创建 StatisticsDetailFeature 基础类结构
  - 注意：现有 `StatusBarFeature` 有基础字数统计可参考
  - _Requirements: 4.1_

- [x] 4.2 实现统计计算逻辑
  - 实现 `_count_chinese_chars` 方法（使用 Unicode 范围判断）
  - 实现 `_count_english_words` 方法（使用正则分词）
  - 实现 `_calculate_reading_time` 方法（中文 300 字/分，英文 200 词/分）
  - 实现 `calculate_statistics` 主方法
  - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.8_

- [ ]* 4.3 编写统计计算属性测试
  - **Property 9: Character count accuracy**
  - **Property 10: Word count accuracy**
  - **Property 11: Paragraph count accuracy**
  - **Property 12: Reading time calculation**
  - **Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.8**

- [x] 4.4 实现状态栏更新
  - 实现 `update_status_bar` 方法
  - 修改现有 StatusBarFeature 显示更多信息（阅读时间）
  - 添加点击事件绑定
  - _Requirements: 4.1, 4.2, 4.3, 4.4_

- [x] 4.5 实现详细统计弹窗
  - 实现 `show_detail_popup` 方法
  - 创建弹窗 UI 显示所有统计项
  - 添加复制统计信息功能
  - _Requirements: 4.5, 4.6_

- [x] 4.6 集成到主应用
  - 在 `gui.py` 中初始化 StatisticsDetailFeature
  - 修改 `preview_sync.py` 在内容变化时更新统计
  - 更新 `ui/features/__init__.py` 导出
  - _Requirements: 4.7_

## 5. 最终集成和测试

- [x] 5.1 更新配置系统
  - 在 `ui/theme.py` 的 DEFAULT_CONFIG 添加新配置项
  - 添加 `preview_zoom_scale` 配置
  - 添加 `open_tabs` 和 `active_tab_id` 配置
  - _Requirements: 2.6, 3.1_

- [ ] 5.2 Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.

- [ ]* 5.3 编写集成测试
  - 测试 PDF 导出完整流程
  - 测试多标签页与文件操作的交互
  - 测试应用重启后状态恢复

- [ ] 5.4 Final Checkpoint - 确保所有测试通过
  - Ensure all tests pass, ask the user if questions arise.
