# Implementation Plan: OCR 图片转 Markdown

- [x] 1. 项目依赖和基础设施
  - [x] 1.1 添加 OCR 相关依赖到 requirements.txt
    - paddlepaddle, paddleocr, pix2tex, pillow
    - _Requirements: 全部_
  - [x] 1.2 创建 ui/features/ocr/ 目录结构
    - __init__.py, image_input.py, ocr_engine.py, markdown_gen.py, dialog.py
    - _Requirements: 全部_

- [x] 2. 图片输入模块
  - [x] 2.1 实现 ImageInputManager 类
    - load_from_file(), load_from_clipboard(), validate_image()
    - _Requirements: 1.1, 1.2, 1.3, 1.5_
  - [x]* 2.2 编写 ImageInputManager 属性测试
    - **Property 1: 图片格式验证**
    - **Validates: Requirements 1.5**
  - [x] 2.3 实现截图功能
    - start_screenshot(), 区域选择, 截图捕获
    - _Requirements: 1.4_
  - [x] 2.4 实现批量加载
    - load_multiple_files()
    - _Requirements: 5.1_

- [x] 3. OCR 识别引擎
  - [x] 3.1 实现 OCREngine 基础类
    - initialize(), recognize(), 延迟加载机制
    - _Requirements: 2.1_
  - [x] 3.2 实现文字识别功能
    - recognize_text(), 中英文支持, 置信度计算
    - _Requirements: 2.2, 2.3, 2.4, 2.5_
  - [x]* 3.3 编写 OCREngine 属性测试
    - **Property 2: OCR 结果完整性**
    - **Validates: Requirements 2.1, 2.4**
  - [x] 3.4 实现表格识别功能
    - recognize_table(), 表格结构提取
    - _Requirements: 3.1, 3.2, 3.4, 3.5_
  - [x] 3.5 实现公式识别功能
    - recognize_formula(), LaTeX 转换
    - _Requirements: 4.1, 4.2_

- [x] 4. Markdown 生成器
  - [x] 4.1 实现 MarkdownGenerator 类
    - generate(), text_to_markdown()
    - _Requirements: 2.4, 2.5_
  - [x] 4.2 实现表格转 Markdown
    - table_to_markdown(), GitHub 风格表格
    - _Requirements: 3.3_
  - [x]* 4.3 编写表格生成属性测试
    - **Property 3: Markdown 表格语法正确性**
    - **Validates: Requirements 3.3**
  - [x] 4.4 实现公式转 Markdown
    - formula_to_markdown(), 行内/块级公式
    - _Requirements: 4.3, 4.4, 4.5_
  - [x]* 4.5 编写公式生成属性测试
    - **Property 4: LaTeX 公式语法正确性**
    - **Validates: Requirements 4.2, 4.3**
  - [x] 4.6 实现列表和标题检测
    - _detect_list_structure(), _detect_heading_structure()
    - _Requirements: 2.5_

- [x] 5. Checkpoint - 确保核心 OCR 功能正常
  - All tests pass.

- [x] 6. 批量处理模块
  - [x] 6.1 实现 BatchOCRProcessor 类
    - process_batch(), cancel(), 进度回调
    - _Requirements: 5.1, 5.2, 5.3_
  - [x]* 6.2 编写批量处理属性测试
    - **Property 5: 批量处理顺序保持**
    - **Property 6: 批量处理容错性**
    - **Validates: Requirements 5.4, 5.5**
  - [x] 6.3 实现结果合并
    - 多图片结果合并为单一文档
    - _Requirements: 5.3_

- [x] 7. OCR 对话框界面
  - [x] 7.1 创建 OCRDialog 主界面
    - 图片预览区, 结果编辑区, 操作按钮
    - _Requirements: 6.1_
  - [x] 7.2 实现图片导入交互
    - 文件选择, 剪贴板粘贴, 截图按钮
    - _Requirements: 1.1, 1.2, 1.3, 1.4_
  - [x] 7.3 实现识别结果预览和编辑
    - 实时预览, 可编辑文本框
    - _Requirements: 6.1, 6.2_
  - [x] 7.4 实现结果导出功能
    - 插入到文档, 复制到剪贴板, 保存文件
    - _Requirements: 6.3, 6.4, 6.5_

- [x] 8. 集成到主应用





  - [x] 8.1 在 gui.py 中添加 OCR 功能入口

    - 工具栏按钮, 菜单项
    - _Requirements: 1.1_
  - [ ] 8.2 添加快捷键支持
    - Ctrl+Shift+O 打开 OCR 对话框
    - _Requirements: 1.1_

- [x] 9. Final Checkpoint - 确保 OCR 功能完整可用
  - All tests pass (76 passed, 1 skipped).
