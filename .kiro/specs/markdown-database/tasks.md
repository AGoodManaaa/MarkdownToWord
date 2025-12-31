# Implementation Plan: Markdown 数据库

- [x] 1. 项目依赖和基础设施
  - [x] 1.1 添加数据库相关依赖
    - watchdog, whoosh/sqlite-fts5, pyyaml, networkx
    - _Requirements: 全部_
  - [x] 1.2 创建 ui/database/ 目录结构
    - __init__.py, vault.py, search.py, tags.py, links.py, graph.py, metadata.py
    - _Requirements: 全部_
  - [x] 1.3 设计并创建 SQLite 数据库 schema
    - documents, tags, document_tags, links, documents_fts
    - _Requirements: 1.2, 2.1_

- [x] 2. 文档库管理模块
  - [x] 2.1 实现 VaultManager 类
    - open_vault(), create_vault(), scan_vault()
    - _Requirements: 1.1, 1.2_
  - [x] 2.2 实现文件监控功能
    - watch_changes(), stop_watching(), 使用 watchdog
    - _Requirements: 1.3, 1.4_
  - [x]* 2.3 编写 VaultManager 属性测试
    - **Property 1: 索引一致性**
    - **Validates: Requirements 1.3, 1.4**
  - [x] 2.4 实现文档信息获取
    - get_document(), get_all_documents(), get_recent_documents()
    - _Requirements: 1.5_

- [x] 3. 全文搜索模块
  - [x] 3.1 实现 SearchEngine 类
    - build_index(), update_index(), remove_from_index()
    - _Requirements: 2.1, 2.8_
  - [x] 3.2 实现搜索功能
    - search(), search_by_tag(), search_by_filename()
    - _Requirements: 2.2, 2.3_
  - [x]* 3.3 编写 SearchEngine 属性测试
    - **Property 2: 搜索结果相关性**
    - **Validates: Requirements 2.1, 2.2**
  - [x] 3.4 实现高级查询解析
    - parse_query(), AND/OR/NOT/引号支持
    - _Requirements: 2.7_
  - [x] 3.5 实现搜索结果高亮
    - 匹配上下文, 关键词高亮
    - _Requirements: 2.4, 2.5_

- [x] 4. Checkpoint - 确保搜索功能正常
  - All tests pass.

- [x] 5. 标签管理模块
  - [x] 5.1 实现 TagManager 类
    - extract_tags(), get_all_tags(), get_documents_by_tag()
    - _Requirements: 3.1, 3.2, 3.3_
  - [x]* 5.2 编写 TagManager 属性测试
    - **Property 3: 标签提取正确性**
    - **Validates: Requirements 3.1**
  - [x] 5.3 实现标签操作功能
    - rename_tag(), delete_tag(), merge_tags()
    - _Requirements: 3.6, 3.7_
  - [x] 5.4 实现层级标签支持
    - get_tag_hierarchy(), 支持 #project/work 格式
    - _Requirements: 3.5_

- [x] 6. 双向链接模块
  - [x] 6.1 实现 LinkManager 类
    - extract_links(), resolve_link()
    - _Requirements: 4.1, 4.5_
  - [x] 6.2 实现链接查询功能
    - get_outgoing_links(), get_backlinks()
    - _Requirements: 4.3, 4.4_
  - [x]* 6.3 编写 LinkManager 属性测试
    - **Property 4: 双向链接对称性**
    - **Property 5: 链接更新传播**
    - **Validates: Requirements 4.3, 4.4, 4.7**
  - [x] 6.4 实现链接维护功能
    - update_links_on_rename(), find_broken_links()
    - _Requirements: 4.7, 4.8_
  - [x] 6.5 实现链接自动完成
    - get_link_suggestions(), [[ 触发
    - _Requirements: 4.2_

- [x] 7. 知识图谱模块
  - [x] 7.1 实现 GraphView 类
    - build_graph(), calculate_layout()
    - _Requirements: 5.1, 5.2, 5.6_
  - [x]* 7.2 编写 GraphView 属性测试
    - **Property 6: 图数据完整性**
    - **Validates: Requirements 5.1, 5.2**
  - [x] 7.3 实现图过滤功能
    - filter_by_tag(), filter_by_depth()
    - _Requirements: 5.8_
  - [x] 7.4 实现图交互功能
    - get_neighbors(), 节点点击/悬停
    - _Requirements: 5.3, 5.4, 5.5, 5.7_

- [x] 8. 元数据解析模块
  - [x] 8.1 实现 MetadataParser 类
    - parse(), extract(), update(), validate()
    - _Requirements: 6.1, 6.4_
  - [x] 8.2 实现默认 frontmatter 生成
    - create_default()
    - _Requirements: 6.5_

- [x] 9. Checkpoint - 确保核心数据库功能正常
  - All tests pass.

- [x] 10. 用户界面
  - [x] 10.1 创建文档库选择界面
    - 打开/创建 vault 对话框
    - _Requirements: 1.1_
  - [x] 10.2 创建搜索面板
    - 搜索框, 结果列表, 高亮预览
    - _Requirements: 2.1, 2.4, 2.5, 2.6_
  - [x] 10.3 创建标签面板
    - 标签列表, 文档计数, 点击过滤
    - _Requirements: 3.2, 3.3_
  - [x] 10.4 创建反向链接面板
    - backlinks 列表, 点击跳转
    - _Requirements: 4.4_
  - [x] 10.5 创建知识图谱视图
    - Canvas 绑定, 节点渲染, 交互处理
    - _Requirements: 5.1-5.8_

- [-] 11. 集成到主应用


  - [ ] 11.1 在侧边栏添加数据库功能入口
    - 文档库, 搜索, 标签, 图谱 标签页
    - _Requirements: 全部_
  - [ ] 11.2 实现 [[ 链接自动完成
    - 编辑器中输入 [[ 触发建议

    - _Requirements: 4.2_
  - [ ] 11.3 添加快捷键支持
    - Ctrl+Shift+F 全局搜索, Ctrl+G 图谱视图
    - _Requirements: 2.1, 5.1_

- [x] 12. Final Checkpoint - 确保数据库功能完整可用
  - All tests pass (76 passed, 1 skipped).
