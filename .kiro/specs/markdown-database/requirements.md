# Requirements Document

## Introduction

本文档定义了 Markdown 数据库功能的需求。该功能将应用转变为一个知识管理系统，支持本地文档全文搜索、标签分类、双向链接和知识图谱可视化。

## Glossary

- **双向链接**: 文档 A 链接到文档 B 时，文档 B 自动显示来自文档 A 的反向链接
- **知识图谱**: 以节点和边的形式可视化文档之间的关系
- **全文搜索**: 在所有文档内容中搜索关键词
- **标签 (Tag)**: 用于分类和组织文档的元数据标记
- **Vault**: 文档库，包含所有 Markdown 文件的根目录
- **Backlink**: 反向链接，指向当前文档的其他文档链接

## Requirements

### Requirement 1: 文档库管理

**User Story:** 作为用户，我希望能够创建和管理文档库，集中存储和组织我的 Markdown 文件。

#### Acceptance Criteria

1. WHEN the user opens the app THEN the system SHALL allow selecting or creating a vault directory
2. WHEN a vault is selected THEN the system SHALL scan and index all Markdown files
3. WHEN new files are added to vault THEN the system SHALL automatically detect and index them
4. WHEN files are modified externally THEN the system SHALL update the index accordingly
5. WHEN the user switches vaults THEN the system SHALL load the new vault's index
6. IF the vault directory is inaccessible THEN the system SHALL display an error and offer alternatives

### Requirement 2: 全文搜索

**User Story:** 作为用户，我希望能够快速搜索所有文档的内容，找到我需要的信息。

#### Acceptance Criteria

1. WHEN the user types in the search box THEN the system SHALL search across all indexed documents
2. WHEN searching THEN the system SHALL support full-text content search
3. WHEN searching THEN the system SHALL support filename search
4. WHEN displaying results THEN the system SHALL show matching context snippets
5. WHEN displaying results THEN the system SHALL highlight matched keywords
6. WHEN the user clicks a result THEN the system SHALL open the document and scroll to match
7. WHEN searching THEN the system SHALL support advanced operators (AND, OR, NOT, quotes)
8. WHEN the search index is outdated THEN the system SHALL rebuild it automatically

### Requirement 3: 标签和分类

**User Story:** 作为用户，我希望能够使用标签对文档进行分类，方便组织和查找。

#### Acceptance Criteria

1. WHEN a document contains #tag syntax THEN the system SHALL extract and index the tags
2. WHEN the user views the tag panel THEN the system SHALL display all tags with document counts
3. WHEN the user clicks a tag THEN the system SHALL filter documents by that tag
4. WHEN the user adds a tag to a document THEN the system SHALL update the index immediately
5. WHEN displaying tags THEN the system SHALL support hierarchical tags (e.g., #project/work)
6. WHEN the user renames a tag THEN the system SHALL update all documents containing that tag
7. IF a tag has no documents THEN the system SHALL remove it from the tag list

### Requirement 4: 双向链接

**User Story:** 作为用户，我希望能够在文档之间创建链接，并自动看到哪些文档链接到当前文档。

#### Acceptance Criteria

1. WHEN the user types [[filename]] THEN the system SHALL create a link to that document
2. WHEN typing [[ THEN the system SHALL show autocomplete suggestions for existing documents
3. WHEN a document is linked THEN the system SHALL track the link relationship
4. WHEN viewing a document THEN the system SHALL display backlinks (documents linking to it)
5. WHEN clicking a wiki-link THEN the system SHALL navigate to the linked document
6. WHEN the linked document does not exist THEN the system SHALL offer to create it
7. WHEN a document is renamed THEN the system SHALL update all links pointing to it
8. WHEN a document is deleted THEN the system SHALL mark links to it as broken

### Requirement 5: 知识图谱可视化

**User Story:** 作为用户，我希望能够以图形方式查看文档之间的关系，发现知识结构。

#### Acceptance Criteria

1. WHEN the user opens the graph view THEN the system SHALL display documents as nodes
2. WHEN documents are linked THEN the system SHALL display edges between nodes
3. WHEN viewing the graph THEN the system SHALL allow zooming and panning
4. WHEN the user clicks a node THEN the system SHALL highlight connected nodes
5. WHEN the user double-clicks a node THEN the system SHALL open that document
6. WHEN displaying the graph THEN the system SHALL use force-directed layout for clarity
7. WHEN the user hovers over a node THEN the system SHALL show document preview
8. WHEN filtering by tag THEN the system SHALL update the graph to show only matching documents

### Requirement 6: 文档元数据

**User Story:** 作为用户，我希望能够为文档添加元数据（如创建日期、作者、状态），便于管理。

#### Acceptance Criteria

1. WHEN a document has YAML frontmatter THEN the system SHALL parse and index metadata
2. WHEN viewing document list THEN the system SHALL display sortable metadata columns
3. WHEN filtering documents THEN the system SHALL support metadata-based filters
4. WHEN the user edits frontmatter THEN the system SHALL validate YAML syntax
5. WHEN creating a new document THEN the system SHALL offer to add default frontmatter
6. IF frontmatter is invalid THEN the system SHALL display a warning but continue loading
