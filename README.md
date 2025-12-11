<p align="center">
  <img src="https://img.shields.io/badge/✨-Markdown_to_Word-6366F1?style=for-the-badge&labelColor=1E293B" alt="Logo"/>
</p>

<h1 align="center">📝 MarkdowntoWord</h1>

<p align="center">
  <b>把你从GPT上复制下来的 Markdown 变成漂亮的 Word 文档，就是这么简单！</b>
</p>

<p align="center">
  <a href="#-功能特色">功能</a> •
  <a href="#-快速开始">安装</a> •
  <a href="#-使用方法">使用</a> •
  <a href="#-技术栈">技术栈</a> •
  <a href="#-贡献">贡献</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python"/>
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License"/>
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?style=flat-square&logo=windows&logoColor=white" alt="Platform"/>
  <img src="https://img.shields.io/github/stars/AGoodManaaa/MarkdowntoWord?style=flat-square&color=yellow" alt="Stars"/>
</p>

---

## 🎯 这是什么？

还在为 Markdown 转 Word 发愁？手动复制粘贴格式全乱？

**MarkdowntoWord** 来拯救你！一键转换，格式完美保留，公式、表格、代码块统统搞定！

> 💡 专为「懒人」打造 —— 因为我们相信，生活已经够累了，转换文档不应该成为负担。

---

## ✨ 功能特色

### 🎨 实时预览
左边写 Markdown，右边立刻看效果，所见即所得！

### 📐 完美公式支持
LaTeX 数学公式？小 case！行内公式 `$E=mc^2$`、块级公式 `$$\int_0^\infty$$` 都能漂亮渲染。

### 📊 表格随便搞
支持复杂表格，表头加粗、单元格合并，导出到 Word 还是那么整齐！

### 💻 代码高亮
代码块保留格式，语法高亮，让你的代码在 Word 里也能闪闪发光。

### 🖼️ 图片自动处理
本地图片、网络图片统统支持，自动下载嵌入，再也不怕图片丢失！

### 📑 多级列表
有序列表、无序列表、多级嵌套，编号格式自动切换：
- 第一级：1. 2. 3.
- 第二级：a) b) c)
- 第三级：i. ii. iii.

### 🎭 富文本样式
**加粗**、*斜体*、~~删除线~~、`行内代码`、[链接](https://github.com)、上标^sup^、下标~sub~ 全部支持！

### 📄 多种页面尺寸
A4、Letter 任你选，满足不同场景需求。

### 🌈 现代化界面
基于 CustomTkinter 打造的现代 UI，清爽美观，用起来心情都好了！

---

## 🚀 快速开始

### 方式一：直接使用 EXE（推荐懒人）

1. 前往 [Releases](https://github.com/AGoodManaaa/MarkdowntoWord/releases) 页面
2. 下载最新的 `MarkdownToWord.exe`
3. 双击运行，完事！

### 方式二：从源码运行

```bash
# 1. 克隆仓库
git clone https://github.com/AGoodManaaa/MarkdowntoWord.git
cd MarkdowntoWord

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动程序
python gui.py
```

---

## 📖 使用方法

### 基本操作

1. **打开程序** → 左侧输入/粘贴 Markdown 内容
2. **实时预览** → 右侧自动显示渲染效果
3. **导出 Word** → 点击「导出」按钮，选择保存路径
4. **大功告成！** 🎉

### 快捷键

| 快捷键 | 功能 |
|--------|------|
| `Ctrl+S` | 保存/导出 |
| `Ctrl+O` | 打开文件 |
| `Ctrl+Z` | 撤销 |
| `Ctrl+Y` | 重做 |

---

## 🛠️ 技术栈

| 组件 | 用途 |
|------|------|
| [python-docx](https://python-docx.readthedocs.io/) | Word 文档生成 |
| [CustomTkinter](https://customtkinter.tomschimansky.com/) | 现代化 GUI |
| [latex2mathml](https://github.com/roniemartinez/latex2mathml) | LaTeX 公式转换 |
| [Pillow](https://pillow.readthedocs.io/) | 图片处理 |
| [Markdown](https://python-markdown.github.io/) | Markdown 解析 |

---

## 📁 项目结构

```
MarkdowntoWord/
├── main.py          # 🚀 程序入口
├── gui.py           # 🎨 图形界面
├── converter.py     # ⚙️ 核心转换逻辑
├── parser.py        # 📝 Markdown 解析器
├── handlers.py      # 🔧 各类处理器
├── styles.py        # 🎭 文档样式
├── utils.py         # 🛠️ 工具函数
└── requirements.txt # 📦 依赖清单
```

---

## 🤝 贡献

欢迎各种形式的贡献！无论是：

- 🐛 提交 Bug 报告
- 💡 提出新功能建议
- 📝 改进文档
- 🔧 提交 Pull Request

都非常欢迎！

### 贡献步骤

1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

---

## 📜 开源协议

本项目采用 [MIT License](LICENSE) 开源协议。

---

## 💖 致谢

感谢所有让这个项目成为可能的开源项目和贡献者们！

特别感谢：
- [python-docx](https://github.com/python-openxml/python-docx) 团队
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) 的作者 Tom Schimansky
- 所有提出建议和反馈的用户

---

<p align="center">
  <b>如果这个项目帮到了你，别忘了给个 ⭐ Star 哦～</b>
</p>

<p align="center">
  Made with ❤️ by <a href="https://github.com/AGoodManaaa">A Good Man</a>
</p>
