# PasteMD
<p align="center">
  <img src="assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

<p align="center">
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/v/release/RICHQAQ/PasteMD?sort=semver&label=Release&style=flat-square&logo=github" alt="Release">
  </a>
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/downloads/RICHQAQ/PasteMD/total?label=Downloads&style=flat-square&logo=github" alt="Downloads">
  </a>
  <a href="LICENSE">
    <img src="https://img.shields.io/github/license/RICHQAQ/PasteMD?style=flat-square" alt="License">
  </a>
  <img src="https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.12+">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Word%20%7C%20WPS-5e8d36?style=flat-square&logo=windows&logoColor=white" alt="Platform">
</p>

<p align="center"> 
  <a href="docs/md/README.en.md">English</a>
  |
  <a href="README.md">简体中文</a>
</p>

> 在写论文或报告时，从 ChatGPT / DeepSeek 等 AI 网站中复制出来的公式在 Word 里总是乱码？Markdown 表格复制到 Excel 总是不行？**PasteMD 就是为了解决这个问题而生的，嘿嘿**
> 
> <img src="docs/gif/atri/igood.gif"
     alt="我可是高性能的"
     width="100">

一个常驻托盘的小工具：
从 **剪贴板读取 Markdown**，调用 **Pandoc** 转换为 DOCX，并自动插入到 **Word/WPS** 光标位置。

**✨ 新功能**：智能识别 Markdown 表格，一键粘贴到 **Excel**！

**✨ 新功能**：智能识别 HTML富文本，方便直接复制网页上的ai回复，一键粘贴到 **Word/WPS**！

---

## 功能特点

### 演示效果

#### Markdown → Word/WPS

<p align="center">
  <img src="docs/gif/demo.gif" alt="演示动图" width="600">
</p>

#### 复制网页中的ai回复 → Word/WPS
<p align="center">
  <img src="docs/gif/demo-html.gif" alt="演示HTML动图" width="600">
</p>

#### Markdown 表格 → Excel
<p align="center">
  <img src="docs/gif/demo-excel.gif" alt="演示Excel动图" width="600">
</p>

#### 设置格式
<p align="center">
  <img src="docs/gif/demo-chage_format.gif" alt="演示设置格式动图" width="600">
</p>


* 全局热键（默认 `Ctrl+B`）一键粘贴 Markdown → DOCX。
* **✨ 智能识别 Markdown 表格**，自动粘贴到 Excel。
* 自动识别当前前台应用：Word 或 WPS。
* 智能打开所需应用为Word/Excel。
* 托盘菜单，可保留文件、查看日志/配置等。
* 支持系统通知提醒。
* 无黑框，无阻塞，稳定运行。

---

## 📊 AI 网站兼容性测试

以下是主流 AI 对话网站的复制粘贴兼容性测试结果：

| AI 网站 | 复制 Markdown<br/>（无公式） | 复制 Markdown<br/>（含公式） | 复制网页内容<br/>（无公式） | 复制网页内容<br/>（含公式） |
|---------|:----------------------------:|:----------------------------:|:---------------------------:|:---------------------------:|
| **Kimi** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ⚠️ 无法显示公式 |
| **DeepSeek** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |
| **通义千问** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ⚠️ 无法显示公式 |
| **豆包\*** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |
| **智谱清言<br/>/ChatGLM** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |
| **ChatGPT** | ✅ 完美支持 | ⚠️ 公式显示为代码 | ✅ 完美支持 | ✅ 完美支持 |
| **Gemini** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |
| **Grok** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |
| **Claude** | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 | ✅ 完美支持 |

**图例说明：**
- ✅ **完美支持**：格式、样式、公式会均正确显示
- ⚠️ **公式显示为代码**：数学公式会以 LaTeX 代码形式显示，需在 Word/WPS 中手动使用公式编辑器
- ⚠️ **无法显示公式**：数学公式会丢失，需在 Word/WPS 中手动使用公式编辑器，自行输入公式内容
- **豆包**：复制网页内容（含公式）前，需要在浏览器中开启“允许读取剪贴板”权限，可在 URL 地址栏左侧的图标中进行设置

**测试说明：**
1. **复制 Markdown**：点击 AI 回复中的"复制"按钮（通常复制的是 Markdown 格式，但是部分网站也会携带上html）
2. **复制网页内容**：直接选中 AI 回复内容进行复制（复制的是 HTML 富文本）

---

## 🚀使用方法

1. 下载可执行文件（[Releases 页面](https://github.com/RICHQAQ/PasteMD/releases/)）：

   * **PasteMD\_vx.x.x.exe**：**便携版**，需要你本机已经安装好 **Pandoc** 并能在命令行运行。
   若未安装，请到 [Pandoc 官网](https://pandoc.org/installing.html) 下载安装即可。
   * **PasteMD\_pandoc-Setup.exe**：**一体化安装包**，自带 Pandoc，不需要另外配置环境。

2. 打开 Word、WPS 或 Excel，光标放在需要插入的位置。

3. 复制 **Markdown** 或者 **网页内容** 到剪贴板，按下热键 **Ctrl+B**。

4. 转换结果会自动插入到文档中：
   - **Markdown 表格** → 自动粘贴到 Excel（如果 Excel 已打开）
   - **普通 Markdown**/**网页内容** → 转换为 DOCX 并插入 Word/WPS

5. 右下角会提示成功/失败。

---

## ⚙️配置

首次运行会生成 `config.json`，可手动编辑：

```json
{
  "hotkey": "<ctrl>+b",
  "pandoc_path": "pandoc",
  "reference_docx": null,
  "save_dir": "%USERPROFILE%\\Documents\\pastemd",
  "keep_file": false,
  "notify": true,
  "enable_excel": true,
  "excel_keep_format": true,
  "auto_open_on_no_app": true,
  "md_disable_first_para_indent": true,
  "html_disable_first_para_indent": true,
  "html_formatting": {
    "strikethrough_to_del": true
  },
  "move_cursor_to_end": true,
  "Keep_original_formula": false,
  "language": "zh"
}
```

字段说明：

* `hotkey`：全局热键，语法如 `<ctrl>+<alt>+v`。
* `pandoc_path`：Pandoc 可执行文件路径。
* `reference_docx`：Pandoc 参考模板（可选）。
* `save_dir`：保留文件时的保存目录。
* `keep_file`：是否保留生成的 DOCX 文件。
* `notify`：是否显示系统通知。
* **`enable_excel`**：**✨ 新功能** - 是否启用智能识别 Markdown 表格并粘贴到 Excel（默认 true）。
* **`excel_keep_format`**：**✨ 新功能** - Excel 粘贴时是否保留 Markdown 格式（粗体、斜体、代码等），默认 true。
* **`auto_open_on_no_app`**：**✨ 新功能** 当未检测到目标应用（如 Word/Excel）时，是否自动创建文件并用系统默认应用打开（默认 true）。
* **`md_disable_first_para_indent`**： - Markdown 转换时是否禁用第一段的特殊格式，统一为正文样式（默认 true）。
* **`html_formatting`**： - HTML 富文本转换时的格式化选项。
  * **`strikethrough_to_del`**： - 是否将删除线 ~~ 转换为 `<del>` 标签，使得转换正确（默认 true）。
* **`html_disable_first_para_indent`**： - HTML 富文本转换时是否禁用第一段的特殊格式，统一为正文样式（默认 true）。
* **`move_cursor_to_end`**：**✨ 新功能** - 插入内容后是否将光标移动到插入内容的末尾（默认 true）。
* **`Keep_original_formula`**：**✨ 新功能** - 是否保留原始数学公式（LaTeX 代码形式）。
* `language`：界面语言，`zh` 中文，`en` 英文。

修改后可在托盘菜单选择 **“重载配置/热键”** 立即生效。

---

## 托盘菜单

* 快捷显示：当前全局热键（只读）。
* 启用热键：开/关全局热键。
* 弹窗通知：开/关系统通知。
* 无应用时自动打开：当未检测到 Word/Excel 时是否自动创建并用默认应用打开。
* 插入后移动光标到末尾：插入内容后是否将光标移动到插入内容的末尾。
* HTML 格式化：切换 **删除线 ~~ 转换为 `<del>`** 等 HTML 自动整理，使得可以正确转换（防止部分网页没有解析这些格式，导致从网页复制粘贴无法显示这些格式）。
* 实验性功能：启用/禁用 **保留原始数学公式** 等实验性功能。
* 设置热键：通过图形界面录制并保存新的全局热键（即时生效）。
* 保留生成文件：勾选后生成的 DOCX 会保存在 `save_dir`。
* 打开保存目录、查看日志、编辑配置、重载配置/热键。
* 版本：显示当前版本；可检查更新；若检测到新版本，会显示条目并可点击打开下载页面。
* 退出：退出程序。

---

## 📦从源码运行 / 打包

建议 Python 3.12 (64位)。

```bash
pip install -r requirements.txt
python main.py
```

使用 PyInstaller：

```bash
pyinstaller --clean -F -w -n PasteMD
  --icon assets\icons\logo.ico
  --add-data "assets\icons;assets\icons"
  --add-data "pastemd\i18n\locales;pastemd\i18n\locales"
  --add-data "pastemd\lua;pastemd\lua"
  --hidden-import plyer.platforms.win.notification
  --hidden-import pastemd.i18n.locales.zh
  --hidden-import pastemd.i18n.locales.en
  main.py
```

生成的程序在 `dist/PasteMD.exe`。

---

## ⭐ Star 

感谢每一位 Star 的帮助，欢迎分享给更多小伙伴~，想要达成1k star🌟，我会努力的喵

<img src="docs/gif/atri/likeyou.gif"
     alt="喜欢你"
     width="150">

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=date&legend=top-left)](https://www.star-history.com/#RICHQAQ/PasteMD&type=date&legend=top-left)

## 🍵支持与打赏

如果有什么想法和好建议，欢迎issue交流！🤯🤯🤯

希望这个小工具对你有帮助，欢迎请作者👻喝杯咖啡☕～你的支持会让我更有动力持续修复问题、完善功能、适配更多场景并保持长期维护。感谢每一份支持！

<img src="docs/gif/atri/flower.gif"
     alt="送你一朵小花"
     width="150">
     
| 支付宝 | 微信 |
| --- | --- |
| ![支付宝打赏](docs/pay/Alipay.jpg) | ![微信打赏](docs/pay/Weixinpay.png) |


---

## License

This project is licensed under the [MIT License](LICENSE).
