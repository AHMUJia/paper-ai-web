# 智能论文整理与 PubMed 校验系统

一个纯前端（HTML/CSS/JS）的小工具，用于：

- 从本地 `docx/txt` 解析参考文献条目
- 可选：通过 PubMed 补全作者/PMID/期刊等信息
- 自动匹配影响因子 / JCR 分区 / 中科院分区
- 导出为 **DOCX**（文件名：`整理后论文格式.docx`）

---

## 启动方式

### 方式一：使用 Python 本地服务器（推荐）

**Python 3.x：**
```bash
# 在项目目录下打开命令行/终端，执行：
python -m http.server 8000
```

**Python 2.x：**
```bash
python -m SimpleHTTPServer 8000
```

然后在浏览器中访问：`http://localhost:8000`

**Windows 快速操作：**
1. 在项目文件夹中，按住 `Shift` 键，右键点击空白处
2. 选择"在此处打开 PowerShell 窗口"或"在此处打开命令窗口"
3. 输入命令：`python -m http.server 8000`
4. 按回车执行
5. 打开浏览器，访问 `http://localhost:8000`

### 方式二：使用其他本地服务器

- **Node.js**: `npx http-server -p 8000`
- **PHP**: `php -S localhost:8000`
- **VS Code**: 安装 "Live Server" 扩展，右键 `index.html` 选择 "Open with Live Server"

### 为什么需要本地服务器？

直接双击打开 HTML 文件（`file://` 协议）可能会遇到浏览器的 CORS（跨域资源共享）限制，导致无法加载本地数据文件（如 `影响因子.txt`、`测试1.docx` 等）。使用本地服务器可以避免这个问题。

---

## 使用流程

1. **启动本地服务器**：按照上述方式启动服务器
2. **打开网页**：在浏览器中访问 `http://localhost:8000`
3. **选择文件**：点击"选择文件"选择你的 `docx/doc/txt`
4. **加载示例文件**（可选）：点击"加载示例文件"，仅会在左侧显示 `测试1.docx`（不会立刻解析）
5. **解析文档**：点击"第一步：解析文档"生成条目列表
6. **PubMed 补全**（可选）：点击"第二步：PubMed 补全"
7. **导出**：点击"导出整理后论文格式"，下载 `整理后论文格式.docx`
