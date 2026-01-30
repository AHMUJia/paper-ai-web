# 智能论文整理与 PubMed 校验系统

一个纯前端（HTML/CSS/JS）的小工具，用于：

- 从本地 `docx/txt` 解析参考文献条目
- 可选：通过 PubMed 补全作者/PMID/期刊等信息
- 自动匹配影响因子 / JCR 分区 / 中科院分区
- 导出为 **DOCX**（文件名：`整理后论文格式.docx`）

---

## 🌐 在线使用（推荐）

**直接访问：** [https://ahmujia.github.io/paper-ai-web/](https://ahmujia.github.io/paper-ai-web/)

打开上述链接即可直接使用，无需本地安装或配置！

---

## 💻 本地启动方式

### 使用 Python 本地服务器（推荐）

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

---

## 使用流程

### 在线使用（推荐）

1. **打开网页**：访问 [https://ahmujia.github.io/paper-ai-web/](https://ahmujia.github.io/paper-ai-web/)
2. **选择文件**：点击"选择文件"选择你的 `docx/doc/txt`
3. **解析文档**：点击"第一步：解析文档"生成条目列表
4. **PubMed 补全**（可选）：点击"第二步：PubMed 补全"
5. **编辑期刊名称**（可选）：点击期刊名称进行编辑，系统会自动更新影响因子和分区
6. **标注作者**（可选）：点击"标注共一/共通讯"按钮，在弹窗中勾选需要标注的作者
7. **导出**：点击"导出整理后论文格式"，下载 `整理后论文格式.docx`

### 本地使用

1. **启动本地服务器**：按照上述方式启动服务器
2. **打开网页**：在浏览器中访问 `http://localhost:8000`
3. 后续步骤与在线使用相同
