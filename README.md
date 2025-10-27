# Word.js API 文档 JSON 化工具链

将 Word.js API 的 Markdown 文档转换为结构化的 JSON 格式，并自动生成使用示例。

## 🔄 完整工作流程

```
Markdown 文档 → JSON 转换 → 示例生成 → 最终 JSON
  (api_docs/)     (jsonfied/)     (processed/)
```

## 📚 核心工具

### 1️⃣ **jsonfy_classes.py** - API 驱动的 JSON 转换
- 使用 OpenAI API 智能转换 Markdown → JSON
- 适合中小型文档（< 50KB）
- 详见：[JSONFY_USAGE.md](JSONFY_USAGE.md)

### 2️⃣ **jsonfy_large_docs.py** - 本地解析的 JSON 转换 ⭐
- 基于标题层级的本地解析器
- 专为大型文档设计（> 50KB）
- 无需 API，速度快
- 已成功处理：Word.Range (153KB), Word.Paragraph (94KB), Word.Body (71KB), Word.Document (63KB)
- 详见：[JSONFY_USAGE.md](JSONFY_USAGE.md)

### 3️⃣ **process.py** - 示例生成和增强工具 ⭐
- 为属性和方法生成使用示例
- 为已有代码生成描述
- 使用 OpenAI API 生成高质量示例
- 详见：[PROCESS_USAGE.md](PROCESS_USAGE.md)

### 4️⃣ **md_parser.py** - Markdown 解析器模块
- 核心解析引擎
- 支持层级结构解析
- 智能提取类、属性、方法信息

## 📁 目录结构

```
word-js-doc/
├── api_docs/              # 原始 Markdown 文档
├── api_docs_all/          # api_docs 的备份
├── jsonfied/              # 转换后的 JSON 文件
├── processed/             # 增强后的 JSON 文件（含生成的示例）
│
├── jsonfy_classes.py      # API 驱动的转换脚本
├── jsonfy_large_docs.py   # 本地解析的转换脚本
├── process.py             # 示例生成脚本
├── md_parser.py           # Markdown 解析器模块
│
├── jsonfy.md              # JSON Schema 规范
├── class.md               # 类列表和处理进度跟踪
│
├── JSONFY_USAGE.md        # JSON 转换工具使用说明
├── PROCESS_USAGE.md       # 示例生成工具使用说明
└── README.md              # 本文件
```

## 🚀 快速开始

### 方式 1：完整流程（从 Markdown 到带示例的 JSON）

```bash
# 步骤 1: 转换大文档为 JSON
python jsonfy_large_docs.py --all-large --overwrite

# 步骤 2: 生成示例和描述
python process.py --all

# 结果：processed/ 目录包含完整的 JSON 文档
```

### 方式 2：单个文件处理

```bash
# 转换单个文档
python jsonfy_large_docs.py --file Word.Body.md

# 生成示例
python process.py --file Word.Body.json
```

### 方式 3：只转换 JSON（不生成示例）

```bash
# 使用本地解析器
python jsonfy_large_docs.py --all-large

# 或使用 API（适合小文档）
python jsonfy_classes.py --limit 10
```

## 📊 处理状态

### 已完成的大文档转换

| 文档 | 大小 | 属性数 | 方法数 | 状态 |
|------|------|--------|--------|------|
| Word.Range.md | 153.4 KB | 62 | 48 | ✅ |
| Word.Paragraph.md | 94.2 KB | 39 | 63 | ✅ |
| Word.Body.md | 71.7 KB | 21 | 24 | ✅ |
| Word.Document.md | 63.2 KB | 24 | 25 | ✅ |

## 📖 详细文档

- **[JSONFY_USAGE.md](JSONFY_USAGE.md)** - JSON 转换工具详细说明
  - 两种转换方案对比
  - 完整的命令行参数说明
  - 使用场景和示例

- **[PROCESS_USAGE.md](PROCESS_USAGE.md)** - 示例生成工具详细说明
  - 自动生成示例的工作原理
  - Prompt 设计说明
  - 批量处理最佳实践

## 🎯 输出格式

所有工具都生成符合 `jsonfy.md` 规范的 JSON：

```json
{
  "class": {
    "name": "Word.Body",
    "package": "word",
    "description": "...",
    "examples": [...]
  },
  "properties": [
    {
      "name": "font",
      "type": "Word.Font",
      "description": "...",
      "examples": [
        {
          "description": "Sets the font size of the document body",
          "usage_code": "await Word.run(async (context) => {...});",
          "output_code": null
        }
      ]
    }
  ],
  "methods": [...]
}
```

## 🔧 安装依赖

```bash
pip install openai
```

## 💡 使用技巧

1. **大文档优先使用本地解析器**
   ```bash
   python jsonfy_large_docs.py --all-large
   ```

2. **先试运行，再实际处理**
   ```bash
   python process.py --all --dry-run
   python process.py --all
   ```

3. **限制处理数量进行测试**
   ```bash
   python process.py --file Word.Body.json --max-per-file 5
   ```

4. **增量处理，逐步验证**
   ```bash
   # 先处理几个文件
   python process.py --files Word.Border.json Word.Shading.json

   # 验证结果后批量处理
   python process.py --all --overwrite
   ```

## 🎨 工具特性

### jsonfy_large_docs.py
✅ 基于 Markdown 标题层级解析
✅ 自动提取 TypeScript 类型信息
✅ 完整提取示例代码
✅ 无需 API，速度快
✅ 适合大型复杂文档

### process.py
✅ 智能生成使用示例
✅ 自动补充示例描述
✅ 精简的 prompt 设计
✅ 批量处理支持
✅ 详细的统计信息

## 📈 性能参考

- **JSON 转换**（本地解析）：~1 秒/文件
- **示例生成**（API 调用）：~2-5 秒/示例
- **完整处理** 单个大文档：~2-5 分钟

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 License

MIT

---

## Legacy Information

### Failed JSONfy Trials (使用 API 转换时失败)
以下大文档使用 `jsonfy_classes.py` (API 驱动) 失败，现已使用 `jsonfy_large_docs.py` (本地解析) 成功转换：

- ✅ Word.Body (已成功)
- ✅ Word.Bookmark (已成功)
- ✅ Word.ContentControl (已成功)
- ✅ Word.ContentControlCollection (已成功)
- ✅ Word.Document (已成功)
- ✅ Word.Field (已成功)
- ✅ Word.Paragraph (已成功)
- ✅ Word.Range (已成功)
- ✅ Word.TableRow (已成功)

这些文档现在都可以在 `jsonfied/` 目录中找到对应的 JSON 文件。
