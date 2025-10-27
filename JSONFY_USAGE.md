# Word.js 文档 JSON 转换工具使用说明

本项目提供了两种方式将 Word.js API 文档从 Markdown 格式转换为结构化的 JSON 格式。

## 工具概览

### 1. **jsonfy_classes.py** - API 驱动转换（适用于中小型文档）
使用 OpenAI API (GPT-5) 进行智能转换，适合处理结构复杂但文件不太大的文档。

### 2. **jsonfy_large_docs.py** - 本地解析转换（专为大文档设计）
使用本地 Markdown 解析器，专门针对大型、复杂的文档设计。不需要 API 调用。

---

## 方案一：API 驱动转换 (jsonfy_classes.py)

### 适用场景
- 中小型文档（通常 < 50KB）
- 文档结构复杂，需要智能理解
- 有 OpenAI API 访问权限

### 使用方法

```bash
# 安装依赖
pip install openai

# 转换 class.md 中列出的所有类
python jsonfy_classes.py

# 只转换前 5 个类（测试用）
python jsonfy_classes.py --limit 5

# 覆盖已存在的 JSON 文件
python jsonfy_classes.py --overwrite

# 试运行（不实际调用 API）
python jsonfy_classes.py --dry-run
```

### 配置参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--class-table` | `class.md` | 类列表文件 |
| `--docs-dir` | `api_docs` | 文档目录 |
| `--prompt-template` | `jsonfy.md` | 转换提示模板 |
| `--output-dir` | `jsonfied` | 输出目录 |
| `--model` | `gpt-5.0-mini` | 使用的模型 |
| `--limit` | 无 | 限制处理数量 |
| `--overwrite` | False | 覆盖已存在文件 |

---

## 方案二：本地解析转换 (jsonfy_large_docs.py) ⭐ 推荐用于大文档

### 适用场景
- **大型文档**（> 50KB）
- **复杂文档**（如 Word.Range, Word.Paragraph, Word.Document, Word.Body）
- 不依赖外部 API
- 需要快速批量处理

### 核心优势
✅ 无需 API，完全本地处理
✅ 基于标题层级智能解析
✅ 自动提取类型信息（从 TypeScript 代码块）
✅ 完整提取示例代码
✅ 处理速度快，适合大批量转换

### 使用方法

#### 1. 处理单个文件
```bash
python jsonfy_large_docs.py --file Word.Body.md
```

#### 2. 处理多个指定文件
```bash
python jsonfy_large_docs.py --files Word.Body.md Word.Range.md Word.Paragraph.md
```

#### 3. 自动处理所有大文件（推荐）
```bash
# 处理所有 > 50KB 的文件
python jsonfy_large_docs.py --all-large

# 处理所有 > 60KB 的文件
python jsonfy_large_docs.py --all-large --size-threshold 60000
```

#### 4. 覆盖已存在的文件
```bash
python jsonfy_large_docs.py --all-large --overwrite
```

#### 5. 自定义输出路径
```bash
python jsonfy_large_docs.py --file Word.Body.md --output custom_output.json
```

#### 6. 试运行（查看会处理哪些文件）
```bash
python jsonfy_large_docs.py --all-large --dry-run
```

### 配置参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--docs-dir` | `api_docs` | 文档目录 |
| `--output-dir` | `jsonfied` | 输出目录 |
| `--file` | - | 处理单个文件 |
| `--files` | - | 处理多个文件 |
| `--all-large` | - | 处理所有大文件 |
| `--size-threshold` | `50000` | 大文件阈值（字节） |
| `--output` | - | 自定义输出文件（仅单文件） |
| `--overwrite` | False | 覆盖已存在文件 |
| `--pretty` | True | 格式化 JSON 输出 |
| `--dry-run` | False | 试运行模式 |

---

## 工作原理

### md_parser.py - Markdown 解析器模块

这是本地转换方案的核心模块，提供了三个主要类：

#### 1. **MarkdownSection**
表示文档中的一个标题区块，包含：
- `level`: 标题级别（1-6）
- `title`: 标题文本
- `content`: 该标题下的内容（不包括子标题）
- `children`: 子标题列表

#### 2. **MarkdownParser**
解析 Markdown 文档的层级结构：
- `parse()`: 解析整个文档为树形结构
- `find_section(title)`: 根据标题查找区块
- `find_sections_by_level(level)`: 查找特定级别的所有标题

#### 3. **APIDocParser**
专门用于解析 Word.js API 文档：
- `extract_class_info()`: 提取类信息（名称、包、继承、API set、描述、示例）
- `extract_properties()`: 提取属性（名称、类型、描述、API版本、示例）
- `extract_methods()`: 提取方法（名称、类型、签名、参数、返回值、示例）
- `to_json_schema()`: 转换为符合 jsonfy.md 规范的 JSON

### 文档结构识别

解析器能够识别以下 Markdown 结构：

```markdown
# Word.ClassName class          ← 类名
Package: [word](...)            ← 包名
Extends: [BaseClass](...)       ← 继承关系
描述文本...                      ← 类描述

## Remarks                       ← API set 和类示例
[API set: WordApi 1.1]
```typescript
示例代码...
```

## Properties                    ← 属性列表概览
- property1
- property2

## Methods                       ← 方法列表概览
- method1()
- method2()

## Property Details              ← 属性详情
### propertyName                 ← 具体属性
描述...
```typescript
readonly propertyName: Type;    ← 类型信息
```
#### Examples                    ← 属性示例
```typescript
示例代码...
```

## Method Details                ← 方法详情
### methodName(param1, param2)   ← 具体方法
描述...
#### Examples                    ← 方法示例
```

---

## 输出格式

两种工具都生成符合 `jsonfy.md` 规范的 JSON 文件：

```json
{
  "class": {
    "name": "Word.Body",
    "package": "word",
    "extends": ["OfficeExtension.ClientObject"],
    "api_set": {
      "name": "WordApi",
      "status": "1.1"
    },
    "description": "Represents the body of a document or a section.",
    "examples": [...]
  },
  "properties": [
    {
      "name": "font",
      "type": "Word.Font",
      "description": "Gets the text format of the body...",
      "since": "WordApi 1.1",
      "examples": [...]
    }
  ],
  "methods": [
    {
      "name": "clear",
      "kind": "delete",
      "description": "Clears the contents of the body object...",
      "signatures": [...],
      "examples": [...]
    }
  ],
  "source": {
    "urls": ["https://docs.microsoft.com/en-us/javascript/api/word"]
  }
}
```

---

## 推荐工作流程

### 对于大型、复杂文档（推荐）

```bash
# 1. 先识别哪些是大文件
python jsonfy_large_docs.py --all-large --dry-run

# 2. 处理所有大文件
python jsonfy_large_docs.py --all-large --overwrite

# 3. 检查输出
ls -lh jsonfied/
```

### 对于中小型文档

```bash
# 使用 API 驱动方案
python jsonfy_classes.py --limit 10  # 先测试几个
python jsonfy_classes.py --overwrite # 全部处理
```

---

## 已成功转换的大文档

以下文档已使用 `jsonfy_large_docs.py` 成功转换：

| 文档 | 大小 | 属性数 | 方法数 | 状态 |
|------|------|--------|--------|------|
| Word.Range.md | 153.4 KB | 62 | 48 | ✅ |
| Word.Paragraph.md | 94.2 KB | 39 | 63 | ✅ |
| Word.Body.md | 71.7 KB | 21 | 24 | ✅ |
| Word.Document.md | 63.2 KB | 24 | 25 | ✅ |

---

## 故障排除

### 问题：属性的 type 字段为 null
**原因**：文档中没有 TypeScript 类型声明
**解决**：检查文档中是否有 \`\`\`typescript 代码块

### 问题：类示例 (class.examples) 为空
**原因**：Remarks 部分没有示例代码
**解决**：检查 `## Remarks` 和 `### Examples` 部分

### 问题：方法签名解析不正确
**原因**：方法标题格式不标准
**解决**：确保方法标题格式为 `### methodName(param1, param2)`

---

## 扩展和定制

如果需要处理其他格式的文档，可以修改 `md_parser.py`：

1. **修改标题识别**：调整 `_extract_sections()` 方法
2. **添加新字段**：在 `extract_properties()` 或 `extract_methods()` 中添加
3. **自定义类型推断**：修改 `_infer_method_kind()` 方法

---

## 总结

- **小文档** → 使用 `jsonfy_classes.py`（API 驱动）
- **大文档** → 使用 `jsonfy_large_docs.py`（本地解析）⭐
- **批量处理** → `python jsonfy_large_docs.py --all-large --overwrite`

两种方案互补，可根据实际需求选择合适的工具！
