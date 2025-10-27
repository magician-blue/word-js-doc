# Process.py - 示例生成和增强工具使用说明

`process.py` 是一个自动化工具，用于为 Word.js API 文档的属性和方法生成或增强示例代码。

## 功能概览

该工具会自动处理 `jsonfied/` 目录中的 JSON 文件，并：

1. ✅ **为缺少示例的属性/方法生成新示例**
   - 包含 `description`（示例说明）
   - 包含 `usage_code`（TypeScript 代码）
   - `output_code` 设置为 `null`

2. ✅ **为已有代码但缺少描述的示例生成描述**
   - 分析现有的 `usage_code`
   - 自动生成简洁的 `description`

3. ✅ **保存处理后的文件到 `processed/` 目录**

## 快速开始

### 基本用法

```bash
# 处理单个文件
python process.py --file Word.Border.json

# 处理多个文件
python process.py --files Word.Border.json Word.Shading.json

# 处理所有文件
python process.py --all
```

### 测试模式

```bash
# 试运行（查看会做什么，不实际调用 API）
python process.py --file Word.Border.json --dry-run

# 限制每个文件最多生成 5 个示例（用于测试）
python process.py --file Word.Body.json --max-per-file 5
```

### 覆盖已存在的文件

```bash
# 默认会跳过已处理的文件
python process.py --all --overwrite
```

## 命令行参数

### 必选参数（三选一）

| 参数 | 说明 | 示例 |
|------|------|------|
| `--file FILE` | 处理单个文件 | `--file Word.Border.json` |
| `--files FILE1 FILE2 ...` | 处理多个文件 | `--files Word.Border.json Word.Shading.json` |
| `--all` | 处理所有文件 | `--all` |

### 可选参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--input-dir DIR` | `jsonfied` | 输入目录 |
| `--output-dir DIR` | `processed` | 输出目录 |
| `--model MODEL` | `gpt-4o-mini` | OpenAI 模型 |
| `--overwrite` | `False` | 覆盖已存在的文件 |
| `--dry-run` | `False` | 试运行模式 |
| `--skip-properties` | `False` | 跳过属性处理 |
| `--skip-methods` | `False` | 跳过方法处理 |
| `--max-per-file N` | 无限制 | 每个文件最多生成 N 个示例 |

## 使用场景

### 场景 1：为新文档生成所有示例

```bash
# 为 Word.Border.json 生成所有缺失的示例
python process.py --file Word.Border.json
```

**处理前**:
```json
{
  "name": "color",
  "type": "string",
  "description": "Specifies the color for the border.",
  "examples": []  // 空数组
}
```

**处理后**:
```json
{
  "name": "color",
  "type": "string",
  "description": "Specifies the color for the border.",
  "examples": [
    {
      "description": "Set the border color of a paragraph to #FF5733",
      "usage_code": "Word.run(async (context) => {\n    const paragraph = context.document.body.insertParagraph(\"Text\", Word.InsertLocation.end);\n    paragraph.getBorder(\"Top\").color = \"#FF5733\";\n    await context.sync();\n});",
      "output_code": null
    }
  ]
}
```

### 场景 2：为已有代码生成描述

如果你的 JSON 中已经有 `usage_code` 但缺少 `description`：

**处理前**:
```json
{
  "description": null,
  "usage_code": "await Word.run(async (context) => {\n  const body = context.document.body;\n  body.load('font/size');\n  await context.sync();\n});"
}
```

**处理后**:
```json
{
  "description": "Load the font size property of the document body.",
  "usage_code": "await Word.run(async (context) => {\n  const body = context.document.body;\n  body.load('font/size');\n  await context.sync();\n});"
}
```

### 场景 3：只处理属性或只处理方法

```bash
# 只生成属性的示例
python process.py --file Word.Border.json --skip-methods

# 只生成方法的示例
python process.py --file Word.Border.json --skip-properties
```

### 场景 4：批量处理多个文件

```bash
# 处理所有已转换的 JSON 文件
python process.py --all --overwrite

# 查看统计信息
# 输出示例：
#   Properties processed: 45
#   Methods processed: 38
#   Examples generated: 67
#   Descriptions generated: 16
```

## 工作流程示例

### 完整的文档处理流程

```bash
# 步骤 1: 将大文档转换为 JSON
python jsonfy_large_docs.py --all-large --overwrite

# 步骤 2: 为所有文档生成示例（先试运行看看）
python process.py --all --dry-run

# 步骤 3: 实际生成示例
python process.py --all

# 步骤 4: 检查结果
ls -lh processed/
```

### 渐进式处理（推荐用于大批量）

```bash
# 先处理几个文件测试
python process.py --files Word.Border.json Word.Shading.json

# 检查结果满意后，处理剩余文件
python process.py --all --overwrite
```

## Prompt 设计

工具使用精心设计的 prompt 来生成**任务-解决方案**格式的示例，**只提供必要的上下文信息**：

### 📝 任务-解决方案格式

- **description**: 具体的任务需求（要做什么）
- **usage_code**: 实现这个任务的代码（怎么做）

**示例:**
```json
{
  "description": "Set the border color to red",  // ← 任务需求
  "usage_code": "await Word.run(async (context) => {...});"  // ← 实现代码
}
```

### 生成属性示例的 Prompt

```
Create a concise TypeScript usage example for the following Word.js API property.

Class: Word.Border
Property: color
Type: string
Description: Specifies the color for the border.

Generate a practical example in task-solution format. Return in this exact format:

DESCRIPTION: [A concrete task/requirement, like "Set the font size to 16" or "Change the border color to red"]

CODE:
```typescript
[TypeScript code implementing the task above using Word.run async pattern]
```

The DESCRIPTION should be a specific task requirement (what to do), and the CODE should implement that task.
```

### 生成方法示例的 Prompt

```
Create a concise TypeScript usage example for the following Word.js API method.

Class: Word.Border
Method: load()
Parameters: options
Description: Queues up a command to load the specified properties...

Generate a practical example in task-solution format. Return in this exact format:

DESCRIPTION: [A concrete task/requirement, like "Insert a paragraph with text" or "Delete the first table"]

CODE:
```typescript
[TypeScript code implementing the task above using Word.run async pattern]
```

The DESCRIPTION should be a specific task requirement (what to do), and the CODE should implement that task.
```

### 生成任务描述的 Prompt（为已有代码）

```
Given this property usage example for Word.Border.color, write a brief task requirement that this code accomplishes.

Usage code:
[现有代码]

Write the task as a concrete requirement (what needs to be done), NOT an explanation of what the code does.

Examples:
- Good: "Set the font size of the paragraph to 14 points"
- Bad: "This code sets the font size of the paragraph to 14 points"

- Good: "Insert a new paragraph at the end of the document"
- Bad: "This example inserts a new paragraph at the end of the document"

Return ONLY the task requirement (1 sentence), no code, no extra formatting.
```

## 输出统计

处理完成后会显示详细统计信息：

```
Statistics:
  Properties processed: 25      ← 处理了多少个属性
  Methods processed: 18         ← 处理了多少个方法
  Examples generated: 32        ← 生成了多少个新示例
  Descriptions generated: 11    ← 生成了多少个描述
```

## 错误处理

工具会优雅地处理错误：

```bash
# 如果 API 调用失败
[ERROR] Failed to generate example: Rate limit exceeded

# 如果文件格式错误
[ERROR] Failed to read Word.Border.json: Invalid JSON

# 如果无法写入输出
[ERROR] Failed to write Word.Border.json: Permission denied
```

**失败的项目会被跳过**，其他项目继续处理。最后会显示失败列表。

## 性能考虑

### API 调用次数估算

对于一个典型的 Word.js 类（20 个属性 + 15 个方法）：
- 如果所有成员都缺少示例：**35 次 API 调用**
- 如果只需要生成描述：**~10-15 次 API 调用**

### 处理时间估算

- 单个示例生成：**~2-5 秒**
- 单个文件（35 个示例）：**~2-3 分钟**
- 批量处理 10 个文件：**~20-30 分钟**

### 建议

```bash
# 大批量处理时，先用 --max-per-file 测试
python process.py --all --max-per-file 5 --dry-run

# 然后逐步增加
python process.py --all --max-per-file 10

# 最后全量处理
python process.py --all
```

## 配置 OpenAI API

工具使用 OpenRouter API，配置在脚本中：

```python
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="sk-or-v1-...",
)
```

如果需要更改模型：

```bash
# 使用 GPT-4（更准确但更慢）
python process.py --file Word.Border.json --model gpt-4

# 使用 GPT-3.5（更快但可能质量稍低）
python process.py --file Word.Border.json --model gpt-3.5-turbo
```

## 输出格式

所有生成的示例都符合 `jsonfy.md` 规范：

```json
{
  "description": "Brief 1-sentence description",
  "usage_code": "TypeScript code using Word.run pattern",
  "output_code": null
}
```

## 常见问题

### Q: 生成的示例质量如何？
A: 工具使用精心设计的 prompt，生成的示例：
- ✅ **任务-解决方案**格式：description 是任务需求，code 是实现
- ✅ 使用标准的 `Word.run(async (context) => {...})` 模式
- ✅ 包含必要的 `load()` 和 `sync()` 调用
- ✅ 简洁实用，专注于当前属性/方法
- ✅ 符合 Word.js API 最佳实践

**示例:**
```json
{
  "description": "Set the font size to 16 points",  // ← 任务需求
  "usage_code": "await Word.run(async (context) => {\n  const paragraph = context.document.body.paragraphs.getFirst();\n  paragraph.font.size = 16;\n  await context.sync();\n});"
}
```

### Q: 可以自定义生成的示例吗？
A: 可以！修改 `ExampleGenerator` 类中的 prompt 模板：
- `_build_property_example_prompt()` - 属性示例
- `_build_method_example_prompt()` - 方法示例

### Q: 如果对某个生成的示例不满意怎么办？
A: 有两种方式：
1. 手动编辑 `processed/` 中的 JSON 文件
2. 删除该示例，重新运行脚本（使用 `--overwrite`）

### Q: 工具会修改原始文件吗？
A: **不会**。原始文件在 `jsonfied/` 中保持不变，所有输出都在 `processed/` 目录。

## 最佳实践

1. **先试运行** - 使用 `--dry-run` 查看会做什么
2. **小批量测试** - 使用 `--max-per-file` 限制生成数量
3. **增量处理** - 先处理几个文件，检查质量后再批量处理
4. **保留原文件** - `jsonfied/` 作为备份，只修改 `processed/`
5. **定期检查** - 处理过程中查看统计信息和错误日志

## 总结

`process.py` 是一个强大的自动化工具，可以：
- 🚀 **快速生成**高质量的 API 使用示例
- 📝 **自动补充**缺失的示例描述
- 🎯 **精准控制**处理范围和数量
- 📊 **详细统计**处理进度和结果

配合 `jsonfy_large_docs.py` 使用，可以完整地将 Markdown 文档转换为结构化、带示例的 JSON 文档！
