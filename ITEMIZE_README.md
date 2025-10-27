# JSON to Markdown Converter (itemize.py)

这个工具可以将 `processed/` 目录中的 JSON 格式 API 文档转换回 Markdown 格式。

## 文件说明

- **itemize.py** - 单文件转换脚本
- **batch_itemize.py** - 批量转换脚本

## 使用方法

### 1. 单文件转换

```bash
# 基本用法（自动生成输出文件名）
python itemize.py processed/Word.Border.json

# 指定输出文件名
python itemize.py processed/Word.Border.json output/Word.Border.md
```

### 2. 批量转换所有文件

```bash
# 使用默认目录（输入: processed/, 输出: markdown_output/）
python batch_itemize.py

# 指定输入目录
python batch_itemize.py processed

# 指定输入和输出目录
python batch_itemize.py processed markdown_docs
```

## 输出格式

生成的 Markdown 文件包含以下部分：

1. **类标题和基本信息**
   - 类名
   - Package 信息
   - API Set 版本
   - 继承关系

2. **描述**
   - 类的详细描述

3. **类级别示例**
   - 完整的使用示例代码

4. **属性 (Properties)**
   - 属性名称
   - 类型
   - 描述
   - Since 版本信息
   - 示例代码

5. **方法 (Methods)**
   - 方法名称
   - Kind（操作类型）
   - 描述
   - 签名（参数和返回值）
   - 示例代码

6. **来源信息**
   - 文档来源 URL

## 示例输出

```markdown
# Word.Border

**Package:** `word`

**API Set:** WordApiDesktop 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the Border object for text, a paragraph, or a table.

## Properties

### color

**Type:** `string`

**Since:** WordApiDesktop 1.1

Specifies the color for the border.

#### Examples

**Example**: Set the border color of the first paragraph to red

\`\`\`typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    border.color = "#FF0000";

    await context.sync();
});
\`\`\`

---

### type

**Type:** `Word.BorderType | "Mixed" | "None" | "Single" | ...`

...
```

## 特点

- ✅ 完整保留所有 JSON 数据
- ✅ 格式化的代码块（TypeScript 语法高亮）
- ✅ 清晰的层次结构
- ✅ 支持多个方法签名（重载）
- ✅ 自动处理 UTF-8 编码
- ✅ 批量转换支持

## 技术细节

- 输入格式：JSON（来自 processed/ 目录）
- 输出格式：Markdown（GitHub 风格）
- 编码：UTF-8
- 代码块语言：TypeScript

## 依赖

- Python 3.6+
- 无第三方依赖（仅使用标准库）
