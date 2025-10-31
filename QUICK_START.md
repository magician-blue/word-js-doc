# 快速开始指南 - TypeScript 示例检查器

## 🚀 5 分钟上手

### 1. 检查单个文件 (推荐新手)

```bash
python check_examples.py processed/Word.Bookmark.json
```

**预期输出:**
```
Checking single file: processed/Word.Bookmark.json

[SETUP] Creating persistent TypeScript environment...
[SETUP] Installing packages...
[SETUP] Environment ready!

[1/16] [PROPERTY] Bookmark.context (property, example 1)
❌ FAIL - 1 error(s)

...

Summary: 0 passed, 16 failed out of 16 examples
Pass rate: 0.0%
Total time: 15.97s (avg 0.998s per example)
```

### 2. 检查所有文件

```bash
python check_examples.py
```

**预期输出:**
```
Found 25 JSON files to check
Setting up TypeScript environment (this happens only once)...

[1/25] Checking: Word.Bookmark.json
  [PROPERTY] Bookmark.context (property, example 1)
  ❌ FAIL - 1 error(s)

[2/25] Checking: Word.Body.json
  [PROPERTY] Body.style (property, example 1)
  ✅ PASS

...

FINAL SUMMARY
Total files checked: 25
Total examples checked: 342
  ✅ Passed: 287
  ❌ Failed: 55
Pass rate: 83.9%
Total time: 42.15s

Detailed results saved to: check_results.json
```

### 3. 验证检查器准确性

```bash
python test_validation.py
```

**预期输出:**
```
[测试 1] 正确的代码 - 期望 PASS
验证: ✅ 正确

[测试 2] 错误的属性名 - 期望 FAIL
验证: ✅ 正确 - 成功检测到错误

...

验证通过: 5/5
🎉 所有验证通过! parser_optimized.py 工作正常!
```

## 📊 理解输出

### ✅ PASS - 代码正确
```
[PROPERTY] Range.text (property, example 1)
Description: Get the text content of a range
✅ PASS
```
说明这个示例代码没有 TypeScript 错误。

### ❌ FAIL - 代码有错误
```
[METHOD] Bookmark.delete() (method, example 1)
Description: Delete a bookmark
❌ FAIL - 1 error(s)
Errors:
  Line 5:43 - TS2339
  Property 'getBookmarkByName' does not exist on type 'Document'.
```

**错误解读:**
- `Line 5:43` - 错误在第 5 行,第 43 列
- `TS2339` - TypeScript 错误代码 (属性不存在)
- 具体信息: `getBookmarkByName` 方法不存在

## 📁 生成的文件

### check_results.json
完整的检查结果,JSON 格式:
```json
{
  "total_files": 25,
  "total_examples": 342,
  "passed": 287,
  "failed": 55,
  "files": [...]
}
```

可以用其他工具处理这个 JSON 文件,例如:
- 生成 HTML 报告
- 导入到数据库
- 统计分析

## ⚡ 性能对比

| 操作 | 时间 | 说明 |
|-----|------|------|
| 首次检查 | ~10秒 | 安装 npm 包 |
| 后续检查 | ~0.5秒/个 | 重用环境 ⚡ |
| 检查 16 个示例 | ~16秒 | 比原版快 10 倍! |
| 检查 100 个示例 | ~60秒 | 比原版快 16 倍! |

## 🔍 常见场景

### 场景 1: 快速验证单个类的示例
```bash
python check_examples.py processed/Word.Range.json
```

### 场景 2: 检查所有文档质量
```bash
python check_examples.py > results.txt 2>&1
```

### 场景 3: 只看失败的示例
```bash
python check_examples.py | grep -A 3 "❌ FAIL"
```

### 场景 4: 统计通过率
```bash
python check_examples.py | grep "Pass rate"
```

## 🛠️ 故障排除

### 问题: npm 未找到
```
RuntimeError: npm not found on PATH
```

**解决**: 安装 Node.js
- 下载: https://nodejs.org/
- 安装后重启终端

### 问题: 编码错误 (Windows)
```
UnicodeEncodeError: 'gbk' codec can't encode character
```

**解决**: 已自动修复,如果仍有问题:
```bash
chcp 65001  # 设置 UTF-8 编码
python check_examples.py
```

### 问题: 检查很慢
**原因**: 第一次运行需要安装 npm 包 (~10秒)
**解决**: 正常现象,后续会快很多

### 问题: 结果不准确
```bash
# 运行验证测试
python test_validation.py
```
如果 5/5 测试通过,说明检查器工作正常。

## 📚 深入学习

- **使用说明**: 查看 `check_examples_usage.md`
- **技术详解**: 查看 `OPTIMIZATION_README.md`
- **完整总结**: 查看 `OPTIMIZATION_SUMMARY.md`

## 💡 小技巧

### 技巧 1: 只看错误摘要
```bash
python check_examples.py | grep -E "(Checking:|FAIL|Summary)"
```

### 技巧 2: 保存结果到文件
```bash
python check_examples.py > check_log.txt 2>&1
```

### 技巧 3: 检查特定模式的文件
```bash
for file in processed/Word.*.json; do
    python check_examples.py "$file"
done
```

### 技巧 4: 统计错误类型
```bash
python check_examples.py | grep "TS[0-9]" | sort | uniq -c
```

## ✅ 检查清单

使用前确认:
- [ ] 已安装 Node.js 和 npm
- [ ] 已安装 Python 3
- [ ] `processed/` 目录存在且包含 JSON 文件
- [ ] 运行 `python test_validation.py` 测试通过

## 🎯 下一步

1. ✅ 检查单个文件熟悉输出
2. ✅ 检查所有文件获得全局视图
3. ✅ 查看 `check_results.json` 详细分析
4. ✅ 根据错误修复文档
5. ✅ 重新检查验证修复

## 📞 获取帮助

遇到问题?
1. 查看文档: `check_examples_usage.md`
2. 运行验证: `python test_validation.py`
3. 检查日志: 查看错误信息

---

**开始检查吧!** 🚀

```bash
python check_examples.py processed/Word.Bookmark.json
```
