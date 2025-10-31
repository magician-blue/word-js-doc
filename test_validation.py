#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 parser_optimized.py 的检查结果是否准确
"""
import sys
import io
from parser_optimized import check_officejs_ts_fast, cleanup_persistent_env

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

print("="*80)
print("验证 TypeScript 检查结果的准确性")
print("="*80)

# Test 1: 正确的代码 - 应该 PASS
print("\n[测试 1] 正确的代码 - 期望 PASS")
print("-"*80)
correct_code = """
/// <reference types="office-js-preview" />
async function test1() {
    await Word.run(async (context) => {
        const range = context.document.body.getRange();
        range.select(Word.SelectionMode.end);
        const newRange = range.expandTo(range.paragraphs.getFirst().getRange());
        newRange.font.highlightColor = "#FFFF00";
        await context.sync();
    });
}
"""
result1 = check_officejs_ts_fast(correct_code, use_preview=True)
print(f"结果: {'✅ PASS' if result1['success'] else '❌ FAIL'}")
print(f"错误数: {result1['error_count']}")
if not result1['success']:
    print(f"错误详情:\n{result1['error_summary']}")
print(f"验证: {'✅ 正确' if result1['success'] else '❌ 错误 - 应该通过但失败了'}")

# Test 2: 错误的属性名 - 应该 FAIL
print("\n[测试 2] 错误的属性名 - 期望 FAIL")
print("-"*80)
wrong_property = """
/// <reference types="office-js-preview" />
async function test2() {
    await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Test", Word.InsertLocation.start);
        p.nonExistentProperty = "this should fail";
        await context.sync();
    });
}
"""
result2 = check_officejs_ts_fast(wrong_property, use_preview=True)
print(f"结果: {'✅ PASS' if result2['success'] else '❌ FAIL'}")
print(f"错误数: {result2['error_count']}")
if not result2['success']:
    print(f"错误详情:\n{result2['error_summary']}")
print(f"验证: {'✅ 正确 - 成功检测到错误' if not result2['success'] else '❌ 错误 - 应该失败但通过了'}")

# Test 3: 错误的方法调用 - 应该 FAIL
print("\n[测试 3] 不存在的方法 - 期望 FAIL")
print("-"*80)
wrong_method = """
/// <reference types="office-js-preview" />
async function test3() {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.nonExistentMethod();
        await context.sync();
    });
}
"""
result3 = check_officejs_ts_fast(wrong_method, use_preview=True)
print(f"结果: {'✅ PASS' if result3['success'] else '❌ FAIL'}")
print(f"错误数: {result3['error_count']}")
if not result3['success']:
    print(f"错误详情:\n{result3['error_summary']}")
print(f"验证: {'✅ 正确 - 成功检测到错误' if not result3['success'] else '❌ 错误 - 应该失败但通过了'}")

# Test 4: 错误的参数类型 - 应该 FAIL
print("\n[测试 4] 错误的参数类型 - 期望 FAIL")
print("-"*80)
wrong_type = """
/// <reference types="office-js-preview" />
async function test4() {
    await Word.run(async (context) => {
        const p = context.document.body.insertParagraph(123, Word.InsertLocation.start);
        await context.sync();
    });
}
"""
result4 = check_officejs_ts_fast(wrong_type, use_preview=True)
print(f"结果: {'✅ PASS' if result4['success'] else '❌ FAIL'}")
print(f"错误数: {result4['error_count']}")
if not result4['success']:
    print(f"错误详情:\n{result4['error_summary']}")
print(f"验证: {'✅ 正确 - 成功检测到类型错误' if not result4['success'] else '❌ 错误 - 应该失败但通过了'}")

# Test 5: 正确的表格操作 - 应该 PASS
print("\n[测试 5] 正确的表格操作 - 期望 PASS")
print("-"*80)
table_code = """
/// <reference types="office-js-preview" />
async function test5() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const table = body.insertTable(3, 4, Word.InsertLocation.end, [
            ["H1", "H2", "H3", "H4"],
            ["C1", "C2", "C3", "C4"],
            ["C5", "C6", "C7", "C8"]
        ]);
        table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
        await context.sync();
    });
}
"""
result5 = check_officejs_ts_fast(table_code, use_preview=True)
print(f"结果: {'✅ PASS' if result5['success'] else '❌ FAIL'}")
print(f"错误数: {result5['error_count']}")
if not result5['success']:
    print(f"错误详情:\n{result5['error_summary']}")
print(f"验证: {'✅ 正确' if result5['success'] else '❌ 错误 - 应该通过但失败了'}")

# Test 6: 测试 Bookmark API (预览版) - 检查是否真的不存在
print("\n[测试 6] Bookmark API 检查 - document.getBookmarkByName")
print("-"*80)
bookmark_code = """
/// <reference types="office-js-preview" />
async function test6() {
    await Word.run(async (context) => {
        const bookmark = context.document.getBookmarkByName("Test");
        bookmark.load("name");
        await context.sync();
    });
}
"""
result6 = check_officejs_ts_fast(bookmark_code, use_preview=True)
print(f"结果: {'✅ PASS' if result6['success'] else '❌ FAIL'}")
print(f"错误数: {result6['error_count']}")
if not result6['success']:
    print(f"错误详情:\n{result6['error_summary']}")
    print(f"说明: ⚠️  'getBookmarkByName' 方法确实不存在于 Word.Document")

# Test 7: 测试正确的 Bookmark 访问方式
print("\n[测试 7] 尝试正确的 Bookmark 访问方式")
print("-"*80)
bookmark_correct = """
/// <reference types="office-js-preview" />
async function test7() {
    await Word.run(async (context) => {
        const bookmarks = context.document.body.getBookmarks();
        bookmarks.load("items");
        await context.sync();
        console.log(bookmarks.items.length);
    });
}
"""
result7 = check_officejs_ts_fast(bookmark_correct, use_preview=True)
print(f"结果: {'✅ PASS' if result7['success'] else '❌ FAIL'}")
print(f"错误数: {result7['error_count']}")
if not result7['success']:
    print(f"错误详情:\n{result7['error_summary']}")

# Summary
print("\n" + "="*80)
print("验证总结")
print("="*80)

tests = [
    ("正确的代码应该通过", result1['success'], True),
    ("错误的属性应该失败", not result2['success'], True),
    ("不存在的方法应该失败", not result3['success'], True),
    ("错误的类型应该失败", not result4['success'], True),
    ("正确的表格操作应该通过", result5['success'], True),
]

passed = sum(1 for _, actual, expected in tests if actual == expected)
total = len(tests)

for name, actual, expected in tests:
    status = "✅" if actual == expected else "❌"
    print(f"{status} {name}")

print(f"\n验证通过: {passed}/{total}")

if passed == total:
    print("\n🎉 所有验证通过! parser_optimized.py 工作正常!")
else:
    print(f"\n⚠️  有 {total - passed} 个验证失败,需要检查!")

# Cleanup
cleanup_persistent_env()
