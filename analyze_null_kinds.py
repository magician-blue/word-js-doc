import json
import os
from collections import Counter

def analyze_null_kinds():
    processed_dir = 'processed'
    json_files = [f for f in os.listdir(processed_dir) if f.endswith('.json')]

    null_kind_methods = []
    method_names = []

    for filename in json_files:
        filepath = os.path.join(processed_dir, filename)

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if 'methods' in data:
                for method in data['methods']:
                    kind = method.get('kind')
                    if kind is None:
                        method_names.append(method['name'])
                        null_kind_methods.append({
                            'name': method['name'],
                            'description': method.get('description', '')
                        })
        except Exception as e:
            pass

    # Count method name frequency
    name_counter = Counter(method_names)

    print("=" * 80)
    print("Kind 为 null 的方法分析")
    print("=" * 80)
    print(f"\n总共有 {len(null_kind_methods)} 个方法的 kind 为 null")
    print(f"涉及 {len(name_counter)} 个不同的方法名\n")

    print("最常见的方法名 (Top 15):")
    print("-" * 80)
    for name, count in name_counter.most_common(15):
        print(f"  {name:40s} 出现 {count:3d} 次")

    print("\n\n这些方法的特点:")
    print("-" * 80)

    # Categorize by function type
    action_verbs = {
        'search': '搜索操作',
        'select': '选择/导航操作',
        'save': '保存操作',
        'close': '关闭操作',
        'open': '打开操作',
        'copy': '复制操作',
        'cut': '剪切操作',
        'paste': '粘贴操作',
        'compare': '比较操作',
        'import': '导入操作',
        'export': '导出操作',
        'apply': '应用操作',
        'generate': '生成操作',
        'detect': '检测操作',
        'set': '设置操作'
    }

    categorized = {}
    for name, count in name_counter.items():
        categorized_flag = False
        for verb, category in action_verbs.items():
            if verb in name.lower():
                if category not in categorized:
                    categorized[category] = []
                categorized[category].append((name, count))
                categorized_flag = True
                break
        if not categorized_flag:
            if '其他操作' not in categorized:
                categorized['其他操作'] = []
            categorized['其他操作'].append((name, count))

    for category, methods in sorted(categorized.items()):
        print(f"\n【{category}】")
        for name, count in sorted(methods, key=lambda x: x[1], reverse=True):
            print(f"  - {name} ({count}次)")

if __name__ == '__main__':
    analyze_null_kinds()
