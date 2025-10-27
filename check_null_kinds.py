import json
import os
from collections import defaultdict

def find_null_kind_methods():
    processed_dir = 'processed'
    json_files = [f for f in os.listdir(processed_dir) if f.endswith('.json')]

    null_kind_methods = []

    for filename in json_files:
        filepath = os.path.join(processed_dir, filename)

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if 'methods' in data:
                for method in data['methods']:
                    kind = method.get('kind')
                    if kind is None or kind == 'null':
                        null_kind_methods.append({
                            'file': filename,
                            'class': data['class']['name'],
                            'method': method['name'],
                            'description': method.get('description', '')[:100]
                        })
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    print(f"找到 {len(null_kind_methods)} 个 kind 为 null 的方法\n")
    print("示例方法 (前 20 个):")
    print("=" * 100)

    for i, method_info in enumerate(null_kind_methods[:20]):
        print(f"{i+1}. 类: {method_info['class']}")
        print(f"   方法: {method_info['method']}")
        print(f"   描述: {method_info['description']}")
        print()

if __name__ == '__main__':
    find_null_kind_methods()
