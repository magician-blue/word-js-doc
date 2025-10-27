import json
import os
import matplotlib.pyplot as plt
from collections import defaultdict

def analyze_examples():
    processed_dir = 'processed'

    # Initialize counters
    properties_examples_count = 0
    methods_examples_count = 0
    kind_examples_count = defaultdict(int)

    # Get all JSON files
    json_files = [f for f in os.listdir(processed_dir) if f.endswith('.json')]

    print(f"正在分析 {len(json_files)} 个文件...")

    for filename in json_files:
        filepath = os.path.join(processed_dir, filename)

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Count examples in properties
            if 'properties' in data:
                for prop in data['properties']:
                    if 'examples' in prop and isinstance(prop['examples'], list):
                        properties_examples_count += len(prop['examples'])

            # Count examples in methods
            if 'methods' in data:
                for method in data['methods']:
                    if 'examples' in method and isinstance(method['examples'], list):
                        example_count = len(method['examples'])
                        methods_examples_count += example_count

                        # Count by kind
                        kind = method.get('kind', 'null')
                        if kind is None:
                            kind = 'null'
                        kind_examples_count[kind] += example_count

        except Exception as e:
            print(f"处理文件 {filename} 时出错: {e}")

    # Print statistics
    print("\n" + "="*50)
    print("统计结果:")
    print("="*50)
    print(f"Properties 提供的 example 数量: {properties_examples_count}")
    print(f"Methods 提供的 example 数量: {methods_examples_count}")
    print(f"\n各个 kind 的 example 数量:")
    for kind, count in sorted(kind_examples_count.items(), key=lambda x: x[1], reverse=True):
        print(f"  {kind}: {count}")

    # Prepare data for pie chart
    total_examples = properties_examples_count + methods_examples_count

    # Create pie chart data
    labels = []
    sizes = []

    # Add properties and methods examples
    if properties_examples_count > 0:
        labels.append('Properties')
        sizes.append(properties_examples_count)

    if methods_examples_count > 0:
        labels.append('Methods')
        sizes.append(methods_examples_count)

    # Create figure with two subplots
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))

    # First pie chart: Properties vs Methods
    colors1 = ['#ff9999', '#66b3ff']
    explode1 = [0.05] * len(sizes)

    ax1.pie(sizes, explode=explode1, labels=labels, colors=colors1, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.set_title(f'Properties vs Methods Examples 分布\n(总计: {total_examples})',
                  fontsize=14, fontweight='bold', pad=20)

    # Second pie chart: Examples by Kind
    kind_labels = []
    kind_sizes = []
    for kind, count in sorted(kind_examples_count.items(), key=lambda x: x[1], reverse=True):
        kind_labels.append(f'{kind}\n({count})')
        kind_sizes.append(count)

    colors2 = plt.cm.Set3(range(len(kind_sizes)))
    explode2 = [0.05] * len(kind_sizes)

    ax2.pie(kind_sizes, explode=explode2, labels=kind_labels, colors=colors2,
            autopct='%1.1f%%', shadow=True, startangle=90)
    ax2.set_title(f'各个 Kind 的 Examples 分布\n(总计: {methods_examples_count})',
                  fontsize=14, fontweight='bold', pad=20)

    plt.tight_layout()
    plt.savefig('examples_distribution.png', dpi=300, bbox_inches='tight')
    print(f"\n扇形图已保存为: examples_distribution.png")

    return {
        'properties_examples': properties_examples_count,
        'methods_examples': methods_examples_count,
        'kind_examples': dict(kind_examples_count),
        'total_examples': total_examples
    }

if __name__ == '__main__':
    # Set Chinese font for matplotlib
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False

    results = analyze_examples()
