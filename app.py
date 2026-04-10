import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import webbrowser  # 新增：用于自动打开网页

# ================= 1. 弹窗选择文件 =================
root = tk.Tk()
root.withdraw() 
root.attributes('-topmost', True) 

print("请选择需要进行【逻辑合理性排查】的 Excel 表格...")
input_path = filedialog.askopenfilename(
    title="请选择需要排查的 Excel 表格",
    filetypes=[("Excel 文件", "*.xlsx *.xls")]
)

if not input_path:
    print("未选择任何文件，程序已取消。")
    exit()

print(f"已成功载入文件：{input_path}")
print("正在执行拓扑逻辑排查，准备生成网页报告...")

# 自动生成输出网页的路径
dir_name = os.path.dirname(input_path)
base_name = os.path.basename(input_path)
name_without_ext, _ = os.path.splitext(base_name)
html_path = os.path.join(dir_name, f"{name_without_ext}_异常数据排查报告.html")

# ================= 2. 读取与构建有向图 =================
df = pd.read_excel(input_path, index_col=0)
common_nodes = set(df.index).intersection(df.columns)

precedes = {node: set() for node in common_nodes}
evidence = {}

for y in common_nodes:
    for x in common_nodes:
        val = df.loc[y, x]
        if pd.isna(val): continue
        try: 
            val = float(val)
        except ValueError: 
            continue
        if val == 0: continue

        if val > 0:
            u, v = x, y 
        elif val < 0:
            u, v = y, x 

        precedes[u].add(v)
        if (u, v) not in evidence:
            evidence[(u, v)] = []
        evidence[(u, v)].append({"X": x, "Y": y, "val": val})

# ================= 3. 检测逻辑矛盾 =================
# 3.1 两点直接矛盾
len2_cycles = set()
for u in common_nodes:
    for v in precedes[u]:
        if u in precedes[v] and u < v:
            len2_cycles.add((u, v))

# 3.2 三点循环矛盾
len3_cycles = set()
for u in common_nodes:
    for v in precedes[u]:
        for w in precedes[v]:
            if u in precedes[w]:
                if u != v and v != w and u != w:
                    cycle = tuple(sorted([u, v, w]))
                    if cycle not in [tuple(sorted(c)) for c in len3_cycles]:
                        len3_cycles.add((u, v, w))

# ================= 4. 生成漂亮的 HTML 网页 =================
total_conflicts = len(len2_cycles) + len(len3_cycles)

html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>异常数据排查报告</title>
    <style>
        body {{ font-family: 'Microsoft YaHei', sans-serif; background-color: #f4f7f6; color: #333; margin: 40px; }}
        h1 {{ color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
        .summary {{ background-color: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 30px; font-size: 16px; }}
        .success {{ color: #27ae60; font-weight: bold; font-size: 18px; text-align: center; }}
        .warning {{ color: #e74c3c; font-weight: bold; font-size: 18px; }}
        .card {{ background-color: #fff; border-left: 5px solid #e74c3c; padding: 15px 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .card-title {{ font-size: 18px; color: #c0392b; font-weight: bold; margin-top: 0; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
        th {{ background-color: #f8f9fa; color: #555; }}
        .highlight {{ color: #e67e22; font-weight: bold; }}
    </style>
</head>
<body>
    <h1>📋 2D 红外数据逻辑排查报告</h1>
    <div class="summary">
        <p><strong>数据源文件：</strong> {base_name}</p>
"""

if total_conflicts == 0:
    html_content += f'<p class="success">🎉 恭喜！未检测到任何两点或三点的逻辑矛盾，数据顺序合理。</p>'
else:
    html_content += f'<p class="warning">⚠️ 警告：共检测到 {total_conflicts} 处逻辑矛盾！请根据下方提示回 Excel 检查数据。</p>'

html_content += "</div>"

# 写入两点矛盾卡片
for idx, (u, v) in enumerate(len2_cycles, 1):
    ev1, ev2 = evidence[(u, v)][0], evidence[(v, u)][0]
    html_content += f"""
    <div class="card">
        <p class="card-title">【矛盾 {idx}】: 两点直接冲突 (相互矛盾)</p>
        <table>
            <tr><th>矛盾环节</th><th>X坐标 (波长)</th><th>Y坐标 (波长)</th><th>Excel中的数值</th></tr>
            <tr><td><span class="highlight">{u} 先于 {v}</span></td><td>{ev1['X']}</td><td>{ev1['Y']}</td><td>{ev1['val']}</td></tr>
            <tr><td><span class="highlight">{v} 先于 {u}</span></td><td>{ev2['X']}</td><td>{ev2['Y']}</td><td>{ev2['val']}</td></tr>
        </table>
    </div>
    """

# 写入三点矛盾卡片
start_idx = len(len2_cycles) + 1
for idx, (u, v, w) in enumerate(len3_cycles, start_idx):
    ev1, ev2, ev3 = evidence[(u, v)][0], evidence[(v, w)][0], evidence[(w, u)][0]
    html_content += f"""
    <div class="card" style="border-left-color: #f39c12;">
        <p class="card-title" style="color: #d35400;">【矛盾 {idx}】: 三点循环冲突 (A &rarr; B &rarr; C &rarr; A 闭环)</p>
        <table>
            <tr><th>矛盾环节</th><th>X坐标 (波长)</th><th>Y坐标 (波长)</th><th>Excel中的数值</th></tr>
            <tr><td><span class="highlight">{u} 先于 {v}</span></td><td>{ev1['X']}</td><td>{ev1['Y']}</td><td>{ev1['val']}</td></tr>
            <tr><td><span class="highlight">{v} 先于 {w}</span></td><td>{ev2['X']}</td><td>{ev2['Y']}</td><td>{ev2['val']}</td></tr>
            <tr><td><span class="highlight">{w} 先于 {u}</span></td><td>{ev3['X']}</td><td>{ev3['Y']}</td><td>{ev3['val']}</td></tr>
        </table>
    </div>
    """

html_content += "</body></html>"

# 保存为 HTML 文件
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n排查完毕！共发现 {total_conflicts} 处矛盾。")
print("正在自动打开网页报告...")

# 自动在默认浏览器中打开生成的网页
webbrowser.open(f'file://{os.path.abspath(html_path)}')