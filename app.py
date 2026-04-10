import pandas as pd
import streamlit as st

# 1. 网页基础设置
st.set_page_config(page_title="2D 红外异常数据排查", page_icon="📋", layout="centered")

st.title("📋 2D 红外数据逻辑排查工具")
st.markdown("上传你的原位红外 Excel 表格，系统将自动进行拓扑逻辑排查，找出所有“自相矛盾”的数据，并输出整体的演变顺序。")

# 2. 网页文件上传模块
uploaded_file = st.file_uploader("请拖拽或点击上传 Excel 表格 (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner('正在执行拓扑逻辑排查，请稍候...'):
        
        # 3. 读取与分析数据
        df = pd.read_excel(uploaded_file, index_col=0)
        common_nodes = set(df.index).intersection(df.columns)

        precedes = {node: set() for node in common_nodes}
        evidence = {}

        # 建立先后顺序关系
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

        # 检测两点矛盾
        len2_cycles = set()
        for u in common_nodes:
            for v in precedes[u]:
                if u in precedes[v] and u < v:
                    len2_cycles.add((u, v))

        # 检测三点矛盾
        len3_cycles = set()
        for u in common_nodes:
            for v in precedes[u]:
                for w in precedes[v]:
                    if u in precedes[w]:
                        if u != v and v != w and u != w:
                            cycle = tuple(sorted([u, v, w]))
                            if cycle not in [tuple(sorted(c)) for c in len3_cycles]:
                                len3_cycles.add((u, v, w))

        total_conflicts = len(len2_cycles) + len(len3_cycles)

        # ==================== 4. 网页结果展示 ====================
        st.markdown("---")
        if total_conflicts == 0:
            st.success("🎉 **恭喜！未检测到任何两点或三点的逻辑矛盾，数据顺序非常完美！**")
        else:
            st.error(f"⚠️ **警告：共检测到 {total_conflicts} 处逻辑矛盾！** 请根据下方的详细表格回 Excel 检查数据。")

            # 展示两点矛盾
            if len2_cycles:
                st.subheader("🔴 两点直接冲突 (相互矛盾)")
                for idx, (u, v) in enumerate(len2_cycles, 1):
                    ev1, ev2 = evidence[(u, v)][0], evidence[(v, u)][0]
                    st.markdown(f"**【矛盾 {idx}】**")
                    
                    conflict_data = pd.DataFrame({
                        "矛盾环节": [f"{u} 先于 {v}", f"{v} 先于 {u}"],
                        "X坐标 (波长)": [ev1['X'], ev2['X']],
                        "Y坐标 (波长)": [ev1['Y'], ev2['Y']],
                        "Excel中的异常值": [ev1['val'], ev2['val']]
                    })
                    st.dataframe(conflict_data, use_container_width=True)

            # 展示三点矛盾
            if len3_cycles:
                st.subheader("🟠 三点循环冲突 (A → B → C → A 闭环)")
                start_idx = len(len2_cycles) + 1
                for idx, (u, v, w) in enumerate(len3_cycles, start_idx):
                    ev1, ev2, ev3 = evidence[(u, v)][0], evidence[(v, w)][0], evidence[(w, u)][0]
                    st.markdown(f"**【矛盾 {idx}】**")
                    
                    conflict_data = pd.DataFrame({
                        "矛盾环节": [f"{u} 先于 {v}", f"{v} 先于 {w}", f"{w} 先于 {u}"],
                        "X坐标 (波长)": [ev1['X'], ev2['X'], ev3['X']],
                        "Y坐标 (波长)": [ev1['Y'], ev2['Y'], ev3['Y']],
                        "Excel中的异常值": [ev1['val'], ev2['val'], ev3['val']]
                    })
                    st.dataframe(conflict_data, use_container_width=True)
        
        # ==================== 5. 新增：输出整体排序与标红 ====================
        st.markdown("---")
        st.subheader("⏱️ 峰位响应时间演变顺序 (由先到后)")
        
        # 1. 揪出所有“内鬼”（参与矛盾的峰位）
        bad_nodes = set()
        for u, v in len2_cycles:
            bad_nodes.update([u, v])
        for u, v, w in len3_cycles:
            bad_nodes.update([u, v, w])
            
        # 2. 计分排序：计算每个节点“先于”其他节点的数量
        scores = {node: len(precedes[node]) for node in common_nodes}
        sorted_nodes = sorted(common_nodes, key=lambda x: scores[x], reverse=True)
        
        # 3. 构造带有颜色的 HTML 输出结果
        sequence_html = "<div style='font-size: 18px; line-height: 2; padding: 10px; background-color: #f0f2f6; border-radius: 8px;'>"
        
        for i, node in enumerate(sorted_nodes):
            if node in bad_nodes:
                sequence_html += f"<span style='color: red; font-weight: bold; padding: 0 4px;'>{node}</span>"
            else:
                sequence_html += f"<span style='color: #333; font-weight: 500; padding: 0 4px;'>{node}</span>"
                
            if i < len(sorted_nodes) - 1:
                sequence_html += "<span style='color: #888;'> ➔ </span>"
                
        sequence_html += "</div>"
        
        # 渲染 HTML
        st.markdown(sequence_html, unsafe_allow_html=True)
        
        # 底部提示语
        if len(bad_nodes) > 0:
            st.caption("💡 **提示：** <span style='color: red;'>红色的峰位</span> 陷入了逻辑死循环，其在此序列中的绝对位置仅供参考，请重点核对红色的数据点。", unsafe_allow_html=True)
        else:
            st.caption("💡 **提示：** 数据完美！未发现逻辑矛盾。")
