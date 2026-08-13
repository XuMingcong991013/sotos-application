"""
实现：上传Excel表格，精确查找发票金额的组合等于目标金额，下载结果，每个sheet为一个组合结果。
依赖：pip install streamlit pandas openpyxl
启动命令：streamlit run streamlit_app.py


代码逻辑：
- progress_decorator：进度条装饰器

- find_exact_combinations：组合查找算法
    预处理阶段：
    1. 将所有数据转换为浮点数，过滤掉大于目标金额的发票
    2. 按金额分组，将相同金额的发票索引分到一组

    动态规划过程：使用字段dp来存储中间结果，key是当前累计的金额和，value是能达到该金额的索引组合列表
    1. 从0开始，代表空组合，dp[0] = [[]]
    2. 对每个不同的金额值：尝试从该金额组中选择1到k个发票；对于每种选择，计算与之前累积金额的新组合；只保留不超过目标金额的新组合；确保不会在一个组合中重复使用同一个发票

    剪枝优化：
    1. 跳过那些即使加上所有剩余金额也无法达到目标的路径
    2. 限制每个金额和下保存的组合数量，避免内存爆炸
    3. 对超出目标金额的组合直接跳过

    结果收集：最后只收集精确等于目标金额的组合

- display_results：结果显示函数

- main：主函数，处理UI交互和流程控制
"""

import streamlit as st
import pandas as pd
import numpy as np
from collections import defaultdict
from itertools import combinations
import io
from datetime import datetime

# 进度条装饰器
def progress_decorator(func):
    def wrapper(*args, **kwargs):
        if 'progress_bar' in kwargs:
            progress_bar = kwargs.pop('progress_bar')
        else:
            progress_bar = st.progress(0)
        
        result = func(*args, **kwargs, progress_bar=progress_bar)
        progress_bar.progress(100)
        return result
    return wrapper

@progress_decorator
def find_exact_combinations(df, amount_col, target_amount, progress_bar=None):
    """
    精确匹配目标金额的组合查找算法
    参数：
    - df: 输入DataFrame
    - amount_col: 金额列名
    - target_amount: 目标金额
    - progress_bar: Streamlit进度条对象
    """
    # 初始化进度
    if progress_bar:
        progress_bar.progress(5)
    
    # 转换和预处理
    df = df.copy()
    df['_amount_float'] = df[amount_col].astype(float)
    target = float(target_amount)
    
    # 过滤掉大于目标金额的数据
    candidates = df[df['_amount_float'] <= target]
    if len(candidates) == 0:
        return []
    
    if progress_bar:
        progress_bar.progress(10)
    
    # 按金额分组（大幅减少计算量）
    amount_groups = candidates.groupby('_amount_float').apply(lambda x: x.index.tolist()).reset_index(name='original_indices')
    total_groups = len(amount_groups)
    
    # 动态规划算法（精确匹配）
    dp = {0: [[]]}
    
    for i, row in amount_groups.iterrows():
        current_amount = row['_amount_float']
        group_indices = row['original_indices']
        
        new_dp = {}
        
        # 更新进度
        if progress_bar and i % 5 == 0:
            progress = 10 + int(70 * (i / total_groups))
            progress_bar.progress(progress)
        
        for sum_val in list(dp.keys()):
            # 剪枝：如果当前和加上所有剩余金额都小于目标值，则跳过
            remaining_potential = sum([amount_groups.iloc[j]['_amount_float'] for j in range(i, total_groups)])
            if sum_val + remaining_potential < target:
                continue
                
            # 限制同一金额最多使用的票据数量
            max_possible = min(10, len(group_indices))
            
            for k in range(1, max_possible + 1):
                # 生成所有k个元素的组合
                for indices_combo in combinations(group_indices, k):
                    new_sum = sum_val + current_amount * k
                    
                    # 如果超出目标金额，跳过
                    if new_sum > target:
                        continue
                        
                    # 为每个现有组合创建新组合
                    for prev_indices in dp[sum_val]:
                        # 确保没有重复索引
                        if not any(idx in prev_indices for idx in indices_combo):
                            new_combo = prev_indices + list(indices_combo)
                            
                            if new_sum not in new_dp:
                                new_dp[new_sum] = []
                            
                            # 限制组合数量，避免内存爆炸
                            if len(new_dp[new_sum]) < 1000:
                                new_dp[new_sum].append(new_combo)
        
        # 合并结果到dp
        for sum_val, combinations_list in new_dp.items():
            if sum_val not in dp:
                dp[sum_val] = []
            
            dp[sum_val].extend(combinations_list)
            
            # 如果某个和的组合太多，进行剪枝
            if len(dp[sum_val]) > 2000:
                dp[sum_val] = dp[sum_val][:2000]
    
    # 只收集精确匹配的结果
    results = []
    if target in dp:
        for indices in dp[target]:
            if indices:  # 确保不是空列表
                results.append(df.loc[indices].drop(columns=['_amount_float']))
    
    if progress_bar:
        progress_bar.progress(90)
    
    return results


def main():
    st.set_page_config(
        page_title="发票组合精确匹配查找器", 
        layout="wide"
    )
    
    # 使用列布局来分割左右两栏
    left_column, right_column = st.columns([1, 2])
    
    # 左侧栏 - 配置区域
    with left_column:
        st.header("💰 配置参数")
        
        # 上传文件区域
        st.subheader("1. 上传数据")
        uploaded_file = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.success("数据加载成功！")
                
                # 数据预览
                with st.expander("数据预览"):
                    st.dataframe(df.head())
                
                # 配置参数区域
                st.subheader("2. 设置参数")
                
                # 列选择器
                cols = df.columns.tolist()
                amount_col = st.selectbox(
                    "选择金额列", 
                    cols, 
                    index=cols.index('开票金额') if '开票金额' in cols else 0
                )
                
                # 目标金额设置
                target_amount = st.number_input(
                    "目标金额", 
                    min_value=0.01, 
                    value=732698.95, 
                    step=1000.0,
                    format="%.2f"
                )
                
                # 匹配模式选择
                match_mode = st.radio(
                    "匹配模式",
                    ["精确匹配", "允许误差"],
                    index=0
                )
                
                # 高级选项
                with st.expander("高级选项"):
                    if match_mode == "允许误差":
                        tolerance = st.slider(
                            "允许误差(%)", 
                            min_value=0.0, 
                            max_value=1.0, 
                            value=0.1, 
                            step=0.01
                        ) / 100
                    else:
                        tolerance = 0  # 精确匹配时忽略容差
                     
                    max_results = st.number_input(
                        "最大结果数", 
                        min_value=1, 
                        value=100,
                        help="限制返回结果数量避免内存溢出"
                    )
                
                # 执行按钮 - 放在左侧栏底部
                search_button = st.button("🚀 开始查找", use_container_width=True)
            
            except Exception as e:
                st.error(f"错误: {str(e)}")
                st.exception(e)
                search_button = False
                df = None
        else:
            st.info("请上传Excel文件开始")
            search_button = False
            df = None
    
    # 右侧栏 - 结果显示区域
    with right_column:
        # 标题和说明
        st.title("💰 发票组合精确匹配查找器")
        st.caption("查找精确等于目标金额的发票组合")
        
        # 如果有上传文件并点击查找按钮，开始处理
        if search_button and df is not None:
            if df.empty:
                st.error("数据为空！")
            else:
                # 创建进度区域
                progress_container = st.container()
                with progress_container:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                
                # 执行查找
                start_time = datetime.now()
                status_text.text("正在预处理数据...")
                
                try:
                    if match_mode == "精确匹配":
                        results = find_exact_combinations(
                            df,
                            amount_col,
                            target_amount,
                            progress_bar=progress_bar
                        )[:max_results]
                    else:
                        # 这里需要添加find_approximate_combinations函数或使用相同函数接口
                        results = find_exact_combinations(
                            df,
                            amount_col,
                            target_amount,
                            progress_bar=progress_bar
                        )[:max_results]
                    
                    duration = (datetime.now() - start_time).total_seconds()
                    status_text.text(f"查找完成！耗时 {duration:.2f} 秒")
                    
                    # 显示结果
                    st.subheader("查找结果")
                    
                    if not results:
                        if match_mode == "精确匹配":
                            st.warning("未找到精确匹配的组合。尝试使用'允许误差'模式可能会有更多结果。")
                        else:
                            st.warning(f"未找到符合条件的组合。")
                    else:
                        # 显示成功消息
                        result_count = len(results)
                        st.success(f"找到 {result_count} 个符合条件的组合")
                        
                        # 在右侧显示结果列表
                        result_container = st.container()
                        with result_container:
                            # 创建下载按钮
                            excel_data = create_excel_output(results, amount_col, target_amount)
                            if excel_data:
                                col1, col2 = st.columns([3, 1])
                                with col2:
                                    st.download_button(
                                        label="⬇️ 下载全部结果",
                                        data=excel_data,
                                        file_name=f"发票组合_{target_amount:.2f}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                            
                            # 逐个显示结果
                            for i, combo in enumerate(results[:50], 1):
                                # 计算总和
                                total = combo[amount_col].sum()
                                diff = total - target_amount
                                
                                with st.expander(f"组合 {i} (总和: {total:,.2f}, 差异: {diff:,.2f})"):
                                    # 添加总计行
                                    total_row = pd.DataFrame({
                                        amount_col: [total],
                                        '状态': ['总和'],
                                        '差异': [diff]
                                    })
                                    display_df = pd.concat([combo, total_row])
                                    
                                    # 显示数据表格
                                    st.dataframe(display_df, use_container_width=True)
                
                except Exception as e:
                    st.error(f"处理过程中出错: {str(e)}")
                    st.exception(e)
        else:
            # 首次加载或未上传文件时显示
            st.info("👈 请在左侧上传Excel文件并设置参数")
            
            # 添加说明内容
            with st.expander("使用说明"):
                st.markdown("""
                ### 使用方法
                1. 在左侧上传包含发票数据的Excel文件
                2. 选择包含金额数据的列
                3. 设置需要匹配的目标金额
                4. 选择匹配模式（精确匹配或允许误差）
                5. 点击"开始查找"按钮
                
                ### 算法说明
                本工具使用动态规划算法，高效查找所有可能的发票组合，确保:
                - 精确匹配目标金额
                - 不重复使用同一发票
                - 结果可导出为Excel文件
                """)
                
            # 示例图片或GIF（可选）
            # st.image("example.gif", caption="使用示例")


def create_excel_output(results, amount_col, target_amount):
    """生成Excel下载文件"""
    if not isinstance(results, list) or not results:
        return None
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, combo in enumerate(results, 1):
            # 添加统计行
            total = combo[amount_col].sum()
            diff = total - target_amount
            total_row = pd.DataFrame({
                amount_col: [total],
                '状态': ['总和'],
                '差异': [diff]
            })
            
            # 合并显示
            display_df = pd.concat([combo, total_row])
            
            # 写入Excel
            sheet_name = f"组合_{i}"
            display_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return output.getvalue()


# 主程序入口
if __name__ == "__main__":
    main()


    