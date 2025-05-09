import streamlit as st
import os
import pandas as pd
import tempfile
import shutil
import io
import gc
import time
import atexit
from datetime import datetime, timedelta
import re


from processor.andaotuo_processor import extract_info_from_filename, excel_num_to_date, read_data_from_excel, process_single_file

# 设置页面配置
st.set_page_config(
    page_title="工时统计工具", 
    page_icon="📝", 
    layout="wide"
)

st.title("📝工时统计工具")

st.markdown("""
## 使用说明
1. 选择待处理项目名称
2. 上传Excel文件(支持多文件批量上传，文件命名按照"员工姓名工时表-公司名-年月"的格式)
3. 点击"开始处理"按钮
""")

# 项目选择
project = st.selectbox(
    "选择项目点",
    options=["安道拓（重庆/合肥）", "test1", "test2"],
    index=0
)

# 临时文件管理
temp_dirs = []

def cleanup_temp_dir(path):
    """清理临时目录"""
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception as e:
        print(f"无法删除临时目录 {path}: {e}")

def cleanup_all_temp_dirs():
    """程序退出时清理所有临时目录"""
    global temp_dirs
    for temp_dir in temp_dirs:
        cleanup_temp_dir(temp_dir)
    temp_dirs = []

atexit.register(cleanup_all_temp_dirs)

def process_files(uploaded_files, update_status, project):
    """处理上传的所有文件"""
    all_results = []
    wrong_employees = []
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    temp_dirs.append(temp_dir)
    
    try:
        # 保存上传的文件到临时目录
        file_paths = []
        for i, uploaded_file in enumerate(uploaded_files):
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            file_paths.append(file_path)
            
            # 更新进度
            update_status((i + 1) / len(uploaded_files), 
                         f"已保存文件: {uploaded_file.name}")
        
        # 根据项目选择处理逻辑
        processor_module = (
            "processor.andaotuo_processor" if "安道拓" in project
            else f"processor.{project}_processor"
        )
        
        # 处理每个文件
        for i, file_path in enumerate(file_paths):
            update_status((i + 1) / len(file_paths), 
                         f"正在处理文件: {os.path.basename(file_path)}")
            
            try:
                if "安道拓" in project:
                    from processor.andaotuo_processor import process_single_file
                else:
                    # 动态导入其他处理模块
                    module = __import__(processor_module, fromlist=['process_single_file'])
                    process_single_file = getattr(module, 'process_single_file')
                
                df_result, wrong_employees = process_single_file(file_path, wrong_employees)
                all_results.append(df_result)
            except Exception as e:
                update_status((i + 1) / len(file_paths), 
                             f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                continue
        
        # 合并所有结果
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            
            # 创建错误数据框
            error_df = pd.DataFrame({
                '工程师': wrong_employees,
                '错误类型': '工时明细与汇总数据不一致'
            })
            
            return final_df, error_df
        else:
            return pd.DataFrame(), pd.DataFrame()
            
    finally:
        cleanup_temp_dir(temp_dir)
        if temp_dir in temp_dirs:
            temp_dirs.remove(temp_dir)

# 主界面
uploaded_files = st.file_uploader(
    "上传工时Excel文件(支持多文件上传)", 
    type=["xlsx"], 
    accept_multiple_files=True,
    help="文件命名按照'员工姓名工时表-公司名-年月'的格式，如'张三工时表-安道拓(重庆)-2025年3月.xlsx'"
)

if uploaded_files:
    st.write(f"已上传 {len(uploaded_files)} 个文件:")
    for uploaded_file in uploaded_files:
        st.write(f"- {uploaded_file.name}")
    
    if st.button("开始处理"):
        # 创建进度条和日志区域
        progress_bar = st.progress(0)
        log_area = st.empty()
        
        def update_status(progress, message):
            progress_bar.progress(progress)
            log_area.text_area("处理日志", message, height=200)
        
        with st.spinner('正在处理文件...'):
            df_result, df_error = process_files(uploaded_files, update_status, project)
        
        # 显示处理结果
        st.success("处理完成！")
        
        if not df_result.empty:
            st.subheader("工时统计结果")
            st.dataframe(df_result)
            
            # 提供下载按钮
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name='工时汇总', index=False)
                if not df_error.empty:
                    df_error.to_excel(writer, sheet_name='错误清单', index=False)
            
            st.download_button(
                label="下载Excel结果文件",
                data=output_buffer.getvalue(),
                file_name="工时统计结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 显示错误信息
            if not df_error.empty:
                st.warning(f"发现 {len(df_error)} 个文件的工时明细与汇总数据不一致:")
                st.dataframe(df_error)
        else:
            st.warning("未生成有效结果，请检查上传的文件格式是否正确")

# 页脚
st.markdown("---")
st.markdown("© 2025 工时统计工具 | Developed by Mingcong Xu | Version 1.0")