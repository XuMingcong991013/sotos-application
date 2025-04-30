import streamlit as st
import os
import pandas as pd
import tempfile
import shutil
import io
import gc
import time
import atexit
import weakref
from contextlib import contextmanager


# 导入安道拓处理模块
from processor.andaotuo_processor import process_andao_files

st.set_page_config(
    page_title="工时统计工具", 
    page_icon="📝", 
    layout="wide"
)

st.title("📝工时统计工具")

st.markdown("""
## 使用说明
1. 从下拉菜单中选择待处理的工时项目
2. 上传Excel文件
3. 点击"开始处理"按钮
""")

# 创建下拉菜单
project_option = st.selectbox(
    "待处理工时项目",
    ["请选择项目", "安道拓（重庆/合肥）", "test"]
)

st.write("请上传工时Excel文件，文件命名按照“员工姓名工时表-公司名-年月”的格式，如“张三工时表-安道拓（重庆）-2025年3月.xlsx”")



# 用于跟踪和清理临时目录
temp_dirs = []

def cleanup_temp_dir(path):
    """尝试清理临时目录的函数，带有重试机制"""
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            # 尝试先删除所有文件
            for root, dirs, files in os.walk(path):
                for file in files:
                    try:
                        file_path = os.path.join(root, file)
                        os.chmod(file_path, 0o777)  # 赋予所有权限
                        os.unlink(file_path)
                    except Exception:
                        pass
            
            # 然后删除目录树
            shutil.rmtree(path, ignore_errors=True)
            return True
        except Exception as e:
            if attempt < max_attempts - 1:
                time.sleep(1)  # 等待1秒再尝试
            else:
                print(f"无法删除临时目录 {path}: {e}")
                return False

# 注册退出处理函数
def cleanup_all_temp_dirs():
    """程序退出时清理所有临时目录"""
    global temp_dirs
    for temp_dir in temp_dirs:
        cleanup_temp_dir(temp_dir)
    temp_dirs = []

atexit.register(cleanup_all_temp_dirs)

@contextmanager
def safe_temp_dir():
    """创建安全的临时目录，确保资源被正确清理"""
    temp_dir = tempfile.mkdtemp()
    global temp_dirs
    temp_dirs.append(temp_dir)
    try:
        yield temp_dir
    finally:
        # 尝试清理，但如果失败也不要抛出异常
        try:
            # 强制垃圾回收
            gc.collect()
            time.sleep(0.5)
            cleanup_temp_dir(temp_dir)
            if temp_dir in temp_dirs:
                temp_dirs.remove(temp_dir)
        except Exception as e:
            st.error(f"清理临时文件时出错: {str(e)}")



# 创建文件上传区域
uploaded_files = st.file_uploader("上传工时文件(支持多文件上传)", 
                                 type=["xlsx"], 
                                 accept_multiple_files=True)

if uploaded_files:
    # 显示上传的文件
    st.write(f"已上传 {len(uploaded_files)} 个文件:")
    for uploaded_file in uploaded_files:
        st.write(f"- {uploaded_file.name}")
    
    # 根据项目选项显示不同的处理逻辑
    if project_option == "安道拓（重庆/合肥）":
        if st.button("开始处理"):
            # 使用安全的临时目录管理器
            with safe_temp_dir() as temp_dir:
                try:
                    # 将上传的文件保存到临时目录
                    for uploaded_file in uploaded_files:
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                    
                    # 确保所有文件已经关闭并且不再被使用
                    uploaded_file = None
                    gc.collect()
                    
                    # 创建进度条
                    progress_bar = st.progress(0)
                    # 创建日志区域
                    log_area = st.empty()
                    
                    # 定义状态更新回调函数
                    def update_status(progress, message):
                        progress_bar.progress(progress)
                        log_area.text_area("处理日志", message, height=400)
                    
                    # 执行处理 - 调用安道拓处理模块
                    with st.spinner('正在处理文件...'):
                        output_file, result_df, error_df = process_andao_files(temp_dir, update_status)
                    
                    # 处理完成后显示结果
                    st.success(f"处理完成！共处理 {len(uploaded_files)} 个文件")
                    
                    # 显示结果表格
                    if not result_df.empty:
                        st.subheader("统计汇总结果")
                        st.dataframe(result_df)
                        
                        # 创建Excel二进制数据用于下载
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            result_df.to_excel(writer, sheet_name='统计汇总', index=False)
                            error_df.to_excel(writer, sheet_name='错误清单', index=False)
                        
                        # 设置文件下载链接
                        st.download_button(
                            label="下载Excel结果文件",
                            data=output_buffer.getvalue(),
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # 如果有错误人员，显示错误清单
                        if not error_df.empty:
                            st.subheader("统计出错人员")
                            st.dataframe(error_df)
                    else:
                        st.warning("未生成任何结果，请检查上传的文件是否正确")
                    
                    # 清理可能的引用
                    result_df = None
                    error_df = None
                    gc.collect()
                    
                except Exception as e:
                    st.error(f"处理过程中发生错误: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
    
    elif project_option == "test":
        st.info("Test项目的处理逻辑尚未实现，此为占位项目。")
else:
    if project_option == "安道拓（重庆/合肥）":
        st.info("请上传Excel文件开始处理。")
    elif project_option == "test":
        st.info("Test项目的处理逻辑尚未实现，此为占位项目。请上传文件后再试。")

# 添加页脚信息
st.markdown("---")
st.markdown("© 2025 工时统计工具 | develped by Mingcong Xu | Version 1.0")