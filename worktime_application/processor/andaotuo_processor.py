import os
import re
import pandas as pd
import numpy as np
from datetime import datetime
import tempfile
import shutil
import time

def process_andao_files(folder_path, status_update_callback=None):
    """
    处理安道拓（重庆/合肥）项目的工时文件
    
    Parameters:
    -----------
    folder_path : str
        包含Excel工时文件的文件夹路径
    status_update_callback : function, optional
        状态更新回调函数，用于更新进度和日志
        
    Returns:
    --------
    tuple
        (output_file_name, result_dataframe, error_dataframe)
    """
    # 存储所有统计结果
    all_results = []
    # 存储统计出错的员工
    wrong_employees = []
    
    # 获取文件夹中的所有Excel文件，排除临时文件（以~$开头）
    excel_files = [f for f in os.listdir(folder_path) 
                  if f.endswith('.xlsx') and ('工时' in f or '工时表' in f) and not f.startswith('~$')]
    
    total_files = len(excel_files)
    
    for file_index, file_name in enumerate(excel_files):
        try:
            # 更新进度
            if status_update_callback:
                progress = (file_index) / total_files
                status_update_callback(progress, f"处理文件 {file_index+1}/{total_files}: {file_name}")
            
            # 1. 从文件名提取信息，修正正则表达式
            match = re.match(r'(.+)工时表?-(.+)-(\d{4})年(\d{1,2})月', file_name)
            if match:
                employee_name = match.group(1)
                company_name = match.group(2)
                work_year = match.group(3)  # 直接提取4位数年份
                work_month = match.group(4).lstrip('0')  # 去除前导零
                work_year_month = work_year + str(work_month).zfill(2)
                
                log_message = f"\n========== 处理文件: {file_name} ==========\n"
                log_message += f"提取信息: 员工={employee_name}, 公司={company_name}, 年={work_year}, 月={work_month}"
                
                # 2. 读取相应的sheet
                file_path = os.path.join(folder_path, file_name)
                
                # 指定Excel引擎，避免格式识别问题
                xl = pd.ExcelFile(file_path, engine='openpyxl')
                log_message += f"\n文件中的sheet名称: {xl.sheet_names}"
                
                # 尝试找到正确的sheet名
                correct_sheet = None
                for sheet_name in xl.sheet_names:
                    if employee_name in sheet_name or sheet_name in employee_name:
                        correct_sheet = sheet_name
                        break
                
                if correct_sheet is None and employee_name in xl.sheet_names:
                    correct_sheet = employee_name
                
                if correct_sheet is None:
                    # 如果找不到匹配的sheet，尝试使用第一个sheet
                    if len(xl.sheet_names) > 0:
                        correct_sheet = xl.sheet_names[0]
                        log_message += f"\n警告: 未找到匹配'{employee_name}'的sheet，使用第一个sheet: '{correct_sheet}'"
                    else:
                        log_message += f"\n错误: 文件 {file_name} 中没有可用的sheet"
                        if status_update_callback:
                            status_update_callback(progress, log_message)
                        continue
                else:
                    log_message += f"\n使用sheet: '{correct_sheet}'"
                
                # 3. 根据月份确定读取范围
                work_month_int = int(work_month)
                work_year_int = int(work_year)
                
                if work_month_int in [1, 3, 5, 7, 8, 10, 12]:
                    end_row = 35
                elif work_month_int in [4, 6, 9, 11]:
                    end_row = 34
                elif work_month_int == 2:
                    if work_year_int % 4 == 0:
                        end_row = 33
                    else:
                        end_row = 32
                
                log_message += f"\n读取工时明细范围: B4:F{end_row}"
                
                if status_update_callback:
                    status_update_callback(progress, log_message)
                
                try:
                    # 读取整个工作表，以便于全面分析
                    full_sheet = pd.read_excel(file_path, sheet_name=correct_sheet, engine='openpyxl')
                    log_message += f"\n工作表总行数: {len(full_sheet)}"
                    
                    # 读取第4行作为表头
                    header_df = pd.read_excel(file_path, sheet_name=correct_sheet, 
                                             usecols="B:F", header=3, nrows=0, engine='openpyxl')
                    header_names = header_df.columns.tolist()
                    log_message += f"\n读取到的表头: {header_names}"
                    
                    # 根据读取到的表头确定列名映射
                    column_mapping = {}
                    for i, col in enumerate(header_names):
                        col_lower = str(col).lower()
                        if '项目' in col_lower and '名称' in col_lower:
                            column_mapping[col] = '项目名称'
                        elif '开始' in col_lower and ('时间' in col_lower or '工作' in col_lower):
                            column_mapping[col] = '开始工作时间'
                        elif '结束' in col_lower and ('时间' in col_lower or '工作' in col_lower):
                            column_mapping[col] = '工作结束时间'
                        elif '加班' in col_lower:
                            column_mapping[col] = '加班时'
                        elif '总' in col_lower and '工作' in col_lower:
                            column_mapping[col] = '总工作时'
                        else:
                            # 使用默认列名
                            column_mapping[col] = ['项目名称', '开始工作时间', '工作结束时间', '加班时', '总工作时'][i]
                    
                    log_message += f"\n列名映射: {column_mapping}"
                    
                    # 读取详细数据，从第5行开始(跳过前4行)
                    df_detail = pd.read_excel(file_path, sheet_name=correct_sheet, 
                                             usecols="B:F", header=3, nrows=end_row-4, engine='openpyxl')
                    
                    # 重命名列
                    df_detail = df_detail.rename(columns=column_mapping)
                    
                    # 显示工时明细的前几行
                    log_message += "\n\n工时明细数据(df_detail)前5行:"
                    log_message += f"\n{df_detail.head(5)}"
                    
                    # 去除空行，只保留项目名称不为空的行
                    df_detail = df_detail.dropna(subset=['项目名称'])
                    
                    # 确保数值列是数值类型
                    df_detail['加班时'] = pd.to_numeric(df_detail['加班时'], errors='coerce').fillna(0)
                    df_detail['总工作时'] = pd.to_numeric(df_detail['总工作时'], errors='coerce').fillna(0)
                    
                    # 更宽松的过滤规则：只过滤纯数字项目名称，允许包含字母和数字的混合项目名称
                    df_detail_filtered = df_detail.copy()
                    # 转换项目名称为字符串
                    df_detail_filtered['项目名称'] = df_detail_filtered['项目名称'].astype(str)
                    # 记录所有可能是纯数字的项目名称，以便调试
                    numeric_projects = df_detail_filtered[df_detail_filtered['项目名称'].str.replace('.', '', 1).str.isdigit()]
                    if not numeric_projects.empty:
                        log_message += "\n\n注意: 发现可能是纯数字的项目名称:"
                        log_message += f"\n{numeric_projects[['项目名称', '加班时', '总工作时']].head()}"
                    
                    # 过滤纯数字项目名称，但保留数字+字母混合的
                    df_detail_filtered = df_detail_filtered[~df_detail_filtered['项目名称'].str.replace('.', '', 1).str.isdigit()]
                    
                    log_message += f"\n\n过滤后的有效项目行数: {len(df_detail_filtered)}"
                    log_message += "\n所有项目名称:"
                    all_projects = df_detail_filtered['项目名称'].unique()
                    log_message += f"\n{all_projects}"
                    
                    if len(df_detail_filtered) > 0:
                        log_message += "\n有效项目示例:"
                        log_message += f"\n{df_detail_filtered[['项目名称', '加班时', '总工作时']].head(3)}"
                    
                    if status_update_callback:
                        status_update_callback(progress, log_message)
                    
                    # 4. 按项目名称分组统计
                    df_medal = df_detail_filtered.groupby('项目名称').agg({
                        '加班时': 'sum',
                        '总工作时': 'sum'
                    }).reset_index()
                    df_medal.columns = ['项目名称', '加班时总和', '总工作时总和']
                    
                    # 显示分组结果
                    log_message += "\n\n分组统计结果(df_medal):"
                    log_message += f"\n{df_medal}"
                    
                    # 5. 尝试识别统计部分
                    # 检查工作表中end_row后面的内容，寻找可能的项目汇总
                    log_message += "\n\n扫描工作表寻找项目汇总:"
                    
                    project_rows = []
                    for i in range(min(len(full_sheet), end_row+15)):
                        row_str = str(full_sheet.iloc[i].values)
                        # 寻找工作表中的项目汇总部分
                        if i >= end_row:
                            # 打印行内容以辅助调试
                            if i < end_row + 10:  # 只打印前10行汇总部分，避免输出过多
                                log_message += f"\nRow {i+1}: {row_str}"
                            
                            # 寻找与分组结果中项目名称匹配的行
                            for _, medal_row in df_medal.iterrows():
                                project_name = str(medal_row['项目名称'])
                                # 检查该项目是否已经在project_rows中
                                if not any(p['项目名称'] == project_name for p in project_rows):
                                    # 检查当前行是否包含这个项目名称
                                    values = full_sheet.iloc[i].values
                                    for j, val in enumerate(values):
                                        if isinstance(val, str) and project_name in val:
                                            # 找到了项目名称，现在寻找对应的工时和加班时间
                                            work_hours = None
                                            overtime_hours = None
                                            
                                            # 尝试在该行中找到数值
                                            for k, v in enumerate(values):
                                                if isinstance(v, (int, float)) and v > 0:
                                                    if k > j:  # 数值在项目名称后面
                                                        if overtime_hours is None:
                                                            overtime_hours = v
                                                        elif work_hours is None:
                                                            work_hours = v
                                            
                                            # 如果找不到，使用计算值
                                            if work_hours is None:
                                                work_hours = medal_row['总工作时总和']
                                            if overtime_hours is None:
                                                overtime_hours = medal_row['加班时总和']
                                            
                                            project_rows.append({
                                                '项目名称': project_name,
                                                '项目编号': '',
                                                '加班': overtime_hours,
                                                '工时': work_hours
                                            })
                                            log_message += f"\n找到项目行: {project_name}, 加班={overtime_hours}, 工时={work_hours}"
                                            break
                    
                    # 如果没有找到所有项目，使用分组结果中的数据补充
                    found_projects = set(p['项目名称'] for p in project_rows)
                    for _, medal_row in df_medal.iterrows():
                        project_name = str(medal_row['项目名称'])
                        if project_name not in found_projects:
                            project_rows.append({
                                '项目名称': project_name,
                                '项目编号': '',
                                '加班': medal_row['加班时总和'],
                                '工时': medal_row['总工作时总和']
                            })
                            log_message += f"\n使用分组结果补充项目: {project_name}, 加班={medal_row['加班时总和']}, 工时={medal_row['总工作时总和']}"
                    
                    # 创建统计DataFrame
                    df_stat = pd.DataFrame(project_rows)
                    
                    log_message += "\n\n构建的统计部分数据(df_stat):"
                    log_message += f"\n{df_stat}"
                    
                    # 检查统计是否正确
                    is_correct = True
                    for _, row in df_stat.iterrows():
                        if pd.notna(row['项目名称']):
                            medal_row = df_medal[df_medal['项目名称'] == row['项目名称']]
                            if not medal_row.empty:
                                overtime_sum = medal_row['加班时总和'].values[0]
                                total_sum = medal_row['总工作时总和'].values[0]
                                
                                if abs(overtime_sum - row['加班']) > 0.01 or abs(total_sum - row['工时']) > 0.01:
                                    is_correct = False
                                    log_message += f"\n统计不匹配: 项目={row['项目名称']}, 加班时总和={overtime_sum}, 统计加班={row['加班']}"
                                    log_message += f"\n            总工作时总和={total_sum}, 统计工时={row['工时']}"
                    
                    if not is_correct:
                        wrong_employees.append(employee_name)
                        log_message += f"\n员工 {employee_name} 统计有误，已添加到错误列表"
                    
                    # 6. 将结果添加到汇总列表
                    log_message += "\n\n添加到汇总列表的数据:"
                    for _, row in df_stat.iterrows():
                        if pd.notna(row['项目名称']) and pd.notna(row['工时']):
                            project_name = str(row['项目名称'])
                            log_message += f"\n项目名称: {project_name}, 工时: {row['工时']}"
                            
                            all_results.append({
                                '业务类别': '结构',
                                '项目点': company_name,  # 按照需求，这里应该是company_name
                                '工程师': employee_name,
                                '工时': row['工时'],
                                '项目名称': project_name
                            })
                    
                    if status_update_callback:
                        status_update_callback(progress, log_message)
                    
                except Exception as e:
                    log_message += f"\n处理工时明细时出错: {str(e)}"
                    import traceback
                    log_message += f"\n{traceback.format_exc()}"
                    if status_update_callback:
                        status_update_callback(progress, log_message)
            else:
                log_message = f"无法从文件名 {file_name} 中提取信息，尝试使用备用正则表达式"
                # 备用正则表达式
                match = re.match(r'(.+)工时表?-(.+)-(\d+)年(\d+)月', file_name)
                if match:
                    employee_name = match.group(1)
                    company_name = match.group(2)
                    work_year = match.group(3)
                    work_month = match.group(4).lstrip('0')
                    work_year_month = work_year + str(work_month).zfill(2)
                    
                    log_message += f"\n使用备用正则提取信息: 员工={employee_name}, 公司={company_name}, 年={work_year}, 月={work_month}"
                    # 继续处理...
                else:
                    log_message += f"\n所有正则表达式都无法匹配文件名: {file_name}"
                
                if status_update_callback:
                    status_update_callback(progress, log_message)
        except Exception as e:
            log_message = f"处理文件 {file_name} 时出错: {str(e)}"
            import traceback
            log_message += f"\n{traceback.format_exc()}"
            if status_update_callback:
                status_update_callback(progress, log_message)
    
    # 创建结果DataFrame
    result_df = pd.DataFrame(all_results)
    
    # 调试: 检查结果DataFrame
    log_message = "\n========== 最终结果 =========="
    log_message += f"\n结果DataFrame共有 {len(result_df)} 行记录"
    if not result_df.empty:
        log_message += "\n每个员工的项目数量:"
        log_message += f"\n{result_df.groupby('工程师')['项目名称'].count()}"
        log_message += "\n\n结果DataFrame前10行:"
        log_message += f"\n{result_df.head(10)}"
    else:
        log_message += "\n结果DataFrame为空"
    
    # 创建错误清单DataFrame
    error_df = pd.DataFrame({'统计出错人员': wrong_employees})
    
    # 确定输出文件名
    if excel_files:
        match = re.match(r'(.+)工时表?-(.+)-(\d{4})年(\d{1,2})月', excel_files[0])
        if match:
            company_name = match.group(2)
            year = match.group(3)
            month = str(match.group(4)).zfill(2)
            output_file = f"{year}{month}{company_name}工时统计结果.xlsx"
        else:
            current_date = datetime.now()
            output_file = f"{current_date.year}{current_date.month:02d}工时统计结果.xlsx"
    else:
        current_date = datetime.now()
        output_file = f"{current_date.year}{current_date.month:02d}工时统计结果.xlsx"
    
    # 写入Excel
    output_path = os.path.join(folder_path, output_file)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='统计汇总', index=False)
        error_df.to_excel(writer, sheet_name='错误清单', index=False)
    
    summary_message = f"处理完成，结果已保存到 {output_file}"
    summary_message += f"\n处理了 {len(excel_files)} 个文件"
    summary_message += f"\n找到 {len(wrong_employees)} 个统计出错的员工"
    summary_message += f"\n汇总表中有 {len(all_results)} 条记录"
    
    if status_update_callback:
        status_update_callback(1.0, summary_message)
    
    return output_file, result_df, error_df