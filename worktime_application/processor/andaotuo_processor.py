import os
import re
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import tempfile
import shutil
import time



def extract_info_from_filename(file_name):
    """从文件名中提取员工姓名、公司名称和工作年月，确定文件读取范围"""
    match = re.match(r'(.+)工时表?-(.+)-(\d{4})年(\d{1,2})月', file_name)
    if match:
        employee_name = match.group(1)
        company_name = match.group(2)
        work_year = match.group(3)  # 直接提取4位数年份
        work_month = match.group(4).lstrip('0')  # 去除前导零
        work_year_month = work_year + str(work_month).zfill(2)

    work_month_int = int(work_month)
    work_year_int = int(work_year)

    if work_month_int in [1, 3, 5, 7, 8, 10, 12]:
        end_row = 34
    elif work_month_int in [4, 6, 9, 11]:
        end_row = 33
    elif work_month_int == 2 and work_year_int % 4 == 0:
        end_row = 32
    else:
        end_row = 31

    return employee_name, company_name, work_year_month, end_row



def excel_num_to_date(serial):
    """将Excel序列号转为Python日期"""
    return datetime(1899, 12, 30) + timedelta(days=serial)



def read_data_from_excel(file_path):
    """
    从Excel文件中读取数据，处理并返回DataFrame
    end_row: 读取的最后一行的行号，根据月份动态调整
    """
    file_name = os.path.basename(file_path)
    print(f"正在读取文件: {file_name}")
    employee_name, company_name, work_year_month, end_row = extract_info_from_filename(file_name)   
    print(f"员工姓名: {employee_name}, 公司名称: {company_name}, 工作年月: {work_year_month}")
 
    xl = pd.ExcelFile(file_path, engine='openpyxl')
    sheet_names = xl.sheet_names
        
    if employee_name in sheet_names:
         sheet_name = employee_name
    else:
        sheet_name = sheet_names[0] if sheet_names else None
        
    full_sheet_data = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    # 1.读取工时明细数据（原表格的第4行作为表头）
    df_detail = pd.read_excel(file_path, sheet_name=sheet_name, usecols="A:F", header=3, nrows=end_row-3, engine='openpyxl')
    df_detail['日期'] = df_detail['日期'].apply(excel_num_to_date)
    df_detail = df_detail.dropna(subset=['总工作时间'])
    df_detail = df_detail.drop(['开始工作时间', '工作结束时间'], axis=1)
    print(f"\n详细工时数据: \n{df_detail}")
    
    # 2.读取工时汇总数据（原表格的第end_row+3行作为表头） 
    df_summary = pd.read_excel(file_path, sheet_name=sheet_name, usecols="B:F", header=end_row+1, nrows=4, engine='openpyxl') 
    df_summary = df_summary.drop('Unnamed: 4', axis=1)
    print(f"原始汇总表数据:\n{df_summary}")

    return df_detail, df_summary, employee_name, company_name



def process_single_file(file_path, wrong_employees: list):
    """计算工作时间和加班时间，比较工时明细与汇总数据是否一致"""
    df_detail, df_summary, employee_name, company_name = read_data_from_excel(file_path)
    df_detail = df_detail.drop('日期', axis=1)
    df_summary = df_summary.drop('项目编号', axis=1)

    print(f"df_detail:\n{df_detail}")
    print(f"df_summary:\n{df_summary}")

    # 分组
    df_grouped = df_detail.groupby('项目名称').agg({
        '加班时间': lambda x: round(x.sum(), 1),
        '总工作时间': lambda x: round(x.sum(), 1)
    }).reset_index()
    df_grouped.columns = ['项目名称', '加班', '工时'] # 重命名列名
    print(df_grouped)

 
    # 执行数据比较逻辑
    df_grouped_sorted = df_grouped.sort_values(by=list(df_grouped.columns)).reset_index(drop=True)
    df_summary_sorted = df_summary.sort_values(by=list(df_summary.columns)).reset_index(drop=True)
    
    # 统一数值类型为浮点数
    df_grouped_sorted[['加班', '工时']] = df_grouped_sorted[['加班', '工时']].astype(float)
    df_summary_sorted[['加班', '工时']] = df_summary_sorted[['加班', '工时']].astype(float)
    
    compare_result = df_grouped_sorted.equals(df_summary_sorted)
    print(f"数据比较结果: {compare_result}")

    if compare_result is False:
        wrong_employees.append(employee_name)

    df_result_single = df_grouped_sorted.copy()
    df_result_single['业务类别'] = "结构"
    df_result_single['项目点'] = company_name
    df_result_single['工程师'] = employee_name
    print("df_result_single:\n", df_result_single)
    return df_result_single, wrong_employees


# wrong_employees=[]
# df_result_single, wrong_employees = compare_worktime(file_path='D:/sotos工作资料/202503 安道拓（重庆）/邹同伟工时表-重庆安道拓-2025年03月.xlsx', wrong_employees=wrong_employees)
# print(wrong_employees)

