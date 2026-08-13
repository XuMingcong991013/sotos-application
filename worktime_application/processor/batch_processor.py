import os
import pandas as pd
from openpyxl import load_workbook
from dotenv import load_dotenv
import logging
import datetime

from config.logging_config import setup_logging
from worktime_application.processor.process_worktime_file import process_worktime_file

load_dotenv()

today_date = datetime.datetime.now().strftime('%Y%m%d')
# 设置日志记录
setup_logging(
    log_file=f'log/log_{today_date}.log',
    log_level=logging.INFO,  # 开发时用DEBUG，生产环境用INFO
    console_log=True,
    file_log=True
)

def batch_process_worktime_files(folder_path):
    """批处理工时文件，合并结果和错误员工列表"""
    file_paths = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            full_path = os.path.join(root, file)
            file_paths.append(full_path)
    logging.info(f"上传文件: {file_paths}")
    error_employee_list = []
    merged_result = []

    files_num = len(file_paths)
    i = 1
    for file_path in file_paths:
        logging.info(f"文件处理进度：{i} / {files_num}")
        i += 1
            
        # result_json, error_name = process_worktime_file(file_path=file_path)
        result_json = process_worktime_file(file_path=file_path)
        merged_result.append(result_json)
        # if error_name:
        #     error_employee_list.append(error_name)
    
    return merged_result


def convert_worktime_result_to_dataframe(data_list):
    """
    将员工工时数据的结果列表转换为DataFrame。返回pd.DataFrame: 包含5列(员工姓名,公司名称,项目名称,加班,工时)的DataFrame
    """
    processed_data = [] 
    for item in data_list:
        employee = item['emplyee_name']
        company = item['company_name']
        
        # 处理每个项目
        for project_name, project_data in item['data'].items():
            # 处理可能的拼写差异(overtime_hours/overtime_hours)
            overtime = project_data.get('overtime_hours') or project_data.get('overtime_hours', 0)
            
            processed_data.append({
                '员工姓名': employee,
                '公司名称': company,
                '项目名称': project_name,
                '加班': overtime,
                '工时': project_data['total_hours']
            })
    result_df = pd.DataFrame(processed_data)
    result_df = result_df.sort_values(by=['公司名称', '员工姓名']).reset_index(drop=True)
    # error_df = pd.DataFrame(error_employee_list, columns=['错误人员明细'])
    return result_df

def save_to_excel_with_sheets(result_df, file_name):
    """将两个DataFrame保存到同一个Excel的不同Sheet中：result_df: 要保存到"统计结果"Sheet的DataFrame，error_df: 要保存到"出错人员明细"Sheet的DataFrame"""
    try:
        # 创建一个新的Excel writer
        writer = pd.ExcelWriter(file_name, engine='openpyxl')
        
        # 保存第一个sheet并确保可见
        result_df.to_excel(
            writer,
            sheet_name="统计结果",
            index=False
        )
        
        # # 保存第二个sheet
        # error_df.to_excel(
        #     writer,
        #     sheet_name="出错人员明细",
        #     index=False
        # )
        
        # 获取workbook对象并确保可见性
        workbook = writer.book
        for sheet in workbook.sheetnames:
            workbook[sheet].sheet_state = 'visible'
        
        # 保存文件
        writer.close()
        
        print(f"成功保存到文件: {file_name}")
        return True
    except Exception as e:
        print(f"保存失败: {str(e)}")
        return False


# Example usage:
folder_path = "D:/sotos工作资料/test/1"

result = batch_process_worktime_files(folder_path)
result_df = convert_worktime_result_to_dataframe(data_list=result)
save_to_excel_with_sheets(result_df, file_name="result/output_1.xlsx")