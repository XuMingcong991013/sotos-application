import os
import re
import datetime
import logging
import json
from dotenv import load_dotenv

from config.logging_config import setup_logging
from worktime_application.call_llm import get_completions
from worktime_application.processor.excel_to_text import excel_to_markdown
from worktime_application.processor.system_prompts import EXTRACT_INFO_PROMPT, EXTRACT_DETAIL_WORKTIME_DATA_PROMPT, EXTRACT_SUMMARY_WORKTIME_DATA_PROMPT

load_dotenv()

today_date = datetime.datetime.now().strftime('%Y%m%d')
# 设置日志记录
setup_logging(
    log_file=f'log/log_{today_date}.log',
    log_level=logging.INFO,  # 开发时用DEBUG，生产环境用INFO
    console_log=True,
    file_log=True
)



def process_worktime_file(file_path: str, error_employee_list: list):
    """
    处理工时文件并提取相关信息
    """
    logging.info(f"开始处理工时文件：{file_path}")
    if error_employee_list is None:
        error_employee_list = []
    
    # 提取Excel文件文本
    markdown_text = excel_to_markdown(file_path)
    logging.info(f"Excel表格提取文本：\n{markdown_text}")
    
    # 1.从表格文本中，提取员工姓名和公司名称
    logging.info("从表格文本中，提取员工姓名和公司名称...")
    system_prompt_1 = EXTRACT_INFO_PROMPT
    user_prompt_1 = markdown_text
    try:
        response_1 = get_completions(user_prompt=user_prompt_1, system_prompt=system_prompt_1)
        response_1 = response_1["choices"][0]["message"]["content"]   
        logging.info(f"模型返回：{response_1}\n")
        try:
            response_1 = json.loads(response_1)
            employee_name = response_1.get("employee_name", "")
            company_name = response_1.get("company_name", "")
            logging.info(f"解析模型返回：\n - 员工姓名：{employee_name},  - 公司名称：{company_name}")
        except Exception as e:
            logging.info(f"模型未返回合法JSON数据。错误：{e}")
    except Exception as e:
        logging.info(f"模型返回错误：{e}")
    
    # 2.从表格文本中，提取工时明细数据,并按照项目名称做分组
    logging.info("\n\n从表格文本中，提取工时明细数据,并按照项目名称做分组...")
    system_prompt_2 = EXTRACT_DETAIL_WORKTIME_DATA_PROMPT
    user_prompt_2 = markdown_text
    try:
        response_2 = get_completions(user_prompt=user_prompt_2, system_prompt=system_prompt_2)
        response_2 = response_2["choices"][0]["message"]["content"]   
        logging.info(f"模型返回：{response_2}\n")
        try:
            response_2 = json.loads(response_2)
            logging.info(f"解析模型返回, 源数据表中的工时明细数据：{response_2}")
        except Exception as e:
            logging.info(f"模型未返回合法JSON数据。错误：{e}")
    except Exception as e:
        logging.info(f"模型返回错误：{e}")

    
    # 3.根据分组的工时明细数据，计算各项目的总工时和加班工时
    def calculate_project_worktime(detail_worktime_data):
        """
        计算每个项目的overtime_hours和total_hours总和
        """
        result = {}
        
        for project_name, daily_data in detail_worktime_data.items():
            overtime_sum = 0
            total_hours_sum = 0
            
            # 遍历项目的每日数据
            for date, hours_data in daily_data.items():
                # 处理可能的拼写错误 (overtours -> overtime_hours)
                if 'overtime_hours' in hours_data:
                    overtime_sum += hours_data['overtime_hours']
                elif 'overtours' in hours_data:  # 处理可能的拼写错误
                    overtime_sum += hours_data['overtours']
                    
                # 累加总工时
                if 'total_hours' in hours_data:
                    total_hours_sum += hours_data['total_hours']
            
            # 将计算结果存储到结果字典中
            result[project_name] = {
                'overtime_hours_sum': overtime_sum,
                'total_hours_sum': total_hours_sum
            } 
        return result
    
    grouped_worktime_data = calculate_project_worktime(detail_worktime_data=response_2)
    logging.info(f"工时分组数据统计：{grouped_worktime_data}")
    
    # 4.从表格文本中，提取汇总数据
    logging.info("\n\n从表格文本中，提取汇总数据...")
    system_prompt_3 = EXTRACT_SUMMARY_WORKTIME_DATA_PROMPT
    user_prompt_3 = markdown_text
    try:
        response_3 = get_completions(user_prompt=user_prompt_3, system_prompt=system_prompt_3)
        response_3 = response_3["choices"][0]["message"]["content"]
        logging.info(f"模型返回：\n{response_3}\n")
        try:
            response_3 = json.loads(response_3)
            logging.info(f"源数据表中的工时汇总数据：{response_3}")
        except Exception as e:
            logging.info(f"模型未返回合法JSON数据。错误：{e}")
    except Exception as e:
        logging.info(f"模型返回错误：{e}")

    
    # 5.将原表中的汇总工时数据与统计的工时数据作比较，判断是否一致，一致是True，不一致是False
    logging.info("比较原表中的汇总工时数据与统计的工时数据...")
    logging.info(f"- 统计的工时数据：{grouped_worktime_data}")
    logging.info(f"- 原表中的工时汇总数据：{response_3}")
    if grouped_worktime_data == response_3:
        compare_result = True
    else:
        compare_result = False
    logging.info(f"原表中的工时汇总数据与统计的工时数据的比较结果：{compare_result}")
    
    # 错误人员明细：
    if not compare_result:
        error_employee_list.append(employee_name)
    logging.info(f"错误人员明细：{error_employee_list}")

    result_json = {"emplyee_name": employee_name, "company_name":company_name, "data": grouped_worktime_data}
    
    return result_json, error_employee_list







