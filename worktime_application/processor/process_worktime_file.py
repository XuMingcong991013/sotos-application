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


def calculate_project_worktime(detail_worktime_data):
    """按照项目名称做分组，计算每个项目的overtime_hours和total_hours总和"""
    if not isinstance(detail_worktime_data, dict):
        logging.error(f"无效的工时数据格式，期望字典类型，实际得到: {type(detail_worktime_data)}")
        return {}
        
    grouped = {}
    
    for date, details in detail_worktime_data.items():
        # 跳过空字典或缺失关键字段的条目
        if not details or 'project_name' not in details:
            continue

        project_name = details["project_name"]
        if not project_name:  # 跳过空项目名的条目
            continue
        
        # 确保有必要的字段
        if 'overtime_hours' not in details or 'total_hours' not in details:
            continue
            
        if project_name not in grouped:
            grouped[project_name] = []
            
        grouped[project_name].append({
            "date": date,
            "overtime_hours": details["overtime_hours"],
            "total_hours": details["total_hours"]
        })
    result = {
        project_name: {
            'overtime_hours': sum(item['overtime_hours'] for item in details),
            'total_hours': sum(item['total_hours'] for item in details)
        } for project_name, details in grouped.items()
    
    }
    return result



def process_worktime_file(file_path: str):
    """
    处理工时文件并提取相关信息
    """
    logging.info(f"开始处理工时文件：{file_path}")
    
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
            logging.error(f"解析模型返回失败，原始返回内容：{response_1}\n错误详情：{str(e)}")
            raise ValueError(f"无法解析模型返回的JSON数据: {str(e)}")
    except Exception as e:
        logging.error(f"模型调用失败，请求参数：{user_prompt_1[:200]}...\n错误详情：{str(e)}")
        raise ValueError(f"模型调用失败: {str(e)}")
    
    # # 2.从表格文本中，提取工时明细数据,并按照项目名称做分组
    # logging.info("\n\n从表格文本中，提取工时明细数据,并按照项目名称做分组...")
    # system_prompt_2 = EXTRACT_DETAIL_WORKTIME_DATA_PROMPT
    # user_prompt_2 = markdown_text
    # try:
    #     response_2 = get_completions(user_prompt=user_prompt_2, system_prompt=system_prompt_2)
    #     response_2 = response_2["choices"][0]["message"]["content"]   
    #     logging.info(f"模型返回：{response_2}\n")
    #     try:
    #         response_2 = json.loads(response_2)
    #         logging.info(f"解析模型返回, 源数据表中的工时明细数据：{response_2}")
    #     except Exception as e:
    #         logging.error(f"解析模型返回失败，原始返回内容：{response_2}\n错误详情：{str(e)}。尝试再次请求模型")
    #         try:
    #             response_2 = get_completions(user_prompt=user_prompt_2, system_prompt=system_prompt_2)
    #             response_2 = response_2["choices"][0]["message"]["content"]
    #             response_2 = json.loads(response_2)
    #         except Exception as e:
    #             logging.error(f"再次请求模型失败，请求参数：{user_prompt_2[:200]}...\n错误详情：{str(e)}")
    #         raise ValueError(f"无法解析模型返回的JSON数据: {str(e)}")
    # except Exception as e:
    #     logging.error(f"模型调用失败，请求参数：{user_prompt_2[:200]}...\n错误详情：{str(e)}")
    #     raise ValueError(f"模型调用失败: {str(e)}")

    
    # # 3.根据工时明细数据，按照项目名称做分组，计算各项目的总工时和加班工时  
    # grouped_worktime_data = calculate_project_worktime(detail_worktime_data=response_2)
    # logging.info(f"工时分组数据统计：{grouped_worktime_data}")
    
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
            logging.error(f"解析模型返回失败，原始返回内容：{response_3}\n错误详情：{str(e)}")
            raise ValueError(f"无法解析模型返回的JSON数据: {str(e)}")
    except Exception as e:
        logging.error(f"模型调用失败，请求参数：{user_prompt_3[:200]}...\n错误详情：{str(e)}")
        raise ValueError(f"模型调用失败: {str(e)}")

    
    # # 5.将原表中的汇总工时数据与统计的工时数据作比较，判断是否一致，一致是True，不一致是False
    # logging.info("比较原表中的汇总工时数据与统计的工时数据...")
    # logging.info(f"- 统计的工时数据：{grouped_worktime_data}")
    # logging.info(f"- 原表中的工时汇总数据：{response_3}")
    # if grouped_worktime_data == response_3:
    #     compare_result = True
    # else:
    #     compare_result = False
    # logging.info(f"原表中的工时汇总数据与统计的工时数据的比较结果：{compare_result}")
    
    # # 错误人员明细：
    # if not compare_result:
    #     error_name = employee_name
    #     logging.info(f"错误人员明细：{error_name}")
    # else:
    #     error_name = ""
    #     logging.info("无错误人员")

    result_json = {"emplyee_name": employee_name, "company_name":company_name, "data": response_3}
    
    return result_json




result_json = process_worktime_file(file_path="C:/Users/admin/Desktop/工时测试数据/蔡爽工时表-重庆安道拓-2025年03月.xlsx")
print("result_json",result_json)
# print("error_name:", error_name)

