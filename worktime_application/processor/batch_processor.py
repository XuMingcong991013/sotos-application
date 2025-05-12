import os
from worktime_application.processor.process_worktime_file import process_worktime_file


def batch_process_worktime_files(file_paths, initial_error_list=None):
    """批处理工时文件，合并结果和错误员工列表"""
    error_employee_list = initial_error_list.copy() if initial_error_list else []
    
    merged_result_json = {}
    
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"Warning: File not found - {file_path}")
            continue
            
        result_json, updated_error_list = process_worktime_file(
            file_path=file_path, 
            error_employee_list=error_employee_list
        )
        
        error_employee_list = updated_error_list
        
        for key, value in result_json.items():
            if key in merged_result_json:
                if isinstance(value, list):
                    merged_result_json[key].extend(value)
                elif isinstance(value, dict):
                    merged_result_json[key].update(value)
                else:
                    merged_result_json[key] = value
            else:
                merged_result_json[key] = value
    
    return merged_result_json, error_employee_list

# Example usage:
file_paths = [
    "D:/test/于浩鑫-工时表-重庆安道拓-2025年03月.xlsx",
    "D:/test/文翊宁工时表-重庆安道拓-2025年03月.xlsx"
]

result_json, error_employee_list = batch_process_worktime_files(file_paths, [])
print("error_employee_list: ", error_employee_list)
print("result_json: ", result_json)

