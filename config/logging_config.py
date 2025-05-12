import logging
import sys


def setup_logging(
    log_file: str,
    log_level: int = logging.INFO,
    console_log: bool = True,
    file_log: bool = True,
):
    """日志配置
    
    Args:
        log_file: 日志文件路径
        log_level: 日志级别 (logging.INFO, logging.DEBUG 等)
        console_log: 是否输出到控制台
        file_log: 是否输出到文件
    """
    handlers = []
    
    if console_log:
        # 确保控制台输出使用 UTF-8（防止 Windows 下 GBK 编码错误）
        console_handler = logging.StreamHandler(sys.stdout)  # 使用 sys.stdout 避免部分环境编码问题
        console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        handlers.append(console_handler)
    
    if file_log:
        # 文件日志强制使用 UTF-8 编码
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        handlers.append(file_handler)
    
    logging.basicConfig(
        level=log_level,
        handlers=handlers,
    )