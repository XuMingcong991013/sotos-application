"""
调用百度智能云文档解析 API，将原始简历转换成 Markdown 文件。
"""

import base64
import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv


# ==================== 配置 ====================
TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
CREATE_TASK_URL = "https://aip.baidubce.com/rest/2.0/brain/online/v2/parser/task"
QUERY_TASK_URL = "https://aip.baidubce.com/rest/2.0/brain/online/v2/parser/task/query"


def get_access_token(api_key: str, secret_key: str) -> str:
    """获取百度 Access Token。"""

    response = requests.post(
        TOKEN_URL,
        params={
            "grant_type": "client_credentials",
            "client_id": api_key,
            "client_secret": secret_key,
        },
        timeout=30,
    )
    response.raise_for_status()

    result = response.json()

    if "access_token" not in result:
        raise RuntimeError(f"获取 Access Token 失败：{result}")

    return result["access_token"]


def create_parse_task(file_path: Path, access_token: str) -> str:
    """上传文件并创建文档解析任务。"""

    with file_path.open("rb") as file:
        file_base64 = base64.b64encode(file.read()).decode("utf-8")

    response = requests.post(
        CREATE_TASK_URL,
        params={"access_token": access_token},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "file_data": file_base64,
            "file_name": file_path.name,
            "language_type": "CHN_ENG",
            "angle_adjust": "true",
            "recognize_formula": "false",
            "analysis_chart": "false",
            "parse_image_layout": "false",
            "html_table_format": "true",
        },
        timeout=120,
    )
    response.raise_for_status()

    result = response.json()

    if result.get("error_code") != 0:
        raise RuntimeError(f"创建解析任务失败：{result}")

    task_id = result.get("result", {}).get("task_id")

    if not task_id:
        raise RuntimeError(f"响应中没有 task_id：{result}")

    return task_id


def wait_for_result(
    task_id: str,
    access_token: str,
    timeout_seconds: int = 300,
) -> str:
    """轮询任务状态，成功后返回 Markdown 下载地址。"""

    start_time = time.time()

    while time.time() - start_time < timeout_seconds:
        response = requests.post(
            QUERY_TASK_URL,
            params={"access_token": access_token},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            data={"task_id": task_id},
            timeout=30,
        )
        response.raise_for_status()

        result = response.json()

        if result.get("error_code") != 0:
            raise RuntimeError(f"查询解析任务失败：{result}")

        task_result = result.get("result", {})
        status = task_result.get("status")

        print(f"    百度任务状态：{status}")

        if status == "success":
            markdown_url = task_result.get("markdown_url")

            if not markdown_url:
                raise RuntimeError(f"任务成功，但没有 markdown_url：{result}")

            return markdown_url

        if status == "failed":
            error_message = task_result.get("task_error", "未知错误")
            raise RuntimeError(f"文档解析失败：{error_message}")

        time.sleep(5)

    raise TimeoutError("等待文档解析结果超时。")


def download_markdown(markdown_url: str, output_path: Path) -> None:
    """下载百度生成的 Markdown 文件。"""

    response = requests.get(markdown_url, timeout=60)
    response.raise_for_status()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(response.content)


def BaiduParser(input_file_path: str, output_dir: str) -> str:
    """
    调用百度文档解析API，将原始文件转换为Markdown。

    Args:
        input_file_path: 原始PDF、Word或图片文件路径。
        output_dir: Markdown输出目录。

    Returns:
        生成的Markdown文件绝对路径。
    """
    load_dotenv()

    api_key = os.getenv("BAIDU_API_KEY")
    secret_key = os.getenv("BAIDU_SECRET_KEY")

    if not api_key or not secret_key:
        raise RuntimeError(
            "没有读取到百度密钥，请检查 .env 中是否配置了 "
            "BAIDU_API_KEY 和 BAIDU_SECRET_KEY。"
        )
    
    input_path = Path(input_file_path)

    if not input_path.exists():
        raise FileNotFoundError(f"找不到输入文件：{input_path.resolve()}")

    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)
    output_path = output_directory / f"{input_path.stem}.md"

    print("    正在获取百度 Access Token……")
    access_token = get_access_token(api_key, secret_key)

    print("    正在上传文件并创建解析任务……")
    task_id = create_parse_task(input_path, access_token)
    print(f"    任务创建成功，task_id：{task_id}")

    print("    正在等待百度解析……")
    markdown_url = wait_for_result(task_id, access_token)

    print("    正在下载 Markdown……")
    download_markdown(markdown_url, output_path)

    print(f"    百度解析完成：{output_path.resolve()}")
    return str(output_path.resolve())
