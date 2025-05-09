def get_completions(user_prompt, system_prompt, stream=False, temperature=0.95, top_p=0.1):
    """请求模型，获取响应
    入参：
        - user_prompt: 用户输入
        - system_prompt: 身份设定
        - stream: 是否使用流式返回
        - temperature: (0,1], 控制生成文本的随机性
        - top_p: (0,1], 模型解码器只考虑从前top_p的概率的候选集中取tokens
    返回：
        - response: 模型返回的完整响应
    """
    url = os.getenv("MODEL_API_URL")
    headers = {
        "Content-Type": "application/json",
        "Authorization": os.getenv("API_KEY")
    }
    payload = {
        "user":"sotos001",
        "model":os.getenv("MODEL_NAME"),
        "messages":[
            {"role":"system", "content":system_prompt},
            {"role": "user", "content": user_prompt}  
        ],
        "stream":stream,
        "temperature": temperature,
        "top_p": top_p,
    }

    try:
        logging.info(f"发送请求到模型API...")
        logging.info(f"system_prompt：{system_prompt}")
        logging.info(f"user_prompt：{user_prompt}")
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()  # 如果请求失败会抛出HTTPError异常
        logging.info(f"请求成功，状态码: {response.status_code}")
        response = response.json()
        logging.info("模型完整响应内容: %s", response)

        return response
    
    except requests.exceptions.HTTPError as errh:
        logging.error(f"HTTP错误: {errh}\n请求URL: {url}\n请求负载: {payload}")
    except requests.exceptions.ConnectionError as errc:
        logging.error(f"连接错误: {errc}\n请检查网络连接和服务地址: {url}")
    except requests.exceptions.Timeout as errt:
        logging.error(f"请求超时: {errt}\n考虑增加超时时间或重试")
    except requests.exceptions.RequestException as err:
        logging.error(f"请求异常: {err}")
    except Exception as e:
        logging.error(f"未处理的异常: {e}", exc_info=True)