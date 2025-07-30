import requests
import jwt
import datetime

# 获取用户信息的示例函数
def get_userinfo():
    return {
        'userName': 'exampleUser',
        'nickName': 'Example Nickname'
    }

# 生成 token
def generate_token():
    cur_info = get_userinfo()
    secret_key = '!@#DFwerw453n'
    payload = {
        'appId': "agent",
        'userId': cur_info['userName'],
        'username': cur_info['nickName'],
        'exp': datetime.datetime.utcnow() + datetime.timedelta(days=3)  # 过期时间 3天
    }
    encoded_jwt = jwt.encode(payload, secret_key, algorithm="HS256")
    return encoded_jwt

def chat_llm():
    data = {
        "messages": [
            {
                "role": "user",
                "content": "1加1为什么等于2"
            }
        ],
        "stream": False,
        "model": "gpt-4o",
        "temperature": 0,
        "presence_penalty": 0,
        "max_tokens": 300,
        "id": "262141969187545089"
    }
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {generate_token()}'
    }
    # url = 'https://172.16.0.80:443/sdw/chatbot/sysai/v1/chat/completions'
    url = 'https://192.168.10.173/sdw/chatbot/sysai/v1/chat/completions'
    response = requests.post(url, json=data, headers=headers, verify=False)
    if response.status_code == 200:
        result = response.json()
        # 假设大模型返回的结果中，内容在某个合适的字段里，比如 'choices' 列表的第一个元素的 'message' 字段的 'content' 字段等，需按实际调整
        content = result.get('choices', [{}])[0].get('message', {}).get('content', "")
        return content
    else:
        return f"请求失败，状态码：{response.status_code}, 响应内容：{response.text}"

# 调用函数
if __name__ == "__main__":
    response_content = chat_llm()
    print(response_content)