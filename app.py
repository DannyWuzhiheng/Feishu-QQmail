import requests
import json
import re
import outlook


def send_feishu_message(webhook, message):
    headers = {
        'Content-Type': 'application/json'
    }
    data = {
    "msg_type": "interactive",
    "card": {
        "elements": [{
                "tag": "markdown",
                "content":message
        }]}}
    response = requests.post(webhook, headers=headers, json=data)
    return response.json()


webhook = 'https://open.feishu.cn/open-apis/bot/v2/hook/e20a0b68-7210-4b64-a9e8-2b68e609c429'
for i in outlook.main():
  response = send_feishu_message(webhook, i)
  print(response)