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


webhook = 'YOUR_FEISHU_BOT_WEBBOOK_API'
for i in outlook.main():
  response = send_feishu_message(webhook, i)
  print(response)
