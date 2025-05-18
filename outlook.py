import imaplib
import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import re

def recv_unread_email_by_imap():
    # 邮箱配置信息
    email_address = "3688974995@qq.com"
    email_password = "srnpbgepmqsucjcj"
    imap_server_host = "imap.qq.com"
    imap_server_port = 993

    try:
        email_server = imaplib.IMAP4_SSL(host=imap_server_host, port=imap_server_port)
        print("imap----connect server success")
    except Exception as e:
        print(f"imap----connect time out: {e}")
        return []

    try:
        email_server.login(email_address, email_password)
        print("imap----login success")
    except Exception as e:
        print(f"imap----login failed: {e}")
        email_server.close()
        return []

    try:
        email_server.select('INBOX')
        print("imap----select inbox success")
    except Exception as e:
        print(f"imap----select inbox failed: {e}")
        email_server.close()
        return []

    status, messages = email_server.search(None, 'UNSEEN')
    if status != 'OK':
        print(f"imap----search failed: {messages}")
        email_server.close()
        return []

    if not messages[0]:
        print("imap----no unread emails")
        email_server.close()
        return []

    unread_emails = []
    message_ids = messages[0].split()
    
    for message_id in message_ids:
        try:
            status, msg_data = email_server.fetch(message_id, '(RFC822)')
            if status != 'OK':
                print(f"imap----fetch email {message_id} failed")
                continue

            raw_email = msg_data[0][1]
            
            # 尝试多种解码方式
            try:
                msg = Parser().parsestr(raw_email.decode('utf-8', errors='ignore'))
            except UnicodeDecodeError:
                try:
                    msg = Parser().parsestr(raw_email.decode('gbk', errors='ignore'))
                except UnicodeDecodeError:
                    print(f"imap----无法解码邮件 {message_id}")
                    continue
            
            parsed_email = parse_email(msg)
            unread_emails.append(parsed_email)
            
            # 标记邮件为已读
            email_server.store(message_id, '+FLAGS', '\\Seen')
            
        except Exception as e:
            print(f"imap----process email {message_id} error: {e}")

    email_server.close()
    email_server.logout()
    return unread_emails

def parse_email(msg):
    email_info = {}
    
    # 解析发件人信息
    from_header = msg.get('From', '')
    if from_header:
        hdr, addr = parseaddr(from_header)
        name = decode_str(hdr)
        email_info['from_name'] = name
        email_info['from_email'] = addr

    # 解析收件人信息
    to_header = msg.get('To', '')
    if to_header:
        hdr, addr = parseaddr(to_header)
        name = decode_str(hdr)
        email_info['to_name'] = name
        email_info['to_email'] = addr

    # 解析标题
    subject = msg.get('Subject', '')
    email_info['subject'] = decode_str(subject)

    # 解析正文
    email_info['content'] = extract_email_content(msg)

    return email_info

def extract_email_content(msg):
    content = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get('Content-Disposition'))
            
            # 跳过附件
            if 'attachment' in content_disposition:
                continue
                
            if content_type == 'text/plain':
                charset = guess_charset(part) or 'utf-8'
                try:
                    content = part.get_payload(decode=True).decode(charset, errors='ignore')
                except UnicodeDecodeError:
                    content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                break
            elif content_type == 'text/html':
                # 如果没有找到纯文本部分，使用HTML部分
                if not content:
                    charset = guess_charset(part) or 'utf-8'
                    try:
                        html_content = part.get_payload(decode=True).decode(charset, errors='ignore')
                        content = html_to_markdown(html_content)
                    except UnicodeDecodeError:
                        content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
    else:
        charset = guess_charset(msg) or 'utf-8'
        try:
            content = msg.get_payload(decode=True).decode(charset, errors='ignore')
        except UnicodeDecodeError:
            content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            
    return content

def html_to_markdown(html):
    # 简单的HTML转Markdown
    # 这是一个简化版本，实际应用中可能需要更复杂的转换
    if not html:
        return ""
        
    # 转换换行标签
    html = re.sub(r'<br\s*\/?>', '\n', html)
    html = re.sub(r'<p>', '', html)
    html = re.sub(r'</p>', '\n\n', html)
    
    # 转换标题
    for i in range(6, 0, -1):
        html = re.sub(rf'<h{i}>(.*?)</h{i}>', f"{'#' * i} \\1\n\n", html)
    
    # 转换列表
    html = re.sub(r'<li>', '- ', html)
    html = re.sub(r'</li>', '\n', html)
    html = re.sub(r'<ul>', '', html)
    html = re.sub(r'</ul>', '\n', html)
    
    # 转换粗体和斜体
    html = re.sub(r'<b>(.*?)</b>', '**\\1**', html)
    html = re.sub(r'<strong>(.*?)</strong>', '**\\1**', html)
    html = re.sub(r'<i>(.*?)</i>', '*\\1*', html)
    html = re.sub(r'<em>(.*?)</em>', '*\\1*', html)
    
    # 移除其他标签
    html = re.sub(r'<[^>]+>', '', html)
    
    # 转换实体
    html = re.sub(r'&nbsp;', ' ', html)
    html = re.sub(r'&lt;', '<', html)
    html = re.sub(r'&gt;', '>', html)
    html = re.sub(r'&amp;', '&', html)
    
    return html.strip()

def decode_str(s):
    if not s:
        return ""
        
    value, charset = decode_header(s)[0]
    if charset:
        try:
            return value.decode(charset)
        except (UnicodeDecodeError, AttributeError):
            try:
                return str(value)
            except:
                return ""
    return str(value) if isinstance(value, bytes) else value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset:
        return charset.name
    
    content_type = msg.get('Content-Type', '').lower()
    for item in content_type.split(';'):
        item = item.strip()
        if item.startswith('charset='):
            return item.split('=')[1]
    
    return None

def format_email_to_markdown(email_info):
    # 格式化邮件信息为Markdown字符串
    markdown = f"发件人:\n"
    markdown += f"{email_info.get('from_name', '未知')}\n"
    markdown += f"{email_info.get('from_email', '未知')}\n\n"
    markdown = f"邮件:\n"
    markdown += f"** {email_info.get('subject', '无主题')} **\n\n"
    content = email_info.get('content', '')
    
    # 确保正文内容不为空
    if not content.strip():
        content = "(无正文内容)"
    
    # 对正文内容进行简单的Markdown格式化
    # 移除多余的空行
    content = re.sub(r'\n{3,}', '\n\n', content)
    
    markdown += f"{content}\n"
    
    return markdown

def main():
    unread_emails = recv_unread_email_by_imap()
    markdown_list = []
    
    for email in unread_emails:
        markdown = format_email_to_markdown(email)
        markdown_list.append(markdown)
    
    return markdown_list

if __name__ == '__main__':
    result = main()
    for i in result:
        print(i)
    print(result)