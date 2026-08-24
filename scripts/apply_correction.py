import re
import sys
import os
import json

sys.stdout.reconfigure(encoding='utf-8')

def parse_issue_body(body_text):
    """解析 GitHub Issue 中的工单字段"""
    data = {}
    m_user = re.search(r'提交人姓名[：:\s]*([^\n]+)', body_text)
    if m_user: data['userName'] = m_user.group(1).strip()
    
    m_phone = re.search(r'联系电话/微信[：:\s]*([^\n]+)', body_text)
    if m_phone: data['userPhone'] = m_phone.group(1).strip()
    
    m_email = re.search(r'回发邮箱[：:\s]*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', body_text)
    if m_email: data['userEmail'] = m_email.group(1).strip()
    
    m_type = re.search(r'修改类型[：:\s]*([^\n]+)', body_text)
    if m_type: data['changeType'] = m_type.group(1).strip()
    
    m_target = re.search(r'目标族人[：:\s]*([^\n]+)', body_text)
    if m_target: data['targetName'] = m_target.group(1).strip()
    
    m_content = re.search(r'具体修改内容[：:\s]*([\s\S]+?)(?:---|📍|$)', body_text)
    if m_content: data['changeContent'] = m_content.group(1).strip()
    
    return data

def main():
    if len(sys.argv) < 2:
        print("Usage: python apply_correction.py '<issue_body_text>'")
        return

    issue_body = sys.argv[1]
    parsed = parse_issue_body(issue_body)
    print("Parsed Issue Data:", parsed)

    # 1. 自动执行编译生成最新的 index.html
    print("Re-running build_html.py to refresh genealogy data...")
    if os.path.exists("build_html.py"):
        os.system(f"{sys.executable} build_html.py")
    
    # 2. 如果存在最新的 HTML，则复制到 index.html
    candidate_html = r"E:\闲杂\族谱\南江宗谱关系网（最新完美版）.html"
    if os.path.exists(candidate_html):
        with open(candidate_html, 'r', encoding='utf-8') as f:
            content = f.read()
        with open("index.html", 'w', encoding='utf-8') as f:
            f.write(content)
        print("index.html refreshed successfully.")

    # 3. 若用户留了邮箱，自动调用发信引擎发送最新版附件
    user_email = parsed.get('userEmail')
    if user_email:
        try:
            from genealogy_mailer import send_user_approved_receipt
            send_user_approved_receipt(
                user_email=user_email,
                user_name=parsed.get('userName', '宗亲'),
                target_name=parsed.get('targetName', '族人条目'),
                html_filepath="index.html"
            )
            print(f"User receipt mail sent to: {user_email}")
        except Exception as e:
            print(f"Failed to send user mail: {e}")

if __name__ == "__main__":
    main()
