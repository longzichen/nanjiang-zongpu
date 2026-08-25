import re
import sys
import os
import json
import subprocess
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

def parse_issue_body(body_text):
    """解析 GitHub Issue 中的工单字段，自动清洗 Markdown 格式"""
    data = {}
    clean_lines = []
    for line in body_text.splitlines():
        l = re.sub(r'^[-\s*#]+', '', line)
        l = l.replace('**', '').replace('__', '').strip()
        clean_lines.append(l)
    
    clean_text = '\n'.join(clean_lines)

    m_user = re.search(r'提交人姓名[：:\s]*([^\n]+)', clean_text)
    if m_user: data['userName'] = m_user.group(1).strip()
    
    m_phone = re.search(r'联系电话[^\n]*[：:\s]*([^\n]+)', clean_text)
    if m_phone: data['userPhone'] = m_phone.group(1).strip()
    
    m_email = re.search(r'回发邮箱[：:\s]*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', clean_text)
    if m_email: data['userEmail'] = m_email.group(1).strip()
    
    m_type = re.search(r'修改类型[：:\s]*([^\n]+)', clean_text)
    if m_type: data['changeType'] = m_type.group(1).strip()
    
    m_target = re.search(r'目标族人[：:\s]*([^\n]+)', clean_text)
    if m_target: data['targetName'] = m_target.group(1).strip()

    m_father = re.search(r'父亲[：:\s]*([^\s，,/]+)', clean_text)
    if m_father: data['clueFather'] = m_father.group(1).strip()
    
    m_content = re.search(r'具体修改内容[：:\s]*([\s\S]+?)(?:---|📍|$)', clean_text)
    if m_content: data['changeContent'] = m_content.group(1).strip()
    
    return data

def apply_to_ledger_and_rebuild(parsed):
    """
    将审批通过的修订持久化记录到 modifications_history.json，并重新编译 index.html
    """
    ledger_path = "modifications_history.json"
    if not os.path.exists(ledger_path):
        ledger_path = os.path.join(os.path.dirname(__file__), "..", "modifications_history.json")

    mod_list = []
    if os.path.exists(ledger_path):
        try:
            with open(ledger_path, 'r', encoding='utf-8') as f:
                mod_list = json.load(f)
        except Exception:
            mod_list = []

    target_str = parsed.get('targetName', '')
    clean_target = re.split(r'[\s(（]', target_str)[0].replace('江', '').strip()
    
    m_gen = re.search(r'(\d+)世', target_str)
    target_gen = int(m_gen.group(1)) if m_gen else None
    clue_father = parsed.get('clueFather', '').replace('江', '').strip()
    content = parsed.get('changeContent', '').strip()
    change_type = parsed.get('changeType', '')
    user_name = parsed.get('userName', '宗亲')
    user_email = parsed.get('userEmail', '')
    user_phone = parsed.get('userPhone', '')

    today_str = datetime.now().strftime("%Y-%m-%d")

    # 智能解析修改语义
    new_entry = {
        "id": f"mod_{datetime.now().strftime('%Y%m%d%H%M%S')}_{len(mod_list)+1}",
        "target_person": clean_target,
        "target_gen": target_gen,
        "target_father": clue_father,
        "modify_type": change_type,
        "content": content,
        "contributor": user_name,
        "contact": user_phone,
        "email": user_email,
        "approved_at": today_str,
        "status": "APPROVED"
    }

    # 检查是否已有完全相同修改，避免重复追加
    is_dup = any(
        m.get('target_person') == clean_target and m.get('content') == content
        for m in mod_list
    )
    if not is_dup:
        mod_list.append(new_entry)
        with open(ledger_path, 'w', encoding='utf-8') as f:
            json.dump(mod_list, f, ensure_ascii=False, indent=2)
        print(f"✅ Added to permanent modifications ledger: {new_entry['id']}")
    else:
        print(f"ℹ️ Modification already in ledger: {clean_target} -> {content}")

    # 调用 build_html.py 重新渲染
    build_script = "build_html.py"
    if not os.path.exists(build_script):
        build_script = os.path.join(os.path.dirname(__file__), "..", "build_html.py")

    if os.path.exists(build_script):
        print("Rebuilding HTML via build_html.py...")
        res = subprocess.run([sys.executable, build_script], capture_output=True, text=True, encoding='utf-8')
        print(res.stdout)
        if res.stderr:
            print("Stderr:", res.stderr)
    else:
        print("Warning: build_html.py not found in current directory.")

    return new_entry

def main():
    body = ""
    if len(sys.argv) >= 2:
        arg = sys.argv[1]
        if os.path.exists(arg):
            with open(arg, 'r', encoding='utf-8') as f:
                body = f.read()
        else:
            body = arg
    else:
        body = os.environ.get('ISSUE_BODY', '')

    if not body:
        print("Warning: No issue body provided via arguments or ISSUE_BODY environment variable.")
        sys.exit(0)

    parsed = parse_issue_body(body)
    print("Parsed Issue Fields:", parsed)

    entry = apply_to_ledger_and_rebuild(parsed)

    # 邮件通知
    if parsed.get('userEmail'):
        try:
            import genealogy_mailer
            print(f"Sending confirmation receipt to: {parsed['userEmail']}")
            genealogy_mailer.send_user_approved_receipt(
                parsed['userEmail'],
                parsed.get('userName', '宗亲'),
                parsed.get('targetName', '目标族人'),
                parsed.get('changeType', '信息修订'),
                parsed.get('changeContent', ''),
                html_filepath="index.html"
            )
        except Exception as e:
            print(f"Mailer notification error: {e}")

if __name__ == '__main__':
    main()
