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
    
    m_phone = re.search(r'联系电话[^\n]*[：:\s]*([^\n]+)', body_text)
    if m_phone: data['userPhone'] = m_phone.group(1).strip()
    
    m_email = re.search(r'回发邮箱[：:\s]*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', body_text)
    if m_email: data['userEmail'] = m_email.group(1).strip()
    
    m_type = re.search(r'修改类型[：:\s]*([^\n]+)', body_text)
    if m_type: data['changeType'] = m_type.group(1).strip()
    
    m_target = re.search(r'目标族人[：:\s]*([^\n]+)', body_text)
    if m_target: data['targetName'] = m_target.group(1).strip()

    m_father = re.search(r'父亲[：:\s]*([^\s，,/]+)', body_text)
    if m_father: data['clueFather'] = m_father.group(1).strip()
    
    m_content = re.search(r'具体修改内容[：:\s]*([\s\S]+?)(?:---|📍|$)', body_text)
    if m_content: data['changeContent'] = m_content.group(1).strip()
    
    return data

def apply_patch_to_build_script(parsed):
    """
    智能语义自动合入引擎：
    根据工单自动将配偶、子嗣、生卒年修改写入 build_html.py 的动态补丁逻辑中
    """
    target_name = parsed.get('targetName', '')
    # 提取纯名字（如 "文 (31世 · 二房)" -> "文"）
    clean_target = re.split(r'[\s(（]', target_name)[0].replace('江', '')
    clue_father = parsed.get('clueFather', '').replace('江', '')
    content = parsed.get('changeContent', '').strip()
    change_type = parsed.get('changeType', '')

    print(f"Applying semantic patch: target='{clean_target}', father='{clue_father}', content='{content}'")

    build_py_path = "build_html.py"
    if not os.path.exists(build_py_path):
        build_py_path = r"C:\Users\longzichen\.gemini\antigravity\scratch\nanjiang-zongpu\build_html.py"
    if not os.path.exists(build_py_path):
        build_py_path = r"C:\Users\longzichen\.gemini\antigravity\scratch\build_html.py"

    with open(build_py_path, 'r', encoding='utf-8') as f:
        code = f.read()

    # 构造补丁代码片段
    patch_code = ""
    
    # 1. 增补配偶 (如 "夫 曾德亮" 或 "妻 李四")
    m_husband = re.search(r'(?:夫|配|适|女婿)[：:\s]*([^\s，,。]+)', content)
    m_wife = re.search(r'(?:妻|配|娶)[：:\s]*([^\s，,。]+)', content)

    if '配偶' in change_type or m_husband or m_wife or '夫' in content or '妻' in content:
        if m_husband or '夫' in content:
            h_name = m_husband.group(1) if m_husband else content.replace('夫', '').strip()
            patch_code = f"""
        if clean_name == '{clean_target}' and '{h_name}' not in detail:
            detail += '。适{h_name}。'
"""
        elif m_wife or '妻' in content:
            w_name = m_wife.group(1) if m_wife else content.replace('妻', '').strip()
            patch_code = f"""
        if clean_name == '{clean_target}' and '{w_name}' not in detail:
            detail += '。妻{w_name}。'
"""

    # 2. 增补子嗣 (如 "增加儿子 江宗泽，2023年生")
    elif '子' in content or '女' in content or '增补' in change_type:
        patch_code = f"""
        if clean_name == '{clean_target}' and '{content[:4]}' not in detail:
            detail += '。{content}。'
"""
    # 3. 常规生平或生卒年修正
    else:
        patch_code = f"""
        if clean_name == '{clean_target}':
            detail += '。{content}。'
"""

    # 将补丁注入到 build_html.py 的 clean_name 判定逻辑后
    anchor = "clean_wife, full_wife = extract_wife(detail)"
    if anchor in code and patch_code:
        if patch_code.strip() not in code:
            code = code.replace(anchor, patch_code + "\n        " + anchor)
            with open(build_py_path, 'w', encoding='utf-8') as f:
                f.write(code)
            print("Successfully injected patch into build_html.py")

    # 触发重新编译
    os.system(f"{sys.executable} {build_py_path}")
    print("Rebuilt HTML successfully.")

def main():
    if len(sys.argv) < 2:
        print("Usage: python apply_correction.py '<issue_body_text>'")
        return

    issue_body = sys.argv[1]
    parsed = parse_issue_body(issue_body)
    print("Parsed Issue Data:", parsed)

    apply_patch_to_build_script(parsed)

    # 同步到 index.html
    candidate_html = r"E:\闲杂\族谱\南江宗谱关系网（最新完美版）.html"
    if not os.path.exists(candidate_html):
        candidate_html = "index.html"
    if os.path.exists(candidate_html):
        with open(candidate_html, 'r', encoding='utf-8') as f:
            content = f.read()
        with open("index.html", 'w', encoding='utf-8') as f:
            f.write(content)

    # 发送回执邮件给用户
    user_email = parsed.get('userEmail')
    if user_email and '@' in user_email:
        try:
            sys.path.append('scripts')
            from genealogy_mailer import send_user_approved_receipt
            send_user_approved_receipt(
                user_email=user_email,
                user_name=parsed.get('userName', '宗亲'),
                target_name=parsed.get('targetName', '族人条目'),
                html_filepath="index.html"
            )
        except Exception as e:
            print(f"Failed to send user mail: {e}")

if __name__ == "__main__":
    main()
