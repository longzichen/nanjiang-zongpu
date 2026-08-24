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

def apply_correction_to_json(parsed):
    """
    直接修改 genealogy_data.json 数据库
    """
    json_path = "genealogy_data.json"
    if not os.path.exists(json_path):
        json_path = r"C:\Users\longzichen\.gemini\antigravity\scratch\nanjiang-zongpu\genealogy_data.json"

    with open(json_path, 'r', encoding='utf-8') as f:
        db = json.load(f)

    nodes = db.get('all_nodes', [])
    target_str = parsed.get('targetName', '')
    clean_target = re.split(r'[\s(（]', target_str)[0].replace('江', '')
    clue_father = parsed.get('clueFather', '').replace('江', '')
    content = parsed.get('changeContent', '').strip()
    change_type = parsed.get('changeType', '')

    print(f"Executing Cloud Auto-Merge: target='{clean_target}', father='{clue_father}', content='{content}'")

    matched_node = None
    # 优先匹配名字和父辈
    for n in nodes:
        if n.get('name') == clean_target:
            if not clue_father or n.get('father_hint') == clue_father:
                matched_node = n
                break
    if not matched_node:
        for n in nodes:
            if n.get('name') == clean_target:
                matched_node = n
                break

    if matched_node:
        print(f"Matched Node in DB: id={matched_node.get('id')}, name={matched_node.get('name')}, gen={matched_node.get('gen')}")
        
        # 1. 增补或修正配偶
        m_husband = re.search(r'(?:夫|适|配)[：:\s]*([^\s，,。]+)', content)
        m_wife = re.search(r'(?:妻|配|娶)[：:\s]*([^\s，,。]+)', content)

        if m_husband or ('配偶' in change_type and matched_node.get('gender') == 'female') or '夫' in content:
            h_name = m_husband.group(1) if m_husband else content.replace('夫', '').strip()
            matched_node['wife'] = h_name
            if '适' not in matched_node.get('detail', ''):
                matched_node['detail'] = matched_node.get('detail', '') + f"。适{h_name}。"
            print(f"Updated spouse (husband): {h_name}")

        elif m_wife or ('配偶' in change_type and matched_node.get('gender') != 'female') or '妻' in content:
            w_name = m_wife.group(1) if m_wife else content.replace('妻', '').strip()
            matched_node['wife'] = w_name
            if '妻' not in matched_node.get('detail', ''):
                matched_node['detail'] = matched_node.get('detail', '') + f"。妻{w_name}。"
            print(f"Updated spouse (wife): {w_name}")

        # 2. 增补子嗣
        elif '子' in content or '女' in content or '增加' in content:
            # 提取名字
            m_child = re.search(r'(?:儿子|女儿|子|女)[：:\s]*([^\s，,。]+)', content)
            c_name = m_child.group(1).replace('江', '') if m_child else content
            is_daughter = '女' in content
            c_gen = matched_node.get('gen', 30) + 1
            new_id = f"node_{len(nodes) + 1}"
            
            # 生年提取
            m_yr = re.search(r'(\d{4})年?', content)
            c_birth = int(m_yr.group(1)) if m_yr else None

            new_node = {
                'id': new_id,
                'name': c_name,
                'full_name': '江' + c_name,
                'gen': c_gen,
                'father_id': matched_node.get('id'),
                'father_hint': matched_node.get('name'),
                'gender': 'female' if is_daughter else 'male',
                'branch': matched_node.get('branch', '二房'),
                'detail': f"父: {matched_node.get('name')} ({matched_node.get('gen')}世)。" + content,
                'wife': '',
                'birth_year': c_birth,
                'children': []
            }
            nodes.append(new_node)
            if 'children' not in matched_node:
                matched_node['children'] = []
            matched_node['children'].append(new_id)
            print(f"Successfully added child node: {c_name} (gen {c_gen}) under {matched_node.get('name')}")

        # 3. 补充生平
        else:
            matched_node['detail'] = matched_node.get('detail', '') + f"。{content}。"
            print("Appended detail to matched node.")

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

    print("Saved updated genealogy_data.json successfully.")

def rebuild_html_from_json():
    """根据最新的 JSON 数据重构 index.html"""
    json_path = "genealogy_data.json"
    if not os.path.exists(json_path):
        json_path = r"C:\Users\longzichen\.gemini\antigravity\scratch\nanjiang-zongpu\genealogy_data.json"

    with open(json_path, 'r', encoding='utf-8') as f:
        db = json.load(f)

    nodes = db.get('all_nodes', [])
    nodes_json_str = json.dumps(nodes, ensure_ascii=False)

    html_path = "index.html"
    if not os.path.exists(html_path):
        html_path = r"C:\Users\longzichen\.gemini\antigravity\scratch\nanjiang-zongpu\index.html"

    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    # 替换其中的 rawNodes 数据
    pattern = r'const rawNodes = \[.*?\];'
    replacement = f'const rawNodes = {nodes_json_str};'
    
    new_html = re.sub(pattern, replacement, html_content, flags=re.DOTALL)
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new_html)

    # 同步更新本地发布文件
    target_local = r"E:\闲杂\族谱\南江江氏宗谱世系关系网.html"
    try:
        with open(target_local, 'w', encoding='utf-8') as f:
            f.write(new_html)
    except Exception:
        pass

    print(f"Rebuilt index.html with {len(nodes)} nodes successfully.")

def main():
    if len(sys.argv) < 2:
        print("Usage: python apply_correction.py '<issue_body_text>'")
        return

    issue_body = sys.argv[1]
    parsed = parse_issue_body(issue_body)
    print("Parsed Issue Data:", parsed)

    apply_correction_to_json(parsed)
    rebuild_html_from_json()

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
