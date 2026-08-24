import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formataddr
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.qq.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "")

def send_admin_alert(issue_info):
    """当有新纠错提交时，向管理员发送通知邮件"""
    try:
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15)
        server.login(SMTP_USER, SMTP_PASS)

        msg = MIMEMultipart()
        msg['From'] = formataddr((str(Header('南江宗谱工单中心', 'utf-8')), SMTP_USER))
        msg['To'] = formataddr((str(Header('宗族管理员', 'utf-8')), ADMIN_EMAIL))
        msg['Subject'] = Header(f"【宗谱纠错申请】{issue_info.get('title', '新修改申请')}", 'utf-8')

        html_content = f"""
        <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 650px; margin: auto; padding: 25px; background: #f8fafc; border-radius: 16px; border: 1px solid #e2e8f0;">
            <div style="display: flex; align-items: center; border-bottom: 2px solid #3b82f6; padding-bottom: 12px; margin-bottom: 20px;">
                <h2 style="color: #1e3a8a; margin: 0; font-size: 20px;">📋 南江宗谱 · 收到新的族人信息纠错工单</h2>
            </div>
            
            <div style="background: #ffffff; padding: 18px; border-radius: 12px; border: 1px solid #cbd5e1; box-shadow: 0 2px 6px rgba(0,0,0,0.04); margin-bottom: 20px;">
                <table style="width: 100%; font-size: 14px; line-height: 1.8; color: #334155;">
                    <tr><td style="width: 120px; font-weight: bold; color: #64748b;">提交人：</td><td style="font-weight: bold; color: #0f172a;">{issue_info.get('userName', '未提供')}</td></tr>
                    <tr><td style="font-weight: bold; color: #64748b;">联系方式：</td><td>{issue_info.get('userPhone', '未提供')}</td></tr>
                    <tr><td style="font-weight: bold; color: #64748b;">回发邮箱：</td><td style="color: #2563eb; font-weight: bold;">{issue_info.get('userEmail', '未提供')}</td></tr>
                    <tr><td style="font-weight: bold; color: #64748b;">修改类型：</td><td><span style="background: #fef3c7; color: #92400e; padding: 2px 8px; border-radius: 6px; font-size: 12px; font-weight: bold;">{issue_info.get('changeType', '信息修正')}</span></td></tr>
                    <tr><td style="font-weight: bold; color: #64748b;">目标族人：</td><td style="color: #1d4ed8; font-weight: bold;">{issue_info.get('targetName', '线索录入')} ({issue_info.get('targetGen', '')}世 · {issue_info.get('targetBranch', '')})</td></tr>
                    <tr><td style="font-weight: bold; color: #64748b;">长辈线索：</td><td>父亲: {issue_info.get('clueFather', '无')} / 爷爷: {issue_info.get('clueGrandfather', '无')}</td></tr>
                </table>
            </div>

            <div style="background: #f1f5f9; padding: 15px; border-radius: 12px; border-left: 4px solid #f59e0b; margin-bottom: 20px;">
                <div style="font-weight: bold; color: #92400e; font-size: 13px; margin-bottom: 6px;">📝 具体修改内容与说明：</div>
                <div style="font-size: 14px; color: #1e293b; line-height: 1.6; white-space: pre-wrap;">{issue_info.get('changeContent', '无具体说明')}</div>
            </div>

            <div style="font-size: 12px; color: #94a3b8; margin-bottom: 20px;">
                📍 提交人 IP：{issue_info.get('clientIp', '未知')} ({issue_info.get('clientLocation', '')}) · 时间：{issue_info.get('submittedAt', '')}
            </div>

            <div style="background: #eff6ff; padding: 15px; border-radius: 12px; border: 1px solid #bfdbfe; text-align: center;">
                <p style="margin: 0 0 10px 0; font-size: 13px; color: #1e40af; font-weight: bold;">审核方式：在 GitHub Issue 中回复 <b>/approve</b> 即可全自动合入并给用户发送最新版 HTML！</p>
                <a href="{issue_info.get('issueUrl', 'https://github.com')}" style="display: inline-block; background: #2563eb; color: #ffffff; text-decoration: none; padding: 10px 24px; border-radius: 8px; font-weight: bold; font-size: 14px;">👉 打开 GitHub Issue 查看与审核</a>
            </div>
        </div>
        """

        msg.attach(MIMEText(html_content, 'html', 'utf-8'))
        server.sendmail(SMTP_USER, [ADMIN_EMAIL], msg.as_string())
        server.quit()
        print(f"✅ 管理员通知邮件已成功送达: {ADMIN_EMAIL}")
        return True
    except Exception as e:
        print(f"❌ 发送管理员通知邮件失败: {e}")
        return False

def send_user_approved_receipt(user_email, user_name, target_name, html_filepath):
    """当管理员审核通过后，给用户发送回执邮件并附带最新 HTML 附件"""
    if not user_email or '@' not in user_email:
        print("ℹ️ 用户未提供有效邮箱，跳过用户邮件发送。")
        return False

    try:
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=20)
        server.login(SMTP_USER, SMTP_PASS)

        msg = MIMEMultipart()
        msg['From'] = formataddr((str(Header('南江江氏宗族理事会', 'utf-8')), SMTP_USER))
        msg['To'] = formataddr((str(Header(user_name or '宗亲', 'utf-8')), user_email))
        msg['Subject'] = Header(f"【南江宗谱】您提交的族人信息修改已审核通过并生效！", 'utf-8')

        html_body = f"""
        <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 600px; margin: auto; padding: 25px; background: #f8fafc; border-radius: 16px; border: 1px solid #e2e8f0;">
            <h2 style="color: #059669; margin-top: 0;">🎉 尊敬的 {user_name or '宗亲'}：</h2>
            <p style="font-size: 14px; color: #334155; line-height: 1.8;">
                您向南江宗族理事会提交的关于【<b>{target_name}</b>】的信息修改/增补申请，已由宗族管理员<b>审核通过</b>并正式合入最新宗谱数据库！
            </p>

            <div style="background: #ecfdf5; padding: 15px; border-radius: 12px; border: 1px solid #a7f3d0; margin: 20px 0;">
                <div style="font-weight: bold; color: #065f46; font-size: 14px; margin-bottom: 8px;">📎 最新版本已作为邮件附件发送：</div>
                <div style="font-size: 13px; color: #047857;">
                    附件名称：<b>南江宗谱关系网（最新完美版）.html</b><br>
                    使用方法：下载附件后，直接在手机或电脑浏览器中双击即可<b>100% 离线高清浏览、查询两人及多人世系关系</b>！
                </div>
            </div>

            <p style="font-size: 13px; color: #64748b; line-height: 1.6;">
                修谱续帙，弘扬祖德。万分感谢您对南江江氏家族谱系完善所作出的宝贵贡献！
            </p>

            <div style="margin-top: 25px; padding-top: 15px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #94a3b8; text-align: center;">
                南江江氏宗族理事会 · 数字化谱牒管理中心
            </div>
        </div>
        """
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        # 添加 HTML 附件
        if os.path.exists(html_filepath):
            with open(html_filepath, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(html_filepath))
            part['Content-Disposition'] = f'attachment; filename="{Header(os.path.basename(html_filepath), "utf-8").encode()}"'
            msg.attach(part)
            print(f"📎 成功装载最新 HTML 附件: {html_filepath}")

        server.sendmail(SMTP_USER, [user_email], msg.as_string())
        server.quit()
        print(f"✅ 已成功给用户发送审核通过邮件及附件: {user_email}")
        return True
    except Exception as e:
        print(f"❌ 发送用户回执邮件失败: {e}")
        return False

if __name__ == "__main__":
    print("Genealogy Mailer module loaded.")
