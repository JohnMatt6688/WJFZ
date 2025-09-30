import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from io import BytesIO
from datetime import datetime

# ======================
# 配置区（从环境变量读取）
# ======================
SMTP_SERVER = "smtp.126.com"
SMTP_PORT = 465
EMAIL_ACCOUNT = os.environ.get("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")      # 直接使用授权码（明文，但由 GitHub Secrets 保护）
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL")
ENCODING = 'UTF-8'  # 苏州房管局网页实际编码

# 安全检查：确保 secrets 已设置
if not all([EMAIL_ACCOUNT, EMAIL_PASSWORD, RECIPIENT_EMAIL]):
    raise EnvironmentError(
        "缺失环境变量！请在 GitHub Secrets 中设置：EMAIL_ACCOUNT, EMAIL_PASSWORD, RECIPIENT_EMAIL"
    )

def debug_log(message):
    """ 调试日志输出 """
    print(f"[DEBUG] {message}")

def fetch_web_data():
    """ 修复编码问题的数据抓取 """
    try:
        response = requests.get(
            "http://clf.zfcjj.suzhou.gov.cn/xsinfo.aspx",
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            },
            timeout=15
        )
        response.raise_for_status()
        soup = BeautifulSoup(response.content.decode(ENCODING), 'html.parser')
        table = soup.find('table', id='ctl00_ContentPlaceHolder1_mytable')
        
        if table is None:
            debug_log("未找到目标表格，可能网页结构已变更")
            return pd.DataFrame()
        
        raw_data = []
        current_region = None
        rowspan_left = 0

        for row in table.find_all('tr')[1:]:
            cells = row.find_all(['th', 'td'])
            if not cells:
                continue
            # 处理区域 rowspan
            if cells[0].name == 'th' and 'rowspan' in cells[0].attrs:
                current_region = cells[0].get_text(strip=True)
                rowspan_left = int(cells[0]['rowspan']) - 1
            elif rowspan_left > 0:
                rowspan_left -= 1

            if len(cells) >= 4:
                record = {
                    '区域': current_region,
                    '类型': cells[1].get_text(strip=True),
                    '套数': cells[2].get_text(strip=True),
                    '面积': cells[3].get_text(strip=True)
                }
                raw_data.append(record)
            elif len(cells) == 3:
                record = {
                    '区域': current_region,
                    '类型': cells[0].get_text(strip=True),
                    '套数': cells[1].get_text(strip=True),
                    '面积': cells[2].get_text(strip=True)
                }
                raw_data.append(record)

        return pd.DataFrame(raw_data)
    except Exception as e:
        debug_log(f"抓取失败: {str(e)}")
        return pd.DataFrame()

def process_data(raw_df):
    """ 数据处理 """
    if raw_df.empty:
        return pd.DataFrame()
    
    processed = []
    current_group = []
    for _, row in raw_df.iterrows():
        if row['类型'] in ['小计', '总计']:
            if current_group:
                processed.append(current_group)
            current_group = [row]
        else:
            current_group.append(row)
    if current_group:
        processed.append(current_group)
    
    result = []
    for group in processed:
        if len(group) >= 2:
            main_row = group[0]
            house_row = group[1]
            result.append({
                '区域': main_row['区域'],
                '小计套数': main_row['套数'],
                '小计面积': main_row['面积'],
                '住宅套数': house_row['套数'],
                '住宅面积': house_row['面积']
            })
    return pd.DataFrame(result)

def send_email_with_excel(df):
    """ 发送Excel文件到邮箱，并在邮件正文附带数据 """
    date_str = datetime.now().strftime("%m-%d")
    filename = f"成交-{date_str}.xlsx"
    
    with BytesIO() as buffer:
        df.to_excel(buffer, index=False, sheet_name=date_str)
        buffer.seek(0)
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ACCOUNT
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = f"苏州市成交数据 - {date_str}"
        
        # 添加Excel附件
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(buffer.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)
        
        # 添加邮件正文（前10行数据）
        preview_df = df.head(10)
        text_content = f"苏州市成交数据 - {date_str}\n\n" + preview_df.to_string(index=False)
        msg.attach(MIMEText(text_content, 'plain', 'utf-8'))
        
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
                server.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
                server.sendmail(EMAIL_ACCOUNT, RECIPIENT_EMAIL, msg.as_string())
            print(f"✅ 邮件已发送至 {RECIPIENT_EMAIL}")
        except Exception as e:
            print(f"❌ 邮件发送失败: {str(e)}")

if __name__ == "__main__":
    print("=== 苏州市房产成交数据抓取开始 ===")
    raw_df = fetch_web_data()
    if not raw_df.empty:
        processed_df = process_data(raw_df)
        if not processed_df.empty:
            send_email_with_excel(processed_df)
        else:
            print("⚠️ 数据处理后为空，未发送邮件")
    else:
        print("❌ 数据获取失败，请检查网络或网页结构是否变更")
