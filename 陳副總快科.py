import os
import oracledb
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from email.header import Header
from datetime import date

SQL_PATH = r"\\192.168.2.26\部門-資訊部\個人_林宏陽\SQL\常用\雅雲塊科.sql"
OUTPUT_DIR = r"D:\\"
SMTP_HOST = "192.168.2.127"
SMTP_PORT = 25
SENDER = "timmy@creation.com.tw"
TO_LIST = ["timmy@creation.com.tw", "yayun@creation.com.tw"]  # ←加入雅雲
# CC_LIST = ["someone@creation.com.tw"]  # 需要再開
TIMEOUT = 15

def export_excel():
    oracledb.init_oracle_client(lib_dir=r"D:\instantclient_11_2")
    conn = oracledb.connect(
        user="mf2000",
        password="yrrah", 
        dsn="192.168.2.7:1521/ORCL"
    )
    with open(SQL_PATH, "r", encoding="cp950") as f:
        sql_query = f.read()

    with conn.cursor() as cur:
        cur.execute(sql_query)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
    conn.close()

    df = pd.DataFrame(rows, columns=cols)
    today = date.today().strftime("%Y%m%d")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_file = os.path.join(OUTPUT_DIR, f"雅雲塊科_{today}.xlsx")

    # 輸出 Excel（index=False：不要把 DataFrame 的 index 寫進 Excel）
    df.to_excel(output_file, index=False)
    return output_file, today

def send_email(file_path, today):
    msg = MIMEMultipart()
    msg["From"] = SENDER
    msg["To"] = ", ".join(TO_LIST)
    # msg["Cc"] = ", ".join(CC_LIST) if CC_LIST else ""
    msg["Subject"] = "陳副總塊科傳票查詢(自動發信)"
    msg.attach(MIMEText("您好，附件為雅雲塊科的 Excel 報表，請查收。", "plain"))

    with open(file_path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition", "attachment",
            filename=Header(f"雅雲塊科_{today}.xlsx", "utf-8").encode()
        )
        msg.attach(part)

    all_rcpts = TO_LIST  # + CC_LIST
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as server:
        server.send_message(msg, from_addr=SENDER, to_addrs=all_rcpts)

if __name__ == "__main__":
    try:
        xlsx, today = export_excel()
        send_email(xlsx, today)
        print("✅ 成功：已產生並寄出報表")
    except Exception as e:
        print(f"❌ 失敗：{e}")
        raise
