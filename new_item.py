# 新品建檔通知（參照 e_invoice_remaining_notify.py 架構）
import os
import oracledb
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from datetime import date

"""
用途：
1) 執行 Oracle SQL（新品建檔通知.sql）
2) 查到資料 → 以 HTML 表格寄信（不附 Excel）
3) 沒資料 → 不寄信
"""


SQL_PATH = r"D:\sql\新品建檔通知.sql"

OUTPUT_DIR = r"D:\python\output"   # 預留 Excel 用路徑（目前不使用）
SMTP_HOST = "192.168.2.127"
SMTP_PORT = 25
SENDER = "erpinfo@creation.com.tw"

# 收件人/副本
TO_LIST = [
    "candy@creation.com.tw",
    "00852@creation.com.tw",
    "04326@creation.com.tw",
    "allenbse@creation.com.tw",
    "regina@creation.com.tw",
    "chunchieh@creation.com.tw",
    "max@creation.com.tw",
    "stellyzheng@creation.com.tw",
]

CC_LIST = [
    "#-801645753@creation.com.tw"
]

TIMEOUT = 15


def run_query():
    """執行 SQL，回傳 DataFrame（不產 Excel）"""
    oracledb.init_oracle_client(lib_dir=r"D:\instantclient_11_2")

    conn = oracledb.connect(
        user="mf2000",
        password="yrrah",
        dsn="192.168.2.7:1521/ORCL"
    )

    # SQL 內有中文 → cp950 最穩；怕遇到怪字就用 errors="ignore"
    with open(SQL_PATH, "r", encoding="cp950", errors="ignore") as f:
        sql_query = f.read()

    with conn.cursor() as cur:
        cur.execute(sql_query)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]

    conn.close()
    return pd.DataFrame(rows, columns=cols)

def format_unit(row):
    """
    將下層/上層數量與單位組成：
    24 BAG / 1 BX
    """
    if pd.notna(row.get("下層數量")) and pd.notna(row.get("上層數量")):
        return f"{int(row['下層數量'])} {row['下層單位']} / {int(row['上層數量'])} {row['上層單位']}"
    else:
        return "-"   # 沒有資料就顯示 -

# 新增一個顯示用欄位


def send_email_with_table(df: pd.DataFrame):
    """
    將查詢結果直接用【HTML 表格】寫在 mail 內文裡（不附 Excel）
    - 欄位中文標題、固定欄位順序
    """
    today_str = date.today().strftime("%Y/%m/%d")
    df["單位換算"] = df.apply(format_unit, axis=1)


    col_mapping = {
        "來源": "來源",
        "料號": "料號",
        "品名": "品名",
        "單位換算":"單位換算",
        "食品雲": "食品雲",
        "批號": "批號",
    }

    # 只取我們要顯示的欄位（固定順序）
    show_cols = [c for c in col_mapping.keys() if c in df.columns]
    df_show = df[show_cols].rename(columns=col_mapping)

    # ★ 用 pandas 產 HTML 表格
    table_html = df_show.to_html(index=False, border=1, justify="center")

    # ★ 美化 table 樣式
    table_html = table_html.replace(
        '<table border="1" class="dataframe">',
        '<table border="1" cellspacing="0" cellpadding="4" '
        'style="border-collapse:collapse;'
        'font-family:Calibri, 微軟正黑體;'
        'font-size:16px; text-align:center;">'
    )

    msg = MIMEMultipart()
    msg["From"] = SENDER
    msg["To"] = ", ".join(TO_LIST)
    msg["Cc"] = ", ".join(CC_LIST)
    msg["Subject"] = Header(f"[開元] 新品建檔通知 {today_str}", "utf-8")

    body_html = f"""
<html>
  <body style="font-family:Calibri, 微軟正黑體; font-size:16px;">
    <p>您好，</p>
    <p>以下為昨日 <b>{today_str}</b> 新品建檔清單：</p>

    {table_html}

    <p style="margin-top:12px; color:#666;">
      ※ 此信為系統自動發送，請勿直接回覆。<br>
    </p>
  </body>
</html>
"""
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as server:
        all_recipients = TO_LIST + CC_LIST
        server.send_message(msg, from_addr=SENDER, to_addrs=all_recipients)


if __name__ == "__main__":
    try:
        df = run_query()

        # ⭐ 沒資料就不寄信
        if df.empty:
            print("✔ 新品建檔資料 0 筆 → 不寄信")
        else:
            print(f"✔ 新品建檔資料 {len(df)} 筆 → 寄信（HTML 表格＋中文欄位）")
            send_email_with_table(df)
            print("📨 成功寄出新品建檔通知")

    except Exception as e:
        print(f"❌ 錯誤：{e}")
        raise
