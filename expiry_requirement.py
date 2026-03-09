# 營所指定效期需求
import os
import oracledb
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase          # ← 目前不附檔，先 mark 起來
# from email import encoders                    # ← 目前不附檔，先 mark 起來
from email.mime.text import MIMEText
from email.header import Header
from datetime import date
"""

"""

# ⚠ 「所指定效期需求.sql」路徑
SQL_PATH = r"D:\sql\營所指定效期需求.sql"
OUTPUT_DIR = r"D:\python\output"   # ← 預留 Excel 用路徑，現在先不用
SMTP_HOST = "192.168.2.127"
SMTP_PORT = 25
SENDER = "erpinfo@creation.com.tw"
# TO_LIST = ["timmy@creation.com.tw","yating@creation.com.tw"]   # 想多加人就自己加
TO_LIST = [
    "leon@creation.com.tw",
    "candy@creation.com.tw",
    "jamy@creation.com.tw",
    "maggie@creation.com.tw",
    "00839@creation.com.tw",
    "00852@creation.com.tw",
    "00937@creation.com.tw",
    "01210@creation.com.tw",
    "kissjin@creation.com.tw",
    "04326@creation.com.tw",
    "allenbse@creation.com.tw",
    "aireli@creation.com.tw",
    "sophiechang@creation.com.tw",
    "steveyeh@creation.com.tw",
    "stellyzheng@creation.com.tw",
    "max@creation.com.tw",
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

    with open(SQL_PATH, "r", encoding="cp950") as f:
        sql_query = f.read()

    with conn.cursor() as cur:
        cur.execute(sql_query)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]

    conn.close()
    return pd.DataFrame(rows, columns=cols)


# === 下面是「產 Excel」功能，現在先不用，但保留以後可開啟 ===
def export_excel(df):
    """
    【預留功能】把 df 輸出成 Excel 檔案。
    目前流程不用產 Excel / 不附檔案，
    之後若 user 要附件時，可以呼叫這個 function。
    """
    today = date.today().strftime("%Y%m%d")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    file_path = os.path.join(OUTPUT_DIR, f"營所指定效期需求_{today}.xlsx")
    df.to_excel(file_path, index=False)
    return file_path, today
# === 產 Excel 功能到這邊先結束 ===


def send_email_with_table(df):
    """
    將查詢結果直接用【HTML 表格】寫在 mail 內文裡（不附 Excel）
    - 欄位改中文標題、固定欄位順序
    - 表格置中、加框線
    """
    today_str = date.today().strftime("%Y/%m/%d")

    # ★ 欄位中文名稱 mapping（也順便固定顯示順序）
    col_mapping = {
        "需求日期": "需求日期",
        "料號": "料號",
        "品名": "品名",
        "批號": "批號",
        "數量": "數量",
        "時段": "時段",
        "原儲位": "原儲位",
        "新儲位": "新儲位",
        "儲位名稱": "儲位名稱",
        "備註7": "備註",
        "狀況": "狀況",
        "部門": "部門",
        "調儲人員工號": "調儲人員工號",
        "調儲人員名稱": "調儲人員名稱",
    }

    # 只取有在 mapping 裡的欄位，並依照 mapping 順序排列
    existing_cols = [c for c in col_mapping.keys() if c in df.columns]
    df_display = df[existing_cols].rename(columns=col_mapping)

    # ★ 日期欄位美化：20251212 -> 2025-12-12
    if "效期" in df_display.columns:
        df_display["效期"] = (
            df_display["效期"]
            .astype(str)
            # .str.replace(r"(\d{4})(\d{2})(\d{2})", r"\1-\2-\3", regex=True)
        )

    # ★ 數量欄位美化（可選）：轉成整數 / 千分位
    if "庫存數量" in df_display.columns:
        df_display["庫存數量"] = (
            pd.to_numeric(df_display["庫存數量"], errors="coerce")
            .fillna(0)
            # .astype(int) #1變成整數
        )

    # ★ 用 pandas 產 HTML 表格
    table_html = df_display.to_html(
        index=False,
        border=1,
        justify="center"
    )

    # ★ 美化 table 樣式（字型、置中、框線等）
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
    msg["Subject"] = Header(f"[開元]營所指定效期需求", "utf-8")

    # ★ HTML 內文
    body_html = f"""
<html>
  <body style="font-family:Calibri, 微軟正黑體; font-size:16px;">
    <p>您好，</p>
    <p>以下為營所指定效期需通知：</p>

    {table_html}

    <p style="margin-top:12px; color:#666;">
      ※ 此信為系統自動發送，請勿直接回覆。<br>
    </p>
  </body>
</html>
"""

    # ★ 用 HTML 格式寄出
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as server:
        all_recipients = TO_LIST + CC_LIST
        server.send_message(msg, from_addr=SENDER, to_addrs=all_recipients)


if __name__ == "__main__":
    try:
        df = run_query()

        # ⭐ 關鍵：沒有資料就什麼都不做（不寄信）
        if df.empty:
            print("✔ 工廠效期資料 0 筆 → 不寄信")
        else:
            print(f"✔ 工廠效期資料 {len(df)} 筆 → 寄信（HTML 表格＋中文欄位）")

            # 如果未來 user 想要 Excel 附件，再打開下面兩行就好：
            # file_path, today = export_excel(df)
            # （再加一個 send_email_with_attachment(file_path, df)）

            # 目前版本：只把資料寫進 mail 本文（HTML 表格）
            send_email_with_table(df)
            print("📨 成功寄出工廠有效期限屆滿通知")

    except Exception as e:
        print(f"❌ 錯誤：{e}")
        raise
