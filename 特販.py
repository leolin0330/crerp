# -*- coding: utf-8 -*-

# ====== 1) 匯入套件 ======
# os：用來處理檔案路徑、建立資料夾等作業系統相關功能
import os

# oracledb：Python 連線 Oracle 的驅動（Thin / Thick 模式）
import oracledb

# pandas：用來把查詢結果轉成 DataFrame，並輸出 Excel
import pandas as pd

# smtplib：Python 內建的 SMTP 寄信模組
import smtplib

# email.*：組合 Email 主體與附件需要的元件
from email.mime.multipart import MIMEMultipart  # 多段式信件（可以同時有文字與附件）
from email.mime.base import MIMEBase           # 附件的基底型別
from email import encoders                     # 附件需要 base64 編碼
from email.mime.text import MIMEText           # 純文字信件內容
from email.header import Header                # 用來正確處理附件中文檔名
from datetime import date


# ====== 2) SQL：直接寫在 Python 裡（單檔版的核心） ======
# r"""...""" 使用 raw string：

SQL_QUERY = r"""
SELECT
lcst.cod_item 料號,item.nam_item 料號名稱,l.content 儲位形式代號,lcst.cod_loc 儲位,loct.nam_loc 儲位名稱,LCST.QTY_STK 儲位存量,u.content 單位,lcst.ser_pcs 批號
FROM loct ,item ,lcst,codd l ,codd u
 WHERE item.cod_item = lcst.cod_item AND loct.cod_loc = lcst.cod_loc  and LCST.QTY_STK > 0
and l.code_id = 'CLSSIZE' and l.code = loct.cls_size
and u.code_id = 'CODUNI' and u.code = lcst.unt_stk
and lcst.cod_item in(
'N0010001',
'F0250022',
'J0090001',
'J0090003',
'V9010J18',
'F0260000',
'J0100002',
'S0370009',
'V3200002',
'V3200003',
'P019G000',
'A024013T',
'A024033T',
'A003050T',
'Q1800000',
'A025010T',
'Q1200000',
'S1806000'
)

"""


# ====== 3) 可調整的設定區（路徑 / 寄信 / 收件人） ======

# Excel 產出
OUTPUT_DIR = r"D:\\"

# SMTP 伺服器設定（公司內部 SMTP）
SMTP_HOST = "192.168.2.127"
SMTP_PORT = 25

# 寄件者（From）
SENDER = "erpinfo@creation.com.tw"

# 收件者清單（To）
TO_LIST = [
    "timmy@creation.com.tw",
    "vickylin@creation.com.tw",
]

CC_LIST = [
    "timmy@reation.com.tw",
    "yating@creation.com.tw"
]

# SMTP 連線逾時秒數（避免卡死）
TIMEOUT = 15


# ====== 4) 查詢 Oracle → 輸出 Excel ======
def export_excel():
    """執行 SQL，將結果輸出成 Excel，回傳（檔案路徑, today字串）。"""

    oracledb.init_oracle_client(lib_dir=r"D:\instantclient_11_2")

    # 建立 Oracle 連線
    # dsn：資料庫位址/服務（host:port/service_name）
    conn = oracledb.connect(
        user="mf2000",
        password="yrrah",
        dsn="192.168.2.7:1521/ORCL"
    )

    # 直接用內嵌 SQL，不再讀 .sql 檔
    sql_query = SQL_QUERY

    # 開 cursor 執行 SQL（用 with 可確保 cursor 正常關閉）
    with conn.cursor() as cur:
        # 送出 SQL 給 Oracle 執行
        cur.execute(sql_query)

        # 把全部資料抓回來（list[tuple]）
        rows = cur.fetchall()

        # 取得欄位名稱（cursor.description 每個欄位一個描述 tuple）
        cols = [d[0] for d in cur.description]

    # 關閉 DB 連線（避免資源占用）
    conn.close()

    # 把結果轉成 DataFrame，欄名用 cols（會是你 SQL 別名/欄位名）
    df = pd.DataFrame(rows, columns=cols)

    # today 字串：用於檔名，例如 20260203
    today = date.today().strftime("%Y%m%d")

    # 確保輸出資料夾存在；不存在就建立
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Excel 檔案完整路徑
    output_file = os.path.join(OUTPUT_DIR, f"特定料號庫存_{today}.xlsx")

    # 輸出 Excel（index=False：不要把 DataFrame 的 index 寫進 Excel）
    df.to_excel(output_file, index=False)

    # 回傳：給寄信函式使用
    return output_file, today


# ====== 5) 寄信（附上 Excel 檔） ======
def send_email(file_path, today):
    """寄出一封信並附上 Excel 檔案。"""

    # 建立多段式 Email（可同時含文字 + 附件）
    msg = MIMEMultipart()

    # 信件基本欄位
    msg["From"] = SENDER
    msg["To"] = ", ".join(TO_LIST)
    msg["Cc"] = ", ".join(CC_LIST)
    msg["Subject"] = "特定料號庫存(自動發信)"

    # 信件內文（plain：純文字）
    msg.attach(MIMEText("您好，附件為特定料號庫存的 Excel 報表，請查收。", "plain"))

    # 讀取檔案內容（binary）
    with open(file_path, "rb") as f:
        # 建立附件物件，指定 MIME 類型
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 把檔案內容放進附件 payload
        part.set_payload(f.read())

        # 附件需要 base64 編碼，信件傳輸才安全
        encoders.encode_base64(part)

        # 設定附件檔名（Header 用 utf-8，避免中文檔名亂碼）
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=Header(f"特定料號庫存_{today}.xlsx", "utf-8").encode()
        )

        # 把附件加進信件
        msg.attach(part)

    # 建立 SMTP 連線並送信
    # with 會自動 close 連線
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as server:
        # send_message：直接送出組好的 msg
        # from_addr / to_addrs 明確指定寄件者與收件者
        all_recipients = TO_LIST + CC_LIST
        server.send_message(msg, from_addr=SENDER, to_addrs=all_recipients)


# ====== 6) 程式進入點 ======
if __name__ == "__main__":
    try:
        # 1) 先查資料並輸出 Excel
        xlsx, today = export_excel()

        # 2) 再把 Excel 當附件寄出
        send_email(xlsx, today)

        # 3) console 顯示成功訊息（方便排程log判斷）
        print("✅ 成功：已產生並寄出報表")

    except Exception as e:
        # 任一步驟出錯都會進到這裡
        print(f"❌ 失敗：{e}")

        # raise 讓排程/監控可以抓到「非 0」的錯誤狀態
        raise
