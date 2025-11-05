import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import requests
from io import BytesIO

# ------------------- 환경 변수 -------------------
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME = os.getenv("HIGHLIGHT_NAME")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO", GMAIL_USER)  # 받는사람 기본값 = 자기 자신

# ------------------- 셀레니움 설정 -------------------
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get(QOO10_URL)
WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".deal_lst li")))

# ------------------- 데이터 추출 -------------------
items = driver.find_elements(By.CSS_SELECTOR, ".deal_lst li")
data = []
for idx, item in enumerate(items, start=1):
    try:
        title = item.find_element(By.CSS_SELECTOR, ".item_tit").text.strip()
        price = item.find_element(By.CSS_SELECTOR, ".prc strong").text.strip()
        total = item.find_element(By.CSS_SELECTOR, ".sum").text.strip()
        img_url = item.find_element(By.CSS_SELECTOR, "img").get_attribute("src")
        data.append((idx, title, price, total, img_url))
    except Exception:
        continue

driver.quit()

# ------------------- 엑셀 저장 -------------------
now = datetime.now().strftime("%Y-%m-%d")
file_name = f"qoo10_ranking_{now}.xlsx"
df = pd.DataFrame(data, columns=["순위", "상품명", "가격", "판매총액", "이미지URL"])
df.to_excel(file_name, index=False)

# ------------------- 강조 표시 -------------------
wb = load_workbook(file_name)
ws = wb.active
highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
bold_font = Font(bold=True, color="000000")

for row in ws.iter_rows(min_row=2):
    if HIGHLIGHT_NAME and HIGHLIGHT_NAME in str(row[1].value):
        for cell in row:
            cell.fill = highlight_fill
            cell.font = bold_font

# ------------------- 이미지 삽입 -------------------
for i, (_, _, _, _, img_url) in enumerate(data, start=2):
    try:
        img_data = requests.get(img_url).content
        img = Image.open(BytesIO(img_data))
        img.thumbnail((80, 80))
        temp_path = f"temp_{i}.png"
        img.save(temp_path)
        excel_img = XLImage(temp_path)
        ws.add_image(excel_img, f"F{i}")
        os.remove(temp_path)
    except:
        continue

wb.save(file_name)

# ------------------- 이메일 발송 -------------------
msg = MIMEMultipart()
msg["From"] = GMAIL_USER
msg["To"] = SEND_TO
msg["Subject"] = f"Qoo10 랭킹 자동 보고서 - {now}"
body = f"{now}자 Qoo10 랭킹 엑셀 파일을 첨부합니다."
msg.attach(MIMEText(body, "plain"))

with open(file_name, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={file_name}")
    msg.attach(part)

with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(GMAIL_USER, GMAIL_PASS)
    server.send_message(msg)

print("[INFO] 자동 메일 전송 완료 ✅")
