import os
import time
import smtplib
import pandas as pd
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import requests

# -----------------------------
# 환경변수 불러오기 (GitHub Secrets)
# -----------------------------
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME = os.getenv("HIGHLIGHT_NAME", "メガ割")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO")

# -----------------------------
# Headless Chrome 설정
# -----------------------------
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# -----------------------------
# Qoo10 페이지 접속
# -----------------------------
print(f"[INFO] Qoo10 페이지 접속: {QOO10_URL}")
driver.get(QOO10_URL)
time.sleep(10)

# iframe 전환
try:
    iframe = driver.find_element(By.TAG_NAME, "iframe")
    driver.switch_to.frame(iframe)
    print("[INFO] iframe 전환 완료")
except:
    print("[WARN] iframe 없음 — 메인 페이지에서 탐색 진행")

# -----------------------------
# 상품 데이터 수집 함수
# -----------------------------
def get_product_elements():
    for attempt in range(3):  # 최대 3회 시도
        try:
            WebDriverWait(driver, 180).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.megasale_rank_list li"))
            )
            print(f"[INFO] 상품 목록 로딩 성공 (시도 {attempt + 1})")
            return driver.find_elements(By.CSS_SELECTOR, "ul.megasale_rank_list li")
        except TimeoutException:
            print(f"[WARN] {180}초 대기 후 실패 (시도 {attempt + 1}) → 새로고침")
            driver.refresh()
            time.sleep(15)
    print("[ERROR] 상품 목록 로딩 실패")
    return []

products = get_product_elements()

# -----------------------------
# 상품 데이터 추출
# -----------------------------
data = []
for p in products[:100]:
    try:
        rank = p.find_element(By.CSS_SELECTOR, ".rank_num").text.strip()
        name = p.find_element(By.CSS_SELECTOR, ".title").text.strip()
        price = p.find_element(By.CSS_SELECTOR, ".price").text.strip()
        total = p.find_element(By.CSS_SELECTOR, ".value").text.strip()
        img = p.find_element(By.CSS_SELECTOR, ".thumb img").get_attribute("src")
        data.append([rank, name, price, total, img])
    except Exception as e:
        print(f"[WARN] 상품 정보 파싱 실패: {e}")
        continue

driver.quit()

# -----------------------------
# 엑셀 파일 생성
# -----------------------------
wb = Workbook()
ws = wb.active
ws.title = "Qoo10 Top 100"
ws.append(["순위", "상품명", "가격", "판매총액", "이미지"])

for row in data:
    ws.append(row[:-1])  # 이미지 제외

# -----------------------------
# 강조 표시 (HIGHLIGHT_NAME 포함 시)
# -----------------------------
for row in ws.iter_rows(min_row=2, max_col=4):
    if HIGHLIGHT_NAME in str(row[1].value):
        for cell in row:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = Font(bold=True, color="000000")

# -----------------------------
# 이미지 삽입 (메모리 기반 + 안전 대기)
# -----------------------------
for i, row in enumerate(data, start=2):
    img_url = row[4]
    try:
        img_data = requests.get(img_url, timeout=15).content
        image = Image.open(BytesIO(img_data))
        image.thumbnail((80, 80))
        bio = BytesIO()
        image.save(bio, format="PNG")
        bio.seek(0)
        img = XLImage(bio)
        ws.add_image(img, f"E{i}")
        time.sleep(0.2)  # GitHub Actions용 안전 대기
    except Exception as e:
        print(f"[WARN] 이미지 처리 실패: {e}")
        continue

# -----------------------------
# 엑셀 저장
# -----------------------------
file_name = "Qoo10_Rank.xlsx"
wb.save(file_name)
print(f"[INFO] 엑셀 저장 완료: {file_name}")

# -----------------------------
# 이메일 전송
# -----------------------------
msg = MIMEMultipart()
msg["From"] = GMAIL_USER
msg["To"] = SEND_TO
msg["Subject"] = f"[Qoo10 자동 리포트] {HIGHLIGHT_NAME} Top 100 결과"

body = MIMEText(f"자동으로 생성된 Qoo10 {HIGHLIGHT_NAME} 순위 엑셀 파일입니다.\n\nURL: {QOO10_URL}", "plain")
msg.attach(body)

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