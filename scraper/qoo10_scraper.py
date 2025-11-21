import os
import time
import smtplib
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import requests
from datetime import datetime


# ==============================================================
# 🕒 Timestamped Print
# ==============================================================
import builtins
_original_print = builtins.print

def timestamped_print(*args, **kwargs):
    now = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    _original_print(now, *args, **kwargs)

builtins.print = timestamped_print


# ==============================================================
# 🔐 GitHub Secrets
# ==============================================================
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME = os.getenv("HIGHLIGHT_NAME", "メガワリ")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO")


# ==============================================================
# 🖥 Headless Chrome 설정
# ==============================================================
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)


# ==============================================================
# 🔍 페이지 구조 자동 감지
# ==============================================================
def detect_page_mode(driver):
    html = driver.page_source

    # 👉 너가 제공한 실제 HTML 구조 기준:
    # <ul class="megasale_rank_list"> ... </ul>
    if 'megasale_rank_list' in html:
        print("[INFO] 최신 메가와리 리스트 페이지 감지됨")
        return "megawari_list"

    print("[WARN] 페이지 구조 자동 감지 실패")
    return "unknown"


# ==============================================================
# 🧩 메가와리 리스트 파서 (너가 제공한 HTML 기반)
# ==============================================================
def parse_megawari_list(driver):
    results = []

    items = driver.find_elements(
        By.CSS_SELECTOR,
        "ul.megasale_rank_list > li"
    )

    print(f"[INFO] 감지된 상품 수: {len(items)}")

    for item in items[:100]:

        try:
            rank = item.find_element(By.CSS_SELECTOR, ".rank_num").text.strip()
        except:
            rank = ""

        try:
            name = item.find_element(By.CSS_SELECTOR, ".title span").text.strip()
        except:
            name = ""

        try:
            price = item.find_element(By.CSS_SELECTOR, ".price").text.strip()
        except:
            price = ""

        try:
            img = item.find_element(By.CSS_SELECTOR, ".thumb img").get_attribute("src")
        except:
            img = ""

        # 메가와리 리스트 구조에는 판매총액 없음 → 빈칸
        results.append([rank, name, price, "", img])

    return results


# ==============================================================
# 🚀 Qoo10 접속
# ==============================================================
print(f"[INFO] Qoo10 접속: {QOO10_URL}")
driver.get(QOO10_URL)
time.sleep(5)


# ==============================================================
# 🎯 구조 감지 후 파싱 실행
# ==============================================================
mode = detect_page_mode(driver)

if mode == "megawari_list":
    data = parse_megawari_list(driver)
else:
    print("[ERROR] 페이지 구조 지원 불가 → 종료")
    driver.quit()
    raise SystemExit

driver.quit()


# ==============================================================
# 📘 엑셀 생성
# ==============================================================
wb = Workbook()
ws = wb.active
ws.title = "Qoo10 Ranking"

ws.append(["순위", "상품명", "가격", "판매총액", "이미지"])

for row in data:
    ws.append(row[:-1])


# 🎯 강조 (상품명에 HIGHLIGHT_NAME 포함 시)
for row in ws.iter_rows(min_row=2, max_col=4):
    if HIGHLIGHT_NAME in str(row[1].value):
        for cell in row:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = Font(bold=True, color="000000")


# ==============================================================
# 🖼 이미지 삽입
# ==============================================================
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/118.0.5993.70 Safari/537.36"
}

for i, row in enumerate(data, start=2):
    img_url = row[4]

    if not img_url:
        continue

    # 1) URL이 // 로 시작하면 https: 붙여주기
    if img_url.startswith("//"):
        img_url = "https:" + img_url

    try:
        # 2) 헤더를 붙여서 403 방지
        resp = requests.get(img_url, headers=headers, timeout=10)
        resp.raise_for_status()
        img_bytes = resp.content

        # 3) WebP 또는 기타 포맷을 PNG로 통일
        image = Image.open(BytesIO(img_bytes))
        image = image.convert("RGB")   # WebP → RGB 변환
        image.thumbnail((80, 80))

        bio = BytesIO()
        image.save(bio, format="PNG")
        bio.seek(0)

        img = XLImage(bio)
        ws.add_image(img, f"E{i}")
        time.sleep(0.2)

    except Exception as e:
        print(f"[WARN] 이미지 실패: {img_url} → {e}")

file_name = "Qoo10_Rank.xlsx"
wb.save(file_name)
print(f"[INFO] 엑셀 저장 완료: {file_name}")


# ==============================================================
# 📧 이메일 전송
# ==============================================================
msg = MIMEMultipart()
msg["From"] = GMAIL_USER
msg["To"] = SEND_TO

today = datetime.now().strftime("%Y-%m-%d")
msg["Subject"] = f"Qoo10 랭킹 자동 보고서 {today}"

body = MIMEText(
    f"안녕하세요,\n\n자동 생성된 Qoo10 {HIGHLIGHT_NAME} 랭킹 보고서입니다.\n"
    f"생성일자: {today}\n"
    f"URL: {QOO10_URL}\n\n좋은 하루 보내세요!",
    "plain"
)
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

print("[INFO] 자동 메일 전송 완료! 🚀🔥")
