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
from datetime import datetime


# ==============================================================
# ✅ print() 로그에 자동 시간 붙이기
# ==============================================================
import builtins
_original_print = builtins.print

def timestamped_print(*args, **kwargs):
    now = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    _original_print(now, *args, **kwargs)

builtins.print = timestamped_print


# ==============================================================
# ✅ GitHub Secrets 로딩
# ==============================================================
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME = os.getenv("HIGHLIGHT_NAME", "メガワリ")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO")


# ==============================================================
# ✅ Headless Chrome 설정
# ==============================================================
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


# ==============================================================
# ✅ Qoo10 접속 + iframe 처리
# ==============================================================
def switch_to_last_iframe(driver):
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"[INFO] 감지된 iframe 수: {len(iframes)}")

    for i, f in enumerate(iframes):
        try:
            driver.switch_to.frame(f)
            print(f"[INFO] iframe {i+1} 전환 성공")

            # 내부에 iframe이 또 있을 경우 반복
            inner = f.find_elements(By.TAG_NAME, "iframe")
            if inner:
                driver.switch_to.frame(inner[0])
                print("[INFO] 내부 iframe 추가 전환 완료")
            return
        except:
            continue

    print("[WARN] iframe 전환 실패 — 메인 페이지에서 처리")


# ==============================================================
# 🧩 페이지 구조 자동 감지
# ==============================================================
def detect_page_mode(driver):
    html = driver.page_source

    # 메가와리 누적금액순 탭 특징
    if "best-accum-price" in html or "accu-price" in html:
        print("[INFO] 메가와리 누적金額順 페이지 감지됨")
        return "megawari_amount"

    # 구버전 메가와리
    if "megasale_rank_list" in html:
        print("[INFO] 구버전 메가와리 페이지 감지됨")
        return "legacy"

    print("[WARN] 페이지 구조를 자동 감지하지 못함")
    return "unknown"


# ==============================================================
# 🧩 메가와리 누적금액순 전용 파서
# ==============================================================
def parse_megawari_amount(driver):
    results = []

    # 메가와리 이벤트 구조 (버전별 대응)
    items = driver.find_elements(
        By.CSS_SELECTOR,
        "div.best-item, li.best-list-item, div.product_item"
    )

    print(f"[INFO] 감지된 아이템 수: {len(items)}")

    for item in items[:100]:
        # 순위
        try:
            rank = item.find_element(By.CSS_SELECTOR, ".rank-num, .best-rank").text.strip()
        except:
            rank = ""

        # 상품명
        try:
            name = item.find_element(
                By.CSS_SELECTOR,
                ".item-title, .best-title, .text-elps"
            ).text.strip()
        except:
            name = ""

        # 가격
        try:
            price = item.find_element(
                By.CSS_SELECTOR,
                ".price__value, .price--discount"
            ).text.strip()
        except:
            price = ""

        # 판매총액
        try:
            amount = item.find_element(
                By.CSS_SELECTOR,
                ".best-accum-price, .accu-price"
            ).text.strip()
        except:
            amount = ""

        # 이미지
        try:
            img_el = item.find_element(By.CSS_SELECTOR, "img")
            image = img_el.get_attribute("data-src") or img_el.get_attribute("src")
        except:
            image = ""

        results.append([rank, name, price, amount, image])

    return results


# ==============================================================
# 🧩 구버전 메가와리/베스트셀러 파서 (호환 유지)
# ==============================================================
def parse_legacy(driver):
    data = []
    items = driver.find_elements(By.CSS_SELECTOR, "ul.megasale_rank_list li")

    print(f"[INFO] 구버전 아이템 수: {len(items)}")

    for p in items[:100]:
        try:
            rank = p.find_element(By.CSS_SELECTOR, ".rank_num").text.strip()
            name = p.find_element(By.CSS_SELECTOR, ".title").text.strip()
            price = p.find_element(By.CSS_SELECTOR, ".price").text.strip()
            total = p.find_element(By.CSS_SELECTOR, ".value").text.strip()
            img = p.find_element(By.CSS_SELECTOR, ".thumb img").get_attribute("src")
            data.append([rank, name, price, total, img])
        except:
            continue

    return data


# ==============================================================
# 🎯 자동 구조 감지 → 해당 파서 실행
# ==============================================================
mode = detect_page_mode(driver)

if mode == "megawari_amount":
    data = parse_megawari_amount(driver)
elif mode == "legacy":
    data = parse_legacy(driver)
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
    ws.append(row[:-1])  # 이미지 제외


# 🎯 강조 (상품명에 HIGHLIGHT_NAME 포함 시)
for row in ws.iter_rows(min_row=2, max_col=4):
    if HIGHLIGHT_NAME in str(row[1].value):
        for cell in row:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = Font(bold=True, color="000000")


# ==============================================================
# 🖼 이미지 삽입
# ==============================================================
for i, row in enumerate(data, start=2):
    img_url = row[4]
    try:
        img_data = requests.get(img_url, timeout=10).content
        image = Image.open(BytesIO(img_data))
        image.thumbnail((80, 80))
        bio = BytesIO()
        image.save(bio, format="PNG")
        bio.seek(0)

        img = XLImage(bio)
        ws.add_image(img, f"E{i}")
        time.sleep(0.1)
    except Exception as e:
        print(f"[WARN] 이미지 실패: {e}")


# ==============================================================
# 💾 저장
# ==============================================================
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
    f"안녕하세요,\n\n자동 생성된 Qoo10 {HIGHLIGHT_NAME} 누적판매금액순 보고서입니다.\n"
    f"생성일자: {today}\n"
    f"참조 URL: {QOO10_URL}\n\n사랑해용.",
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

print("[INFO] 자동 메일 전송 완료 ✅")
