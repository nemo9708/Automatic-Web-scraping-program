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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import requests
from datetime import datetime, timedelta


# ============================================================== 
# 🕒 Timestamped Print (JST)
# ============================================================== 
import builtins
_original_print = builtins.print

def timestamped_print(*args, **kwargs):
    now = (datetime.utcnow() + timedelta(hours=9)).strftime("[%Y-%m-%d %H:%M:%S]")
    _original_print(now, *args, **kwargs)

builtins.print = timestamped_print


# ============================================================== 
# 🔐 GitHub Secrets
# ============================================================== 
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME1 = os.getenv("HIGHLIGHT_NAME1")
HIGHLIGHT_NAME2 = os.getenv("HIGHLIGHT_NAME2")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO")


# ============================================================== 
# 🖥 Headless Chrome 설정
# ============================================================== 
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)


# ============================================================== 
# ⏳ 메가와리 리스트 로딩 대기
# ============================================================== 
def wait_list(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "ul.megasale_rank_list > li")
            )
        )
        print("[INFO] 메가와리 랭킹 리스트 로딩 완료")
        return True
    except:
        print("[WARN] 메가와리 리스트 감지 실패")
        return False


# ============================================================== 
# 🧩 메가와리 파서 (이미지 개선 포함)
# ============================================================== 
def parse_megawari(driver):
    results = []
    items = driver.find_elements(By.CSS_SELECTOR, "ul.megasale_rank_list > li")
    print(f"[INFO] 감지된 아이템 수: {len(items)}")

    for item in items[:100]:

        try:
            rank = item.find_element(By.CSS_SELECTOR, ":scope .rank_num").text.strip()
        except:
            rank = ""

        try:
            name = item.find_element(By.CSS_SELECTOR, ":scope .title span").text.strip()
        except:
            name = ""

        try:
            total = item.find_element(By.CSS_SELECTOR, ":scope .price").text.strip()
        except:
            total = ""

        try:
            img_el = item.find_element(
                By.CSS_SELECTOR,
                ":scope .thumb_area img, :scope .thumb img"
            )
            img_url = (
                img_el.get_attribute("data-src")
                or img_el.get_attribute("data-original")
                or img_el.get_attribute("data-lazy")
                or img_el.get_attribute("src")
            )
        except:
            img_url = ""

        results.append([rank, name, total, img_url])

    return results


# ============================================================== 
# 🚀 Qoo10 접속
# ============================================================== 
print(f"[INFO] Qoo10 접속: {QOO10_URL}")
driver.get(QOO10_URL)

wait_list(driver)
data = parse_megawari(driver)
driver.quit()


# ============================================================== 
# 📘 엑셀 생성
# ============================================================== 
wb = Workbook()
ws = wb.active
ws.title = "Qoo10 Ranking"

ws.append(["순위", "상품명", "총판매액", "이미지"])

for row in data:
    ws.append(row[:-1])


# ============================================================== 
# 🎨 하이라이트 적용 (2개 키워드)
# ============================================================== 
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

for row in ws.iter_rows(min_row=2, max_col=3):
    name = str(row[1].value)

    if HIGHLIGHT_NAME1 and HIGHLIGHT_NAME1 in name:
        for cell in row:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = Font(bold=True)

    if HIGHLIGHT_NAME2 and HIGHLIGHT_NAME2 in name:
        for cell in row:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = Font(bold=True)


# ============================================================== 
# 🖼 이미지 삽입
# ============================================================== 
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

for i, row in enumerate(data, start=2):
    img_url = row[3]
    if not img_url:
        continue

    if img_url.startswith("//"):
        img_url = "https:" + img_url

    try:
        resp = requests.get(img_url, headers=headers, timeout=10)
        resp.raise_for_status()
        img_bytes = resp.content

        image = Image.open(BytesIO(img_bytes))

        if image.mode in ("RGBA", "P"):
            bg = Image.new("RGB", image.size, (255, 255, 255))
            try:
                bg.paste(image, mask=image.split()[3])
            except:
                bg.paste(image)
            image = bg
        else:
            image = image.convert("RGB")

        image.thumbnail((80, 80))

        bio = BytesIO()
        image.save(bio, format="PNG")
        bio.seek(0)

        img = XLImage(bio)
        ws.add_image(img, f"D{i}")

        time.sleep(0.15)

    except Exception as e:
        print(f"[WARN] 이미지 실패: {img_url} → {e}")


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

today = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d")
msg["Subject"] = f"Qoo10 랭킹 자동 보고서 {today}"

body = MIMEText(
    f"안녕하세요,\n\n자동 생성된 Qoo10 랭킹 보고서입니다.\n"
    f"생성일자: {today}\n"
    f"URL: {QOO10_URL}\n\n윤쨩 왕사랑해오! 파이팅구에오!",
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
