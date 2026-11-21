import os
import sys
import time
import smtplib
import base64
import requests
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
from datetime import datetime, timedelta
import builtins

# ==============================================================
# 📂 경로 동적 설정 (중요: 하드코딩 방지)
# ==============================================================
def get_base_path():
    """실행 파일(.exe) 또는 스크립트(.py)가 위치한 폴더 경로를 반환합니다."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_path()

GCP_KEY_FILENAME = "qoo10-scraper-auto-22518a3e5c96.json"
GCP_KEY_PATH = os.path.join(BASE_DIR, GCP_KEY_FILENAME)

# ==============================================================
# 🕒 Timestamped Print
# ==============================================================
_original_print = builtins.print
def timestamped_print(*args, **kwargs):
    now = (datetime.utcnow() + timedelta(hours=9)).strftime("[%Y-%m-%d %H:%M:%S]")
    _original_print(now, *args, **kwargs)
builtins.print = timestamped_print

# ==============================================================
# 🔐 환경 변수 로드
# ==============================================================
QOO10_URL = os.getenv("QOO10_URL")
HIGHLIGHT_NAME1 = os.getenv("HIGHLIGHT_NAME1")
HIGHLIGHT_NAME2 = os.getenv("HIGHLIGHT_NAME2")
GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_PASS = os.getenv("GMAIL_PASS")
SEND_TO = os.getenv("SEND_TO")

# ==============================================================
# 🖥 Chrome 설정 (리눅스 헤드리스 타임아웃 방지 최적화)
# ==============================================================
chrome_options = Options()
# 🌟 핵심: 광고/트래킹 스크립트 완료를 기다리지 않고 DOM 로드 즉시 진행
chrome_options.page_load_strategy = 'eager'

chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--window-size=1920,1080")
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
chrome_options.add_argument(f"--user-agent={user_agent}")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)
# 페이지 로드 최대 대기 시간을 30초로 제한 (120초 행 걸림 방지)
driver.set_page_load_timeout(30)
driver.set_script_timeout(30)

# ==============================================================
# 📜 스크롤 다운 (중요: 이미지 Lazy Loading 해제)
# ==============================================================
def scroll_to_bottom(driver):
    print("[INFO] 페이지 스크롤 시작 (이미지를 깨우기 위해 천천히 스크롤)...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    current_position = 0
    step = 800
    max_scrolls = 30  # 무한 스크롤 늪 방지용 상한선
    scroll_count = 0

    while current_position < last_height and scroll_count < max_scrolls:
        current_position += step
        driver.execute_script(f"window.scrollTo(0, {current_position});")
        time.sleep(0.4)
        last_height = driver.execute_script("return document.body.scrollHeight")
        scroll_count += 1
        
    driver.execute_script("window.scrollTo(0, 0);")
    print("[INFO] 페이지 스크롤 완료")

# ==============================================================
# 🧩 메가와리 파서 (새로운 'list_v2' 구조 완벽 대응 버전)
# ==============================================================
def parse_megawari(driver):
    results = []
    
    possible_selectors = [
        "ul.list_v2 > li",           # 🌟 최신 Qoo10 개편 사이트 구조
        "ul.megasale_rank_list > li", 
        "ul.type_gallery > li", 
        "ul.list_item > li",
        "div[data-itemcode]",
        "li.list-item"
    ]

    items = []
    for selector in possible_selectors:
        items = driver.find_elements(By.CSS_SELECTOR, selector)
        if len(items) > 0:
            print(f"[INFO] '{selector}' 구조에서 상품 {len(items)}개 감지 성공!")
            break

    if not items:
        print("[ERROR] 등록된 어떠한 형태의 리스트도 감지되지 않았습니다.")
        return results

    for idx, item in enumerate(items[:100], start=1):
        rank = str(idx)
        
        try:
            name_elem = item.find_element(By.CSS_SELECTOR, ".item_info .title, .list_v2_thumbnail img.thumbnail_image")
            if name_elem.tag_name == "img":
                name = name_elem.get_attribute("alt").strip()
            else:
                name = name_elem.text.strip()
        except:
            name = ""
        
        try:
            total = item.find_element(By.CSS_SELECTOR, ".item_info .price, .item_info .sale_price, .prc strong").text.strip()
        except: 
            total = ""
       
        try:
            img_el = item.find_element(By.CSS_SELECTOR, "img.thumbnail_image, img")
            gd_src = img_el.get_attribute("gd_src")
            data_src = img_el.get_attribute("data-src")
            src = img_el.get_attribute("src")
            
            # 회색 기본 플레이스홀더(no_image 등)가 아닌 실제 주소 우선 선택
            img_url = gd_src or data_src or src
            if img_url and ("no_image" in img_url or "blank" in img_url):
                img_url = gd_src or data_src or ""
        except: 
            img_url = ""

        if not name:
            continue

        results.append([rank, name, total, img_url])
        
    return results

# ==============================================================
# 🚀 메인 로직 실행
# ==============================================================
try:
    print(f"[INFO] Qoo10 접속: {QOO10_URL}")
    if not QOO10_URL:
        raise ValueError("QOO10_URL 환경변수가 없습니다.")

    max_retries = 3
    data = []

    for attempt in range(1, max_retries + 1):
        try:
            driver.get(QOO10_URL)
        except TimeoutException:
            print("[WARNING] 페이지 로딩 30초 초과 -> 로딩 강제 중지 후 데이터 추출 진행")
            driver.execute_script("window.stop();")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)  # 동적 리스트 렌더링을 위한 짧은 대기
        scroll_to_bottom(driver)
        
        data = parse_megawari(driver)

        if data:
            print(f"[INFO] {attempt}번째 시도 성공: 상품 {len(data)}개 수집 완료!")
            break
        else:
            print(f"[WARNING] {attempt}번째 시도 실패: 데이터를 찾지 못했습니다.")
            if attempt < max_retries:
                print("[INFO] 5초 대기 후 페이지를 새로고침하여 재시도합니다...")
                time.sleep(5)
            else:
                print("[ERROR] 최대 재시도 횟수를 초과했습니다. 빈 상태로 엑셀을 생성합니다.")

    # 📘 엑셀 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "Qoo10 Ranking"
    ws.append(["순위", "상품명", "총판매액", "이미지"])

    if data:
        for row in data:
            ws.append(row[:-1])

    # 🎨 하이라이트 적용
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    bold_font = Font(bold=True)

    for row in ws.iter_rows(min_row=2, max_col=3):
        product_name = str(row[1].value) if row[1].value else ""
        should_highlight = False
        if HIGHLIGHT_NAME1 and HIGHLIGHT_NAME1 in product_name: should_highlight = True
        if HIGHLIGHT_NAME2 and HIGHLIGHT_NAME2 in product_name: should_highlight = True

        if should_highlight:
            for cell in row:
                cell.fill = yellow_fill
                cell.font = bold_font

    # 🖼 이미지 삽입 (브라우저 세션 연동)
    print("[INFO] 이미지 다운로드 시작...")
    session = requests.Session()
    session.headers.update({"User-Agent": user_agent, "Referer": driver.current_url})
    
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])

    ws.column_dimensions['D'].width = 12

    for i, row in enumerate(data, start=2):
        img_url = row[3]
        if not img_url or img_url.startswith("data:image"): continue
        if img_url.startswith("//"): img_url = "https:" + img_url

        for retry in range(3):  # 최대 3회까지 재시도
            try:
                resp = session.get(img_url, timeout=15)
                if resp.status_code == 200:
                    img_bytes = resp.content
                    # 1차 검증: 파일 크기가 2KB(2048 bytes) 미만이면 깨진 기본 아이콘으로 간주
                    if len(img_bytes) < 2048:
                        raise ValueError(f"비정상적으로 작은 이미지 용량 ({len(img_bytes)} bytes)")
                    image = Image.open(BytesIO(img_bytes))
                    image = image.convert("RGB")
                    # 2차 검증: 이미지 전체의 고유 색상이 15가지 미만이면 단색 회색 아이콘으로 간주
                    colors = image.getcolors(maxcolors=20)
                    if colors is not None and len(colors) < 15:
                        raise ValueError("단색/깨진 아이콘 이미지 감지")
                    image.thumbnail((80, 80))
                    bio = BytesIO()
                    image.save(bio, format="PNG")
                    bio.seek(0)
                    img = XLImage(bio)
                    ws.row_dimensions[i].height = 65 
                    ws.add_image(img, f"D{i}")
                    time.sleep(0.1)
                    break  # 성공 시 재시도 종료
            except Exception as e:
                if retry == 2:
                    print(f"[ERROR] 이미지 삽입 최종 실패 (순위 {row[0]}): {e}")
                else:
                    time.sleep(1)

    # 💾 저장 및 메일 전송
    file_path = os.path.join(BASE_DIR, "Qoo10_Rank.xlsx")
    wb.save(file_path)
    print(f"[INFO] 엑셀 저장 완료: {file_path}")

    if GMAIL_USER and GMAIL_PASS and SEND_TO:
        print(f"[INFO] 메일 전송 시도: {GMAIL_USER} -> {SEND_TO}")
        msg = MIMEMultipart()
        msg["From"] = GMAIL_USER
        msg["To"] = SEND_TO
        today = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d")
        msg["Subject"] = f"Qoo10 랭킹 자동 보고서 {today}"
        
        body = MIMEText(f"스크래퍼에 의해 자동 생성된 Qoo10 보고서입니다.\n파이팅구! 많이 사랑해용♡\n일자: {today}\nURL: {QOO10_URL}", "plain")
        msg.attach(body)

        with open(file_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment; filename=Qoo10_Rank.xlsx")
        msg.attach(part)

        with smtplib.SMTP("smtp.naver.com", 587) as server:
            server.starttls()
            server.login(GMAIL_USER, GMAIL_PASS)
            server.send_message(msg)
        print("[INFO] 메일 전송 완료")
    else:
        print("[WARNING] 메일 환경변수(GMAIL_USER, GMAIL_PASS, SEND_TO) 중 누락된 값이 있어 발송을 건너뜁니다.")
    
except Exception as e:
    print(f"[ERROR] 프로그램 실행 중 오류: {e}")

finally:
    if 'driver' in locals():
        driver.quit()
        print("[INFO] 브라우저 종료")