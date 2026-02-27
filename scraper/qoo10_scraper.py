import os
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
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from datetime import datetime, timedelta
import builtins

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
# 🖥 Chrome 설정
# ==============================================================
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
# Qoo10 차단 방지용 User-Agent 설정
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
chrome_options.add_argument(f"--user-agent={user_agent}")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

# ==============================================================
# 📜 스크롤 다운 (중요: 이미지 Lazy Loading 해제)
# ==============================================================
def scroll_to_bottom(driver):
    print("[INFO] 페이지 스크롤 시작 (이미지를 깨우기 위해 천천히 스크롤)...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    current_position = 0
    step = 800  # 모니터 한 화면 정도의 높이씩 쪼개서 내려감

    while current_position < last_height:
        current_position += step
        driver.execute_script(f"window.scrollTo(0, {current_position});")
        time.sleep(0.4)  # 이미지가 뜰 수 있도록 0.4초 대기
        
        # 스크롤을 내리면서 페이지 길이가 길어졌는지 매번 확인
        last_height = driver.execute_script("return document.body.scrollHeight")
        
    # 다시 맨 위로 (혹시 모를 오류 방지)
    driver.execute_script("window.scrollTo(0, 0);")
    print("[INFO] 페이지 스크롤 완료")

# ==============================================================
# 🧩 메가와리 파서
# ==============================================================
def parse_megawari(driver):
    results = []
    # 선택자 수정 (공통 리스트)
    items = driver.find_elements(By.CSS_SELECTOR, "ul.megasale_rank_list > li, ul.type_gallery > li, ul.list_item > li")
    print(f"[INFO] 감지된 아이템 수: {len(items)}")

    for item in items[:100]:
        try:
            rank = item.find_element(By.CSS_SELECTOR, ".rank_num").text.strip()
        except:
            rank = ""

        try:
            name = item.find_element(By.CSS_SELECTOR, ".title, .sbj").text.strip()
        except:
            name = ""

        try:
            total = item.find_element(By.CSS_SELECTOR, ".price, .prc strong").text.strip()
        except:
            total = ""

try:
            img_el = item.find_element(By.CSS_SELECTOR, "img")
            # Qoo10 특유의 원본 이미지 저장 속성인 gd_src를 1순위로 찾습니다.
            img_url = (
                img_el.get_attribute("gd_src")
                or img_el.get_attribute("data-src")
                or img_el.get_attribute("data-original")
                or img_el.get_attribute("data-lazy")
                or img_el.get_attribute("src")
            )
        except:
            img_url = ""

        results.append([rank, name, total, img_url])

    return results

# ==============================================================
# 🚀 메인 로직 실행
# ==============================================================
try:
    print(f"[INFO] Qoo10 접속: {QOO10_URL}")
    if not QOO10_URL:
        raise ValueError("QOO10_URL 환경변수가 없습니다.")

    driver.get(QOO10_URL)
    
    # 1. 리스트 대기
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    
    # 2. 스크롤 다운 (이미지 URL 활성화를 위해 필수)
    scroll_to_bottom(driver)
    
    # 3. 데이터 파싱
    data = parse_megawari(driver)

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
    # 🎨 하이라이트 적용
    # ==============================================================
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    bold_font = Font(bold=True)

    for row in ws.iter_rows(min_row=2, max_col=3):
        product_name = str(row[1].value) if row[1].value else ""
        should_highlight = False

        if HIGHLIGHT_NAME1 and HIGHLIGHT_NAME1 in product_name:
            should_highlight = True
        if HIGHLIGHT_NAME2 and HIGHLIGHT_NAME2 in product_name:
            should_highlight = True

        if should_highlight:
            for cell in row:
                cell.fill = yellow_fill
                cell.font = bold_font

    # ==============================================================
    # 🖼 이미지 삽입 (브라우저 세션 연동)
    # ==============================================================
    print("[INFO] 이미지 다운로드 시작 (브라우저 세션 연동)...")
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": user_agent,
        "Referer": driver.current_url,
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"
    })
    
    cookies = driver.get_cookies()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    # 엑셀 D열(이미지 열)의 너비를 조금 넓혀줍니다.
    ws.column_dimensions['D'].width = 12

    for i, row in enumerate(data, start=2):
        img_url = row[3]
        if not img_url:
            continue

        if img_url.startswith("//"):
            img_url = "https:" + img_url
            
        # Data URI (Base64) 더미 이미지인 경우 스킵 (보통 레이지 로딩이 덜 풀렸을 때 발생)
        if img_url.startswith("data:image"):
            print(f"[WARN] Base64 더미 이미지 감지되어 스킵 (순위 {row[0]})")
            continue

        try:
            resp = session.get(img_url, timeout=10)
            
            if resp.status_code != 200:
                print(f"[WARN] 이미지 접근 거부 ({resp.status_code}): {img_url}")
                continue

            img_bytes = resp.content
            image = Image.open(BytesIO(img_bytes))

            if image.mode in ("RGBA", "LA"):
                background = Image.new("RGB", image.size, (255, 255, 255))
                background.paste(image, mask=image.split()[-1])
                image = background
            elif image.mode == "P":
                image = image.convert("RGBA")
                background = Image.new("RGB", image.size, (255, 255, 255))
                background.paste(image, mask=image.split()[-1])
                image = background
            else:
                image = image.convert("RGB")

            image.thumbnail((80, 80))
            bio = BytesIO()
            image.save(bio, format="PNG")
            bio.seek(0)
            
            img = XLImage(bio)
            # 엑셀 행 높이를 80 픽셀 크기에 맞게 조절 (약 60 포인트)
            ws.row_dimensions[i].height = 65 
            
            # 셀 안에 이미지 삽입
            ws.add_image(img, f"D{i}")
            
            time.sleep(0.1) # 서버 부하 방지를 위해 0.1초로 조금 늘림

        except Exception as e:
            # 🌟 핵심: 이제 에러를 숨기지 않고 C# 로그 창에 출력되도록 합니다!
            print(f"[ERROR] 이미지 삽입 실패 (순위 {row[0]}): {e} / URL: {img_url}")

    # ==============================================================
    # 💾 저장 및 메일 전송
    # ==============================================================
    file_name = "Qoo10_Rank.xlsx"
    wb.save(file_name)
    print(f"[INFO] 엑셀 저장 완료: {file_name}")

    if GMAIL_USER and GMAIL_PASS and SEND_TO:
        msg = MIMEMultipart()
        msg["From"] = GMAIL_USER
        msg["To"] = SEND_TO
        today = (datetime.utcnow() + timedelta(hours=9)).strftime("%Y-%m-%d")
        msg["Subject"] = f"Qoo10 랭킹 자동 보고서 {today}"
        
        body = MIMEText(
            f"안녕하세요,\n\n자동 생성된 Qoo10 랭킹 보고서입니다.\n"
            f"생성일자: {today}\n"
            f"URL: {QOO10_URL}\n\n짱 마아아아아니 마아아아아아아아니 사랑해오!!!!!",
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
        print("[INFO] 메일 전송 완료")
    
except Exception as e:
    print(f"[ERROR] 프로그램 실행 중 오류: {e}")

finally:
    # 🚀 중요: 모든 작업이 끝난 후에 브라우저 종료
    if 'driver' in locals():
        driver.quit()
        print("[INFO] 브라우저 종료")
