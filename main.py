import os
import smtplib
from email.message import EmailMessage
import requests
import warnings
import urllib3
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime, timedelta
from io import StringIO
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

warnings.filterwarnings('ignore')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# 1. Session / Headers (멀티스레딩 & GitHub Actions 최적화)
# =========================================================
session = requests.Session()
session.verify = False

# GitHub Actions의 일시적 네트워크 오류와 멀티스레드 환경을 위한 어댑터 설정
retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=retries)
session.mount('http://', adapter)
session.mount('https://', adapter)

headers = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Content-Type': 'application/x-www-form-urlencoded',
    'Referer': 'https://www.ccfgroup.com/member/member.php',
}

BASE_URL = "https://www.ccfgroup.com"

# =========================================================
# 2. Login Function
# =========================================================
def login_ccfgroup(session, headers, login_data):
    login_url = "https://www.ccfgroup.com/member/member.php"
    resp = session.post(login_url, data=login_data, headers=headers, timeout=30)
    resp.raise_for_status()
    return session

# =========================================================
# 3. Daily / Weekly Finder (lxml 파서 적용)
# =========================================================
today = datetime.today().date()
offset_days = 1
target_date = today - timedelta(days=offset_days)

def find_market_daily(list_url: str, title_prefix: str):
    resp = session.get(list_url, headers=headers, timeout=30)
    resp.raise_for_status()
    # lxml 파서로 교체하여 파싱 속도 향상
    soup = BeautifulSoup(resp.text, "lxml")
    candidates = []

    for a in soup.find_all("a"):
        text = a.get_text(strip=True)
        if not text.startswith(title_prefix):
            continue
        try:
            date_str = text[text.find("(") + 1 : text.find(")")]
            post_date = datetime.strptime(date_str, "%b %d, %Y").date()
        except Exception:
            continue

        if post_date <= target_date:
            full_url = urljoin(BASE_URL, a.get("href"))
            candidates.append((post_date, full_url))

    if not candidates: return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def find_market_weekly(list_url: str, title_prefix: str):
    resp = session.get(list_url, headers=headers, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    for a in soup.find_all("a"):
        if a.get_text(strip=True).startswith(title_prefix):
            return urljoin(BASE_URL, a.get("href"))
    return None

# =========================================================
# 4. URL Extract (멀티스레딩 병렬 처리)
# =========================================================
# 대상 URL 매핑 딕셔너리 구성
url_tasks = {
    "pta_daily": ("daily", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=100000&Prod_ID=100001", "PTA market daily"),
    "meg_daily": ("daily", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=100000&Prod_ID=100002", "MEG market daily"),
    "yarn_daily": ("daily", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=100000&Prod_ID=100005", "Polyester filament yarn market daily"),
    "fiber_daily": ("daily", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=100000&Prod_ID=100006", "Polyester staple fiber market daily"),
    "bottle_daily": ("daily", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=100000&subclassid=100000&Prod_ID=100004", "PET bottle chip market daily"),
    "px_weekly": ("weekly", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=200000&subclassid=100000&Prod_ID=100001", "PX market weekly"),
    "yarn_weekly": ("weekly", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=200000&subclassid=100000&Prod_ID=100005", "Polyester filament yarn market weekly"),
    "fiber_weekly": ("weekly", "https://www.ccfgroup.com/newscenter/index.php?Class_ID=200000&subclassid=100000&Prod_ID=100006", "Polyester staple fiber market weekly")
}

urls = {}
# 서버 IP 차단 방지를 위해 최대 스레드 4개로 제한
with ThreadPoolExecutor(max_workers=4) as executor:
    futures = {}
    for name, (type_, link, prefix) in url_tasks.items():
        if type_ == "daily":
            futures[executor.submit(find_market_daily, link, prefix)] = name
        else:
            futures[executor.submit(find_market_weekly, link, prefix)] = name
            
    for future in as_completed(futures):
        name = futures[future]
        urls[name] = future.result()

# 기존 로직 하위 호환을 위해 변수 할당
pta_daily = urls.get("pta_daily")
meg_daily = urls.get("meg_daily")
yarn_daily = urls.get("yarn_daily")
fiber_daily = urls.get("fiber_daily")
bottle_daily = urls.get("bottle_daily")
px_weekly = urls.get("px_weekly")
yarn_weekly = urls.get("yarn_weekly")
fiber_weekly = urls.get("fiber_weekly")

print("=== Extracted URLs (Parallel, No Login) ===")
for k, v in urls.items():
    print(f"{k}: {v}")
df_url = pd.Series(urls).to_frame(name='URL')

# =========================================================
# 5. Login
# =========================================================
USERNAME = "SKGlobalKorea"
PASSWORD = "Sk15001657"

login_data = {
    'custlogin': '1',
    'action': 'login',
    'username': USERNAME,
    'password': PASSWORD,
    'savecookie': 'savecookie'
}

session = login_ccfgroup(session, headers, login_data)
print("✅ 로그인 완료 (session 유지됨)")

# =========================================================
# 6. 테이블 및 텍스트 데이터 추출 (lxml 강제 적용)
# =========================================================
def fetch_tables_as_df(session, url, headers):
    if not url: return []
    resp = session.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    html_file_like = StringIO(resp.text)
    # lxml 엔진을 사용하여 DataFrame 변환 속도 극대화
    return pd.read_html(html_file_like, flavor='lxml')

def fetch_average_from_text(session, url, headers, start_marker, end_marker):
    if not url: return None
    resp = session.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'lxml')
    content = soup.find('div', id='fontzoom')
    if not content: return None
    text = content.get_text(separator='\n', strip=True)
    try:
        target_part = text.split(start_marker)[1].rsplit(end_marker, 1)[0]
        nums = [int(n) for n in re.findall(r'\d+', target_part)]
        return sum(nums) / len(nums) if nums else None
    except (IndexError, ValueError):
        return None

# =========================================================
# 8. 데이터 추출 (멀티스레딩 병렬 처리)
# =========================================================
# 추출할 데이터 작업 리스트
fetch_tasks = {
    "pta_daily": pta_daily, "meg_daily": meg_daily, "yarn_daily": yarn_daily,
    "fiber_daily": fiber_daily, "bottle_daily": bottle_daily,
    "px_weekly": px_weekly, "yarn_weekly": yarn_weekly, "fiber_weekly": fiber_weekly
}

extracted_dfs = {}
yarn_avg = None

with ThreadPoolExecutor(max_workers=4) as executor:
    # 1. 테이블 추출 작업 할당
    future_to_name = {executor.submit(fetch_tables_as_df, session, url, headers): name 
                      for name, url in fetch_tasks.items()}
    # 2. 텍스트 추출 작업 할당
    future_yarn_avg = executor.submit(fetch_average_from_text, session, yarn_daily, headers, 'assessed to ', ' near')
    
    # 완료된 순서대로 결과 회수
    for future in as_completed(future_to_name):
        name = future_to_name[future]
        extracted_dfs[name] = future.result()
        
    yarn_avg = future_yarn_avg.result()

# 기존 로직 하위 호환을 위해 변수 재할당
df_pta_daily = extracted_dfs.get("pta_daily", [])
df_meg_daily = extracted_dfs.get("meg_daily", [])
df_yarn_daily = extracted_dfs.get("yarn_daily", [])
df_fiber_daily = extracted_dfs.get("fiber_daily", [])
df_bottle_daily = extracted_dfs.get("bottle_daily", [])
df_px_weekly = extracted_dfs.get("px_weekly", [])
df_yarn_weekly = extracted_dfs.get("yarn_weekly", [])
df_fiber_weekly = extracted_dfs.get("fiber_weekly", [])

print("✅ 본문 테이블 및 텍스트 데이터 병렬 수집 완료")

# =========================================================
# 9. 데이터프레임 처리 (기존 코드 완벽하게 동일)
# =========================================================
df_pta_daily_f = df_pta_daily[2].iloc[:4,:-1].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_meg_daily_f = df_meg_daily[2].iloc[:-1,:].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_yarn_daily_f = df_yarn_daily[2].iloc[:-1,:].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_fiber_daily_f = df_fiber_daily[2].iloc[:-1,:].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))
df_bottle_daily_f = df_bottle_daily[2].iloc[:-1,:].reset_index(drop=True).pipe(lambda d: d.rename(columns=d.iloc[0]).drop(d.index[0]).reset_index(drop=True))

df_px_weekly_f = df_px_weekly[2].iloc[(df_px_weekly[2][0].str.contains('operating', case=False, na=False).idxmax())+1:df_px_weekly[2][0].str.contains('imports', case=False, na=False).idxmax()].T.drop_duplicates(keep='first').T.reset_index(drop=True).pipe(lambda d: d.set_axis(d.iloc[0], axis=1).iloc[1:]).reset_index(drop=True).pipe(lambda df: df.rename(columns={df.columns[0]: 'Operating rate'}))
df_yarn_weekly_f = df_yarn_weekly[8].T.drop_duplicates(keep='first').T.pipe(lambda df: df.set_axis(df.iloc[0], axis=1)).iloc[1:].reset_index(drop=True)
df_fiber_weekly_f = df_fiber_weekly[7].iloc[1:, :].T.drop_duplicates(keep='first').T.pipe(lambda df: df.set_axis(df.iloc[0], axis=1)).iloc[1:].reset_index(drop=True).pipe(lambda df: df.rename(columns={df.columns[0]: 'Index'}))

# 1. Margin Backdata
data_map_margin_backdata_1 = {
    "RMB | by cash, ex-CMP (yuan/mt)": df_pta_daily_f.iloc[0, 2],
    "Spot, East China": df_meg_daily_f.iloc[1, 1],
    "SD POY white | 150D/48F": df_yarn_daily_f.iloc[3, 4],
    "SD FDY white | 150D/96F": df_yarn_daily_f.iloc[11, 4],
    "SD DTY large companies white | 150D/48F,   non-intermingled": df_yarn_daily_f.iloc[19, 4],
    "Virgin   PSF 1.4D*38mm | Daily average(yuan/mt by cash ex-works)": df_fiber_daily_f.iloc[0, 1],
    "PET water/hot fill bottle chip, by cash EXW": df_bottle_daily_f.iloc[0, 2],
}
df_px_margin_backdata = pd.Series(data_map_margin_backdata_1).to_frame(name='Value')
data_map_margin_backdata_2 = {
    "RMB | by cash, ex-CMP (yuan/mt)": df_pta_daily[3].iloc[1, 0].split('(')[1].split(')')[0],
    "Spot, East China": df_meg_daily[3].iloc[1, 0].split('(')[1].split(')')[0],
    "SD POY white | 150D/48F": df_yarn_daily_f.columns[4],
    "SD FDY white | 150D/96F": df_yarn_daily_f.columns[4],
    "SD DTY large companies white | 150D/48F,   non-intermingled": df_yarn_daily_f.columns[4],
    "Virgin   PSF 1.4D*38mm | Daily average(yuan/mt by cash ex-works)": df_fiber_daily[3].iloc[1, 0].split('(')[1].split(')')[0],
    "PET water/hot fill bottle chip, by cash EXW": df_bottle_daily[3].iloc[1, 0].split('(')[1].split(')')[0],
}
df_px_margin_backdata['Date'] = df_px_margin_backdata.index.map(data_map_margin_backdata_2)
df_px_margin_backdata['Value'] = df_px_margin_backdata['Value'].astype(float)

# 2. Margin
px_margin_calc = (0.855*df_px_margin_backdata.iloc[0, 0]) + (0.335*df_px_margin_backdata.iloc[1, 0])
data_map_margin_1 = {
    "POY margin": df_px_margin_backdata.iloc[2, 0] - (px_margin_calc + 1150),
    "FDY margin": df_px_margin_backdata.iloc[3, 0] - (px_margin_calc + 1550),
    "DTY margin": df_px_margin_backdata.iloc[4, 0] - (px_margin_calc + 2450),
    "PSF margin": df_px_margin_backdata.iloc[5, 0] - (px_margin_calc + 900),
    "PET Bottle Chip margin": df_px_margin_backdata.iloc[6, 0] - (px_margin_calc + 700),
    "Polyester 복합 margin": 0
}
df_px_margin = pd.Series(data_map_margin_1).to_frame(name='Value')
df_px_margin.iloc[5,0] = (0.16 * df_px_margin.iloc[0, 0]) + (0.28 * df_px_margin.iloc[1, 0]) + (0.16 * df_px_margin.iloc[2, 0]) + (0.22 * df_px_margin.iloc[3, 0]) + (0.17 * df_px_margin.iloc[4, 0])

# 3. Final Result
data_map_px_1 = {
    "중국 PX": df_px_weekly_f.iloc[0, 1],
    "중국 PTA": df_px_weekly_f.iloc[2, 1],
    "중국 Polyester": df_px_weekly_f.iloc[3, 1],
    "Polyester Filament Yarn": df_yarn_weekly_f.iloc[2, 3],
    "Polyester Staple Fiber": df_fiber_weekly_f.iloc[1, 4],
    "강소/절강 Texturing machine": df_yarn_weekly_f.iloc[8, 3],
    "강소/절강 DTY Machine": df_yarn_weekly_f.iloc[7, 3],
    "중국 Polyester Fiber Yarn": yarn_avg,
    "POY": df_yarn_weekly_f.iloc[4, 3],
    "FDY": df_yarn_weekly_f.iloc[5, 3],
    "DTY": df_yarn_weekly_f.iloc[6, 3],
    "PSF": df_fiber_weekly_f.iloc[0, 4],
    "Polyester 복합": df_px_margin.iloc[5,0],
    "장섬유": df_px_margin.iloc[0,0],
    "단섬유": df_px_margin.iloc[3,0],
    "Bottle Chip": df_px_margin.iloc[4,0]
}
df_px_result = pd.Series(data_map_px_1).to_frame(name='Value')
data_map_px_2 = {
    "중국 PX": df_px_weekly_f.columns[1],
    "중국 PTA": df_px_weekly_f.columns[1],
    "중국 Polyester": df_px_weekly_f.columns[1],
    "Polyester Filament Yarn": df_yarn_weekly_f.columns[3],
    "Polyester Staple Fiber": df_fiber_weekly_f.columns[4],
    "강소/절강 Texturing machine": df_yarn_weekly_f.columns[3],
    "강소/절강 DTY Machine": df_yarn_weekly_f.columns[3],
    "중국 Polyester Fiber Yarn": df_yarn_daily_f.columns[4],
    "POY": df_yarn_weekly_f.columns[3],
    "FDY": df_yarn_weekly_f.columns[3],
    "DTY": df_yarn_weekly_f.columns[3],
    "PSF": df_fiber_weekly_f.columns[4],
    "Polyester 복합": "",
    "장섬유": "",
    "단섬유": "",
    "Bottle Chip": ""
}
df_px_result['Date'] = df_px_result.index.map(data_map_px_2)
df_px_result['Value'] = df_px_result['Value'].astype(float).round(1)

today = datetime.now().strftime('%Y-%m-%d')
file_name = f"px_result_{today}.xlsx"

targets = {"px_result": df_px_result, "px_margin": df_px_margin, 
           "px_margin_backdata": df_px_margin_backdata, "url": df_url}

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    workbook = writer.book
    
    border_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#BCD1E4', 'border': 1, 'align': 'center'})

    for sheet_name, df in targets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=True)
        worksheet = writer.sheets[sheet_name]
        
        rows, cols = df.shape
        worksheet.conditional_format(0, 0, rows, cols, {
            'type': 'formula',
            'criteria': '=1',  # 무조건 참이 되도록 설정하여 빈 칸 포함 전체 영역에 테두리 적용
            'format': border_format
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        worksheet.write(0, 0, df.index.name or 'Index', header_format)

        idx_max_len = max(df.index.astype(str).map(len).max(), len(str(df.index.name or "Index"))) + 3
        worksheet.set_column(0, 0, idx_max_len)
        
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 3
            worksheet.set_column(i + 1, i + 1, column_len)

print(f"✅ '{file_name}' 저장 완료!")

# =========================================================
# 10. 이메일 발송 (Gmail 앱 비밀번호 사용)
# =========================================================
print("=== 메일 발송 준비 ===")

# GitHub Secrets에서 환경변수로 주입받을 이메일 계정 정보
sender_email = os.environ.get("GMAIL_USER")
app_password = os.environ.get("GMAIL_APP_PASSWORD")
receiver_email = "jp_lee@sk.com"

# 메일 제목 (yyyy-mm-dd 형식)
subject = f"PX CCF {today}"

# 본문 데이터프레임을 HTML 표로 변환 (엑셀 서식인 테두리와 중앙 정렬, 음영을 CSS로 구현)
html_table = df_px_result.to_html(border=0, justify='center', index=True)
html_body = f"""
<html>
<head>
<style>
    table {{
        border-collapse: collapse;
        text-align: center;
        font-family: Arial, sans-serif;
    }}
    th {{
        background-color: #BCD1E4; /* 엑셀 헤더와 동일한 배경색 */
        font-weight: bold;
        border: 1px solid black;
        padding: 5px 10px;
    }}
    td {{
        border: 1px solid black;
        padding: 5px 10px;
    }}
</style>
</head>
<body>
    <p>안녕하세요,</p>
    <p>오늘자 CCF 추출 결과입니다. 상세 내용은 첨부파일을 확인해 주시기 바랍니다.</p>
    <br>
    {html_table}
</body>
</html>
"""

# 메일 객체 생성 및 설정
msg = EmailMessage()
msg['Subject'] = subject
msg['From'] = sender_email
msg['To'] = receiver_email
msg.set_content("HTML 뷰어를 지원하는 메일 클라이언트를 사용해 주세요.") # Fallback 텍스트
msg.add_alternative(html_body, subtype='html')

# 작성된 엑셀 파일 첨부
with open(file_name, 'rb') as f:
    excel_data = f.read()
    
msg.add_attachment(
    excel_data, 
    maintype='application', 
    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
    filename=file_name
)

# SMTP 서버 연결 및 전송
if sender_email and app_password:
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)
        print("✅ 이메일 발송 완료!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")
else:
    print("⚠️ GMAIL_USER 또는 GMAIL_APP_PASSWORD 환경변수가 설정되지 않아 메일을 발송하지 않았습니다.")
