# 📈 PX CCF Market Data Auto-Scraper

중국의 화학/섬유 시장 정보 사이트인 **CCFGroup**에서 PX, PTA, MEG 및 폴리에스터 관련 일간/주간 시장 데이터를 자동으로 수집하고, 마진(Margin)을 계산하여 이메일로 리포팅하는 파이썬 자동화 파이프라인입니다.

## 🚀 주요 기능 (Features)
* **병렬 웹 스크래핑**: `ThreadPoolExecutor`와 `lxml` 파서를 활용하여 여러 리포트 URL과 본문 데이터를 멀티스레딩으로 빠르고 안정적으로 수집합니다.
* **데이터 전처리 및 마진 계산**: 수집된 Raw Data를 Pandas를 이용해 정제하고, 사전에 정의된 공식(예: `0.855*PTA + 0.335*MEG` 등)에 따라 각 제품별 마진을 자동 계산합니다.
* **엑셀 리포트 자동 생성**: `xlsxwriter`를 사용하여 시각적으로 깔끔한 포맷(테두리, 헤더 배경색 지정, 셀 너비 자동 맞춤)이 적용된 엑셀 파일(`.xlsx`)을 매일 생성합니다.
* **자동화된 이메일 발송**: 매일 생성된 엑셀 리포트와 요약된 본문 표를 지정된 수신처(`jp_lee@sk.com`)로 자동 발송합니다.
* **CI/CD 스케줄링**: GitHub Actions를 통해 로컬 PC 구동 없이 **매주 평일(월~금) KST 오전 10시**에 클라우드 환경에서 스크립트가 자동으로 실행됩니다.

## 📁 저장소 구조 (Repository Structure)
```text
PX/
├── .github/
│   └── workflows/
│       └── px_daily_report.yml  # GitHub Actions 스케줄러 설정 파일
├── main.py                      # 스크래핑, 데이터 처리 및 이메일 발송 메인 스크립트
└── README.md                    # 프로젝트 가이드
⚙️ 환경 설정 (Setup & Installation)
이 프로젝트를 포크(Fork)하거나 클론하여 새로 구축할 경우, 이메일 발송을 위한 환경 변수 설정이 반드시 필요합니다.

1. GitHub Secrets 등록 (필수)

자동 이메일 발송 기능을 위해 발신용 Gmail 계정 정보가 필요합니다. GitHub 저장소의 [Settings] -> [Secrets and variables] -> [Actions] 로 이동하여 아래 두 가지 변수를 등록해 주세요.

GMAIL_USER: 발송에 사용할 Gmail 주소 (예: example@gmail.com)

GMAIL_APP_PASSWORD: 해당 Gmail 계정의 16자리 앱 비밀번호 (띄어쓰기 없이 입력)

2. 로컬 실행 시 의존성 패키지 설치

로컬 PC에서 스크립트를 수동으로 테스트하거나 수정할 경우 아래 라이브러리들을 설치해야 합니다.

Bash
pip install requests pandas beautifulsoup4 lxml xlsxwriter urllib3
(실행 시 환경 변수가 로컬에 설정되어 있지 않으면 메일 발송은 건너뛰고 엑셀 파일만 로컬에 생성됩니다.)

⏰ 스케줄 (Schedule)
실행 주기: 매주 월요일 ~ 금요일

실행 시간: KST 오전 10:00 (UTC 01:00)

구동 방식: GitHub Actions Ubuntu 서버 워크플로우를 통한 자동 실행 (cron: '0 1 * * 1-5')
