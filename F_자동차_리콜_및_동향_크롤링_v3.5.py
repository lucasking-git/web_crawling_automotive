import os                                                           # 운영체제와 상호작용하기 위한 모듈
import re                                                           # 정규 표현식을 사용하기 위한 모듈
import ssl                                                          # SSL 관련 작업을 위한 모듈
import sys                                                          # 파이썬 인터프리터와 상호작용하기 위한 모듈
import msvcrt                                                       # 윈도우에서 키보드 입력 등의 기능을 제공하는 모듈
import pyperclip                                                    # 클립보드에 접근하기 위한 모듈
import pyautogui                                                    # GUI 자동화(키보드, 마우스 조작)를 위한 모듈
import time                                                         # 시간 관련 모듈
from getpass import getpass                                         # 비밀번호와 같은 민감한 정보를 안전하게 입력받기 위한 함수
import xlwings as xw                                                # 엑셀과 상호작용하기 위한 라이브러리
import pandas as pd                                                 # 데이터 처리와 분석을 위한 라이브러리
import datetime                                                     # 날짜와 시간 관련 작업을 위한 모듈
import colorama                                                     # 텍스트 색상 조작을 위한 모듈
from colorama import Fore, Back, Style                              # 텍스트 글자/배경 변경
from selenium import webdriver                                      # 웹 브라우저를 자동으로 조작하기 위한 모듈
from selenium.webdriver.common.keys import Keys                     # 키보드 키 상수(예: Keys.ENTER)
from selenium.webdriver.common.by import By                         # 웹 페이지의 요소를 찾을 때 사용하는 방법을 정의
from selenium.webdriver.chrome.service import Service               # ChromeDriver를 시작하는 데 사용하는 서비스
from webdriver_manager.chrome import ChromeDriverManager            # ChromeDriver의 자동 설치 및 관리를 위한 클래스
from selenium.webdriver.support.ui import Select                    # 웹 페이지의 <select> 요소와 상호작용하기 위한 클래스
from selenium.webdriver.support.ui import WebDriverWait             # 웹 페이지의 요소가 로드될 때까지 대기하는 기능을 제공
from selenium.webdriver.support import expected_conditions as EC    # 특정 요소가 특정 상태에 도달할 때까지 대기하는 조건을 정의
from selenium.webdriver.chrome.options import Options               # Chrome 브라우저를 시작할 때 사용할 수 있는 추가 옵션을 설정
from selenium.webdriver.common.alert import Alert                   # 웹 페이지의 경고창과 상호작용하기 위한 클래스
from selenium.webdriver import ActionChains                         # 복잡한 사용자 액션(예: 드래그 앤 드롭)을 구현하기 위한 클래스


# colorama 초기화 (Windows에서 colorama의 ANSI 코드 해석을 활성화하는 데 필수적)
colorama.init(autoreset=True, wrap=True)

# 라인추가 함수
def add_line():
    print("----------------------------------------------------------------------------------------------------------------")

# 프로그램 정보 관련 변수 설정
program_version = "v3.5"
program_release_date = "2024-04-01"
program_creator = "CQO / Quality Management / Quality Planning / sunsik.wang"
# 3.1   20240318    텍스트 색상변경
# 3.2   20240329    colorma 초기화 코드 추가
# 3.3   20240320    메일본문 입력 오타코드 수정 (무상수리 부분)
# 3.4   20240320    program_info 내용 일부 수정 (주의사항)
# 3.5   20240401    crawl_nhtsa() 코드변경 (Last 7 Days 클릭 부분)



def print_program_info(version, release_date, creator):
    info = f"""
----------------------------------------------------------------------------------------------------------------
Creator : {creator}
Program Version : {version}
Program Release Date : {release_date}
Description : 자동차 리콜 및 동향 등의 정보를 수집하여 자동으로 이메일로 송부하는 프로그램
----------------------------------------------------------------------------------------------------------------
Contact and Support: For assistance or inquiries, please contact sunsik.wang@hlcompany.com
----------------------------------------------------------------------------------------------------------------
{Fore.RED + Style.BRIGHT}주의사항{Fore.RESET + Back.RESET + Style.RESET_ALL}
  본 프로그램은 크롬(Chrome) 기반으로 실행 되어서 {Fore.YELLOW}크롬이 기본적으로 설치{Fore.RESET} 되어 있어야 합니다.
  
  가급적 프로그램 실행 중 다른 작업을 하지 말아주세요.
  
  {Fore.YELLOW}사내망이 허용한 IP 주소로 접속{Fore.RESET}이 되어 있어야 한마루 접속 시 문제없이 코드가 작동합니다.
    - 사내가 아닌 자택 혹은 다른장소의 IP에 접속이 되어 있으면 한마루 접속시 문제가 발생합니다.
----------------------------------------------------------------------------------------------------------------"""
    print(info)


# 변수설정
now = datetime.datetime.now().strftime("%Y%m%d")
file_path = f"{now}_자동차 리콜 및 산업동향 정보 조회결과.xlsx"
current_directory = os.getcwd()
save_file_path = os.path.join(os.getcwd(), file_path)


def get_secure_password(prompt=f">> {Fore.YELLOW}패스워드를 입력{Fore.RESET}해주세요 (대소문자 구분) : "):
    password = ""
    print(prompt, end="", flush=True)

    while True:
        key = msvcrt.getch()
        key = key.decode("utf-8")

        if key == "\r" or key == "\n":
            break
        elif key == "\x08":  # Backspace key
            if password:
                print("\b \b", end="", flush=True)
                password = password[:-1]
        else:
            print("*", end="", flush=True)
            password += key

    print()  # Move to the next line after password entry
    return password

# # 환경 변수에서 아이디, 비밀번호를 가져오기
# Hanmaru_ID = os.environ.get('HANMARU_ID')
# Hanmaru_PW = os.environ.get('HANMARU_PW')

# URL 변수 설정
url_naver = "https://search.naver.com/search.naver?sm=tab_hty.top&where=news&query=%EC%B0%A8%EB%9F%89+%EB%A6%AC%EC%BD%9C&oquery=%EC%B0%A8%EB%9F%89+%EB%A6%AC%EC%BD%9C&tqi=iL7LCsp0JXVss5Q0N1Rssssstx4-283104"
url_recallcenter = "https://www.car.go.kr/ri/stat/list.do"
url_recallcenter_무상수리 = "https://www.car.go.kr/ri/grts/list.do"
url_nthsa = "https://www.nhtsa.gov/search-safety-issues#recall"
url_autowein = "https://autowein.com/"
url_autoway = "https://partners.hmc.co.kr/"
url_hanmaru = "https://ep.hlcompany.com/"

# 드라이버 셋업
def setup_driver():
    options = Options()
    options.add_argument("disable-infobars")
    options.add_argument("start-maximized")
    options.add_experimental_option("detach", True)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# 엘리먼트 대기
def wait_for_element(driver, selector_type, selector, wait_time=5):
    selector_types = {
        'css': By.CSS_SELECTOR,
        'class': By.CLASS_NAME,
        'id': By.ID,
        'tag': By.TAG_NAME,
        'xpath': By.XPATH
    }
    return WebDriverWait(driver, wait_time).until(
        EC.presence_of_element_located((selector_types[selector_type], selector))
    )

# 엘리먼트 클릭
def click_element(driver, selector_type, selector, wait_time=5):
    element = wait_for_element(driver, selector_type, selector, wait_time)
    element.click()

# 텍스트 작성
def input_text(driver, selector_type, selector, text, wait_time=5):
    element = wait_for_element(driver, selector_type, selector, wait_time)
    element.clear()
    element.send_keys(text)


# 메인 창 제외 모든 팝업 창 닫기
def close_popup(driver):
    try:
        time.sleep(3)
        
        for i in range(len(driver.window_handles) - 1):
            driver.switch_to.window(driver.window_handles[1])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
    except:
        pass

# 팝업 대기 및 처리
def close_alert_popup(driver):
    try:
        # WebDriverWait를 사용하여 팝업이 나타날 때까지 3~5초간 대기
        popup = WebDriverWait(driver, 5).until(EC.alert_is_present())

        # 팝업 확인 버튼 클릭
        popup.accept()
        
    except Exception as e:
        # 기타 예외 발생 시
        print("Error while waiting for popup:", e)


# 첨부파일 추가여부 확인 (한마루용)
def wait_for_one_ul_tag(driver):
    element = driver.find_element(By.CSS_SELECTOR, "#RAON_K_UP_file_temp")
    ul_tags = element.find_elements(By.TAG_NAME, "ul")
    return len(ul_tags) == 1

def get_user_input_with_recipients():
    print(f"크롤링 결과를 최종적으로 송부하기 위해 {Fore.YELLOW}본인의 한마루 아이디/패스워드를 입력{Fore.RESET}해야 합니다.")
    add_line()
    
    # 입력받기
    Hanmaru_ID = input(f">> {Fore.YELLOW}본인의 한마루 아이디{Fore.RESET}를 입력해주세요 {Fore.RED}(@hlcompany.com 제외){Fore.RESET} : ")
    Hanmaru_PW = get_secure_password()
    # Hanmaru_PW = getpass("패스워드 입력 : ")
    add_line()
    

    # 이메일 수신자 설정
    recipients_list = []
    확인 = ""
    
    for _ in enumerate(range(1, 100)):  # 100번까지 루프 수행 (무한 루프 방지용)
        continue_input = input(f">> 본인 외에 {Fore.YELLOW}크롤링 결과 수신자{Fore.RESET}를 더 {Fore.YELLOW}추가{Fore.RESET}하시겠습니까? (Y/N) : ")
        add_line()
        
        if continue_input.upper() != 'Y':
            recipients_list.append(Hanmaru_ID)  # Hanmaru_ID를 맨 뒤에 추가
            print("현재 받는사람 리스트 :", recipients_list)
            확인 = input(f"위에 있는 메일 수신자가 {Fore.YELLOW}오타없이 올바르게 입력{Fore.RESET}되었습니까? (Y/N) : ")
            add_line()

            if 확인.upper() != 'Y':
                recipients_list.clear()
                continue
            else:
                break

        email_input = input(f"{Fore.YELLOW}추가 수신자의 이메일주소{Fore.RESET}를 입력해주세요 {Fore.RED}(@hlcompany.com 제외){Fore.RESET} : ")
        recipients_list.append(email_input)
        add_line()

    return Hanmaru_ID, Hanmaru_PW, recipients_list



# 크롤링_뉴스 기사
def crawl_naver(driver, url):
    driver.get(url)
    
    elems = driver.find_element(By.CSS_SELECTOR, "#main_pack > section > div > div.group_news > ul").find_elements(By.CLASS_NAME, "bx")
    data_naver = []
    for idx, elem in enumerate(elems):
        언론사 = elem.find_element(By.CLASS_NAME, "info_group").find_element(By.TAG_NAME, "a").text.replace("언론사 선정", "")
        제목 = elem.find_element(By.CLASS_NAME, "news_tit").text
        if len(elem.find_element(By.CLASS_NAME, "info_group").find_elements(By.TAG_NAME, "span")) == 2:
            기사_발행시간 = elem.find_element(By.CLASS_NAME, "info_group").find_elements(By.TAG_NAME, "span")[1].text
        else:
            기사_발행시간 = elem.find_element(By.CLASS_NAME, "info_group").find_elements(By.TAG_NAME, "span")[2].text
        
        주요내용 = elem.find_element(By.CLASS_NAME, "dsc_wrap").text
        접속URL = elem.find_element(By.CLASS_NAME,"news_tit").get_attribute("href")
        data_naver.append([idx + 1, 언론사, 제목, 기사_발행시간, 주요내용, 접속URL])

    cols_naver = "No, 언론사, 제목, 기사_발행시간, 주요내용, 접속URL".replace(" ", "").split(",")
    
    return pd.DataFrame(data_naver, columns=cols_naver)
    

# 크롤링_국내 리콜
def crawl_recallcenter(driver, url):
    driver.get(url)

    data_recallcenter = []
    for i in range(5):
        driver.find_element(By.CSS_SELECTOR, f"#content > div > ul.board-hrznt-list > li:nth-child({i+1}) > a").click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#content > div > div.stat-tit > div.info > dl > dd:nth-child(4)")))

        작성일 = driver.find_element(By.CSS_SELECTOR, "#content > div > div.stat-tit > div.info > dl > dd:nth-child(4)").text
        제작사 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(1) > td:nth-child(2)").text
        차종 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(1) > td:nth-child(4)").text
        내용 = driver.find_element(By.CSS_SELECTOR, "#content > div > div.stat-tit > div.subject").text.split("]")[1].strip().split("-")[1].strip()
        생산기간 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(2) > td:nth-child(2)").text
        대상수량 = int(driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(3) > td:nth-child(2)").text.split("대")[0].strip().replace(",",""))
        결함내용 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(4) > td").text
        시정방법 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(5) > td").text

        data_recallcenter.append([i + 1, 작성일, 제작사, 차종, 내용, 생산기간, 대상수량, 결함내용, 시정방법])
        driver.back()

    cols_recallcenter = "No, 작성일, 제작사, 차종, 내용, 생산기간, 대상수량, 결함내용, 시정방법".replace(" ", "").split(",")
    
    return pd.DataFrame(data_recallcenter, columns=cols_recallcenter)

# 크롤링_국내 리콜(무상수리)
def crawl_recallcenter_무상수리(driver, url):
    driver.get(url)
    data_recallcenter_무상수리 = []
    for i in range(5):
        driver.find_element(By.CSS_SELECTOR, f"#content > div > ul.board-hrznt-list > li:nth-child({i+1}) > a").click()
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#content > div > div.stat-tit > div.info > dl > dd:nth-child(4)")))

        작성일 = driver.find_element(By.CSS_SELECTOR, "#content > div > div.stat-tit > div.info > dl > dd:nth-child(4)").text
        제작사 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(1) > td:nth-child(2)").text
        차종 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(1) > td:nth-child(4)").text
        내용 = driver.find_element(By.CSS_SELECTOR, "#content > div > div.stat-tit > div.subject").text.split("]")[1].strip().split("-")[1].strip()
        생산기간 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(2) > td:nth-child(2)").text
        대상수량 = int(driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(3) > td:nth-child(2)").text.split("대")[0].strip().replace(",",""))
        결함내용 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(4) > td").text
        시정방법 = driver.find_element(By.CSS_SELECTOR, "#content > div > table > tbody > tr:nth-child(5) > td").text

        data_recallcenter_무상수리.append([i + 1, 작성일, 제작사, 차종, 내용, 생산기간, 대상수량, 결함내용, 시정방법])
        driver.back()

    cols_recallcenter_무상수리 = "No, 작성일, 제작사, 차종, 내용, 생산기간, 대상수량, 결함내용, 시정방법".replace(" ", "").split(",")

    return pd.DataFrame(data_recallcenter_무상수리, columns=cols_recallcenter_무상수리)


# 크롤링_북미 리콜
def crawl_nhtsa(driver, url):
    driver.get(url)

    # Data Range 날짜구간 클릭
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CLASS_NAME, "input-group-addon")))
    driver.find_element(By.CLASS_NAME, "input-group-addon").click()


    # 'ranges' 클래스를 포함하는 요소 내에서 'Last 7 Days' 텍스트가 포함된 li 요소를 찾아 클릭하기 위한 WebDriverWait
    last_7_days_element = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='ranges']/ul/li[contains(text(), 'Last 7 Days')]"))
    )
    last_7_days_element.click()
    
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CLASS_NAME, "panel-item")))
    elems = driver.find_elements(By.CLASS_NAME, "panel-item")

    data_nhtsa = []

    def extract_data(css_selector, is_text=True, split_char=None, index=None, replace_old=None, replace_new=None, is_int=False):
        try:
            data = elem.find_element(By.CSS_SELECTOR, css_selector)
            if is_text:
                data = data.text
                if split_char and index:
                    data = data.split(split_char)[index]
                if replace_old and replace_new:
                    data = data.replace(replace_old, replace_new)
                if is_int:
                    data = int(data)
            else:
                data = data.get_attribute("aria-expanded")
            return data
        except:
            return ""

    for idx, elem in enumerate(elems):
        # + 버튼이 아니면 + 버튼 클릭
        aria_expanded = elem.find_element(By.TAG_NAME, "a").get_attribute("aria-expanded")
        if aria_expanded == "false":
            elem.find_element(By.CLASS_NAME, "faq-Icon").click()

        이슈날짜 = elem.text.split("NHTSA")[0].rstrip()
        캠페인번호 = elem.text.split("NHTSA")[1].split(":")[1].split("\n")[0].lstrip()
        제목 = elem.find_element(By.CLASS_NAME, "panel-title-caption").text.split('\n')[0]
        제조업체 = extract_data(".panel-body p:nth-of-type(2) span")
        제품 = extract_data(".panel-body p:nth-of-type(3) span")
        주요내용 = extract_data(".panel-title-caption", split_char="\n", index=1)
        대상수량 = elem.find_element(By.CLASS_NAME, "panel-body").find_elements(By.TAG_NAME, "p")[3].find_element(By.TAG_NAME, "span").text.replace(",", "")
        내용요약 = extract_data(".panel-body div:nth-of-type(1) p:nth-of-type(2)")
        조치 = extract_data(".panel-body div:nth-of-type(2) p:nth-of-type(2)")

        data_nhtsa.append([idx + 1, 이슈날짜, 캠페인번호, 제목, 제조업체, 제품, 주요내용, 대상수량, 내용요약, 조치])

    global cnt_nhtsa
    cnt_nhtsa = idx + 1
    
    cols_nhtsa = "No, 이슈날짜, 캠페인번호, 제목, 제조업체, 제품, 주요내용, 대상수량, 내용요약, 조치".replace(" ", "").split(",")
    
    return pd.DataFrame(data_nhtsa, columns=cols_nhtsa)


# 크롤링_품질 동향
def crawl_autowein(driver, url):
    driver.get(url)
    close_alert_popup(driver)

    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#grid-wrap > div:nth-child(1) > div.bg-img.pointer")))

    data_autowein = []
    elems = driver.find_element(By.CSS_SELECTOR, "#grid-wrap").find_elements(By.CLASS_NAME, "py-3")

    elems[0].find_element(By.CSS_SELECTOR, "#grid-wrap > div:nth-child(1) > div.news-wrap.ml-4 > span").text # 발행날짜
    elems[0].find_element(By.CSS_SELECTOR, "#grid-wrap > div:nth-child(1) > div.news-wrap.ml-4 > a.news-title.mb-1.pointer").text # 제목
    elems[0].find_element(By.CSS_SELECTOR, "#grid-wrap > div:nth-child(1) > div.news-wrap.ml-4 > a.news-text.pointer").text.split(".")[0] + "." # 주요내용

    for idx, elem in enumerate(elems):
        발행날짜 = elem.find_element(By.CSS_SELECTOR, f"#grid-wrap > div:nth-child({idx+1}) > div.news-wrap.ml-4 > span").text # 발행날짜
        제목 = elem.find_element(By.CSS_SELECTOR, f"#grid-wrap > div:nth-child({idx+1}) > div.news-wrap.ml-4 > a.news-title.mb-1.pointer").text # 제목
        주요내용 = elem.find_element(By.CSS_SELECTOR, f"#grid-wrap > div:nth-child({idx+1}) > div.news-wrap.ml-4 > a.news-text.pointer").text.split("다.")[0] + "다." # 주요내용
        
        data_autowein.append([idx+1, 발행날짜, 제목, 주요내용])

    cols_autowein = "No, 발행날짜, 제목, 주요내용".replace(" ", "").split(",")
    
    return pd.DataFrame(data_autowein, columns=cols_autowein)

# 크롤링_HKMC 대외문
def crawl_autoway(driver, url, username, password):
    driver.get(url)

    driver.find_element(By.CSS_SELECTOR, "#userID").send_keys(username) # ID 입력
    driver.find_element(By.CSS_SELECTOR, "#password").send_keys(password) # PW 입력
    driver.find_element(By.CSS_SELECTOR, "#Login").click() # 로그인 버튼 클릭

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#divGNBHeaderSub > div.l-functions__list > div.l-functions__item.l-functions__item--sign.c-tooltip > button")))
    driver.find_element(By.CSS_SELECTOR, "#divGNBHeaderSub > div.l-functions__list > div.l-functions__item.l-functions__item--sign.c-tooltip > button").click() # 대외문 버튼 클릭

    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#ifrSubSys")) # iframe 요소 전환

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#documentSubject > a")))
    elems = driver.find_element(By.CLASS_NAME, "c-table__tbody").find_elements(By.CLASS_NAME, "c-table__tr") # 대외문 문서 elem

    data_hkmc = []
    for idx, elem in enumerate(elems):
        수신일 = elem.find_elements(By.TAG_NAME, "td")[8].text
        기안부서 = elem.find_elements(By.TAG_NAME, "td")[6].text
        기안자 = elem.find_elements(By.TAG_NAME, "td")[7].text
        제목 = elem.find_elements(By.TAG_NAME, "td")[5].text
        
        data_hkmc.append([idx+1, 수신일, 기안부서, 기안자, 제목])

    # 기본 컨텐츠로 전환
    driver.switch_to.default_content()
    
    cols_hkmc = "No, 대외문_수신일, 기안부서, 기안자, 제목".replace(" ","").split(",")
    return pd.DataFrame(data_hkmc, columns=cols_hkmc)


# 엑셀 저장
def save_excel(dataframes, file_path):
    with pd.ExcelWriter(path=file_path) as writer:
        for name, df in dataframes.items():
            df.to_excel(writer, sheet_name=name, index=False)


# 한마루 접속 및 로그인
def access_hanmaru(driver, url, username, password):
    driver.get(url)
    # 한마루 접속 후 발생할 수 있는 경고창 처리
    close_alert_popup(driver)
    
    # 로그인
    input_text(driver, "css", "#lvLogin_LoginID", username)
    input_text(driver,"css", "#lvLogin_Password", password +  Keys.ENTER)

    # 팝업창 닫기
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#TopNav > div > ul > li:nth-child(2) > a"))) # 웹페이지 로딩 기다리기
    close_popup(driver)


# 한마루에서 메일 작성
def write_email_hanmaru(driver, df_naver, df_recallcenter, df_recallcenter_무상수리, df_nhtsa, df_autowein, df_autoway, lst_toInput):
    # 메일 클릭
    driver.find_element(By.CSS_SELECTOR, "#TopNav > div > ul > li:nth-child(2) > a").click()

    # 1. 메일쓰기 클릭 (이후 창 전환)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnItemSend > span"))) # 웹페이지 로딩 기다리기
    driver.find_element(By.CSS_SELECTOR, "#btnItemSend > span").click() # 메일쓰기 클릭
    driver.switch_to.window(driver.window_handles[1]) # 클릭 이후 창 전환 (메일쓰기 창)

    # 2. 변수 설정 (이름, 이메일, 메일쓰기 창 내 변수들)
    toInput = driver.find_element(By.CSS_SELECTOR, "#toInput") # 수신
    ccInput = driver.find_element(By.CSS_SELECTOR, "#ccInput") # 참조
    subjectInput = driver.find_element(By.CSS_SELECTOR, "#tbSubject") # 제목
    btSendMail = driver.find_element(By.CSS_SELECTOR, "#btSendMail") # 메일쓰기 버튼
    btAreaAttachDisplay = driver.find_element(By.CSS_SELECTOR, "#btAreaAttachDisplay") # 첨부파일 아래 화살표 버튼
    subject = f"{now}_자동차 리콜 및 산업동향 정보 조회결과" # 메일 내 들어갈 제목
    
    # 3. 수신, 제목 입력
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it(driver.find_element(By.CSS_SELECTOR, "#dext_frame_writeEditor"))) # 본문 활성화 될 때 까지
    time.sleep(0.5)
    driver.switch_to.window(driver.window_handles[1])

    
    
    for l in lst_toInput:
        toInput.send_keys(f"{l}@hlcompany.com" + Keys.ENTER)
    
    time.sleep(1)
    subjectInput.send_keys(subject)

    # 4. 본문 클릭
    # iframe 요소 전환
    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#dext_frame_writeEditor"))
    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#dext5_design_writeEditor"))

    driver.find_element(By.CSS_SELECTOR, "#dext_body > p:nth-child(1)").click() # 본문클릭

    # 5. 본문내용 입력
    actions = ActionChains(driver)
    actions.send_keys(f"안녕하세요").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.send_keys("본 메일은 자동차 리콜 정보 및 자동차 동향을 수집하여 자동으로 발송되는 메일입니다.").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    actions.send_keys(f"상세 내용은 첨부파일 참고 부탁드립니다.").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    
    # 6-1. 내용 입력_뉴스기사
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ 차량 리콜 주요기사 (상위 5개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : 네이버 뉴스 (차량 리콜)").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    for i in range(5):
        actions.send_keys(f"   {i+1} 번째 기사").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ① 언론사 : {df_naver.iloc[i,1]} ({df_naver.iloc[i,3]})").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ② 제목 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_naver.iloc[i,2]}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ③ 내용 : {df_naver.iloc[i,4].split('다.')[0]}다.").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ④ 링크 : {df_naver.iloc[i,5]}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.key_down(Keys.ENTER).perform()

    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()

    # 6-2-1. 자동차 리콜센터 주요내용 입력
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ 국내 리콜 (최근 발행 5개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : 자동차 리콜센터").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    for i in range(5):
        actions.send_keys(f"   {i+1} 번째 리콜").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ① 발행일 : {df_recallcenter.iloc[i,1]}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ② 제조사 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_recallcenter.iloc[i,2]}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.key_down(Keys.ENTER).perform()
        
        if len(df_recallcenter.iloc[i,3].split(",")) == 1:
            actions.send_keys(f"     ③ 차종 / 내용 : ").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
            actions.send_keys(f"{df_recallcenter.iloc[i,3]} / {df_recallcenter.iloc[i,4]}").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        else:
            actions.send_keys(f"     ③ 차종 / 내용 : ").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
            actions.send_keys(f"{df_recallcenter.iloc[i,3].split(',')[0]} 등 {len(df_recallcenter.iloc[i,3].split(','))}차종 / {df_recallcenter.iloc[i,4]}").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
            
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ④ 생산기간 : {df_recallcenter.iloc[i,5]}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ⑤ 대상수량 : {df_recallcenter.iloc[i,6]}대").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.key_down(Keys.ENTER).perform()

    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    
    # 6-2-2. 자동차 리콜센터 주요내용 입력 (무상수리)
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ 국내 무상점검/수리 (최근 발행 5개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : 자동차 리콜센터").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    for i in range(5):
        actions.send_keys(f"   {i+1} 번째 무상점검/수리").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ① 발행일 : {df_recallcenter_무상수리.at[i,'작성일']}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ② 제조사 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_recallcenter_무상수리.at[i,'제작사']}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.key_down(Keys.ENTER).perform()
        
        if len(df_recallcenter.at[i,'차종'].split(",")) == 1:
            actions.send_keys(f"     ③ 차종 / 내용 : ").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
            actions.send_keys(f"{df_recallcenter_무상수리.at[i,'차종']} / {df_recallcenter_무상수리.at[i,'내용']}").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        else:
            actions.send_keys(f"     ③ 차종 / 내용 : ").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
            actions.send_keys(f"{df_recallcenter_무상수리.at[i,'차종'].split(',')[0]} 등 {len(df_recallcenter_무상수리.at[i,'차종'].split(','))}차종 / {df_recallcenter_무상수리.at[i,'내용']}").perform()
            actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ④ 생산기간 : {df_recallcenter_무상수리.at[i,'생산기간']}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ⑤ 대상수량 : {df_recallcenter_무상수리.at[i,'대상수량']}대").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.key_down(Keys.ENTER).perform()

    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    
    
    # 6-3. NHTSA(북미) 리콜 주요내용 입력
    global cnt_nhtsa
    if cnt_nhtsa >= 5:
        cnt_nhtsa = 5
    else:
        pass

    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ 북미 리콜 (최근 발행 {cnt_nhtsa}개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : NHTSA").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()

    for i in range(cnt_nhtsa):
        actions.send_keys(f"   {i+1} 번째 리콜").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ① 이슈일 (Campain No) : {df_nhtsa.at[i,'이슈날짜']} ({df_nhtsa.at[i,'캠페인번호']})").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ② 제조사 / 제품 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_nhtsa.at[i,'제조업체']}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f" / {df_nhtsa.at[i,'제품']}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ③ 내용 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_nhtsa.at[i,'제목']}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ④ 대상수량 : {df_nhtsa.iloc[i,7]}대").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.key_down(Keys.ENTER).perform()

    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    # 6-4. 아우토바인 주요내용 입력
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ 아우토바인 주요기사 (5개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : 아우토바인").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    for i in range(5):
        actions.send_keys(f"   {i+1} 번째 기사").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ① 발행일 : {df_autowein.iloc[i,1]}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ② 제목 : ").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_autowein.iloc[i,2]}").perform()
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.key_down(Keys.ENTER).perform()
        actions.send_keys(f"     ③ 내용 : {df_autowein.iloc[i,3]}").perform()
        actions.key_down(Keys.ENTER).perform()
        actions.key_down(Keys.ENTER).perform()
        
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()
    
    # 6-5. HKMC 대외문 내용 입력
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"■ HKMC 대외문 (최근 10개)").perform()
    actions.key_down(Keys.CONTROL).send_keys('u').send_keys('b').key_up(Keys.CONTROL).perform()
    actions.send_keys(f"   * 출처 : HKMC Autoway").perform()
    actions.key_down(Keys.ENTER).perform()
    actions.key_down(Keys.ENTER).perform()


    for i in range(10):
        actions.send_keys(f"     {df_autoway.iloc[i,1]} : ")
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f"{df_autoway.iloc[i,4]}")
        actions.key_down(Keys.CONTROL).send_keys('b').key_up(Keys.CONTROL).perform()
        actions.send_keys(f" ({df_autoway.iloc[i,2]})")
        actions.key_down(Keys.ENTER).perform()
        
    actions.key_down(Keys.ENTER).perform()
    actions.send_keys(f"이상입니다.").perform()
    

    # 7. 메일 보내기
    driver.switch_to.window(driver.window_handles[1])
    btSendMail.click()
    

    # 8. 메인 창으로 전환
    driver.switch_to.window(driver.window_handles[0])


# 메인함수
def main():
    
    print_program_info(program_version, program_release_date, program_creator) # 프로그램 정보 출력
    # 입력받기
    Hanmaru_ID, Hanmaru_PW, lst_toInput = get_user_input_with_recipients()
    print("인터넷 연결 중입니다.")
    print("잠시만 기다려주세요 (인터넷 환경에 따라 소요시간 상이)")
    add_line()
    
    # 드라이버 셋업
    driver = setup_driver()
    
    # 크롤링
    data_naver = crawl_naver(driver, url_naver)
    data_recallcenter = crawl_recallcenter(driver, url_recallcenter)
    data_recallcenter_무상수리 = crawl_recallcenter_무상수리(driver, url_recallcenter_무상수리)
    data_nhtsa = crawl_nhtsa(driver, url_nthsa)
    data_autowein = crawl_autowein(driver, url_autowein)
    data_autoway = crawl_autoway(driver, url_autoway, "RFP3000", "RFP3000a@")
    
    # 데이터 프레임 저장
    dataframes = {
        "뉴스기사" : data_naver,
        "국내 리콜" : data_recallcenter,
        "국내 무상점검_수리" : data_recallcenter_무상수리,
        "북미 리콜" : data_nhtsa,
        "품질동향" : data_autowein,
        "HKMC 대외문" : data_autoway
    }
    
    # 엑셀 파일로 변환 저장
    save_excel(dataframes, save_file_path)
    print()

    # 한마루 접속 및 로그인
    access_hanmaru(driver, url_hanmaru, Hanmaru_ID, Hanmaru_PW)

    # 메일 본문 작성 및 이메일 발송
    write_email_hanmaru(driver, data_naver, data_recallcenter, data_recallcenter_무상수리, data_nhtsa, data_autowein, data_autoway, lst_toInput)
    time.sleep(1)
    
    print("작업을 완료하였습니다. 프로그램을 종료합니다.")
    print("강제 종료 안하셔도 자동으로 실행창이 닫힙니다.")
    
    # 드라이버 종료
    driver.quit()

main()