# 크롬 드라이버 기본 모듈
import time
import easyocr
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

# 크롬 드라이버 자동 업데이트을 위한 모듈
from webdriver_manager.chrome import ChromeDriverManager

# 브라우저 꺼짐 방지 옵션
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# 불필요한 에러 메시지 삭제
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

# 크롬 드라이버 최신 버전 설정
service = Service(executable_path=ChromeDriverManager().install())

driver = webdriver.Chrome(service=service, options=chrome_options)


# -------------------- 매크로 시작 -------------------- #
# 브라우저 사이즈 조절
driver.set_window_size(1400, 1000)  # (가로, 세로)
driver.get('https://ticket.interpark.com/Gate/TPLogin.asp') # 페이지 이동
# driver.implicitly_wait(10)

# 로그인
driver.switch_to.frame(driver.find_element(By.XPATH, "//div[@class='leftLoginBox']/iframe[@title='login']"))
userId = driver.find_element(By.ID, 'userId')
userId.send_keys('') # 로그인 할 계정 id
userPwd = driver.find_element(By.ID, 'userPwd')
userPwd.send_keys('') # 로그인 할 계정의 패스워드
userPwd.send_keys(Keys.ENTER)
driver.implicitly_wait(10)

# 예매할 티켓의 페이지로 이동
goodsCode = '23012727'
goods_url = 'http://ticket.interpark.com/Ticket/Goods/GoodsInfo.asp?GoodsCode=' + goodsCode 
driver.get(goods_url)

if driver.current_url != goods_url:
    driver.get(goods_url)

time.sleep(1)

# 팝업 닫기
# driver.find_element(By.XPATH, "//*[@id='popup-prdGuide']/div/div[3]/button")
# 팝업이 있는지 확인
try:
    popup_button = driver.find_element(By.XPATH, "//*[@id='popup-prdGuide']/div/div[3]/button")
    popup_button.click()
except NoSuchElementException:
    pass
try:
    popup_button2 = driver.find_element(By.XPATH, "//*[@id='productSide']/div/div[2]/a[1]")
    popup_button2.click()
except NoSuchElementException:
    pass

# 월 체크
calen = driver.find_elements(By.CSS_SELECTOR, ".datepicker-panel")
uls = calen[0].find_elements(By.TAG_NAME, "ul")
year_month = uls[0].find_elements(By.TAG_NAME, "li")[1].text.split('. ')
year = year_month[0]  # 년
month = year_month[1]  # 월

# yearC = int(wantYear) - int(year) # 예매 원하는 년
yearC = int(2023) - int(year)
# monthC = int(wantMonth) - int(month) # 예매 원하는 달
monthC = int(10) - int(month)

prev = uls[0].find_elements(By.TAG_NAME, "li")[0]
next = uls[0].find_elements(By.TAG_NAME, "li")[2]

s = yearC * 12 + monthC
i = 0
if s > 0:
    while i < s:
        next.click()
        i = i + 1
elif s < 0:
    while i < s:
        prev.click()
        i = i - 1

# 선택 가능한 날짜 모두 가져오기
CellPlayDate = driver.find_elements(By.XPATH, "//ul[@data-view='days']/li[@class!='disabled']")
for cell in CellPlayDate:
    # if cell.text == wantDate: # 예매 원하는 일
    if cell.text == "29":
        cell.click()
        break

# 선택 가능한 시간 가져오기
time_li = driver.find_elements(By.XPATH, "//a[@class='timeTableLabel']/span")

# hour_min = hour + ":" + min_
hour_min = "18" + ":" + "00"

for li in time_li:
    if li.text == hour_min:
        li.click()
        break

# 예매하기 클릭
# driver.find_element(By.XPATH, "//*[@id='productSide']/div/div[2]/a[1]").click()
a = driver.find_element(By.CLASS_NAME, "is-primary")
a.click()

# 예매하기 눌러서 팝업창이 뜨면 포커스를 새창으로 바꿔준다
driver.switch_to.window(driver.window_handles[1])

# iframe 이동
time.sleep(1)
driver.switch_to.frame(driver.find_element(By.XPATH, "//*[@id='ifrmSeat']"))

# 입력해야될 문자 이미지 캡쳐하기.
capchaPng = driver.find_element(By.XPATH, "//*[@id='imgCaptcha']")

# easyocr 이미지내 인식할 언어 지정
reader = easyocr.Reader(['en'])

# 캡쳐한 이미지에서 문자열 인식하기
result = reader.readtext(capchaPng.screenshot_as_png, detail=0)

# 이미지에 점과 직선이 포함되어있어서 문자 인식이 완벽하지 않아서 데이터를 수동으로 보정
capchaValue = result[0].replace(' ', '').replace('5', 'S').replace('0', 'O').replace('$', 'S').replace(',', '')\
    .replace(':', '').replace('.', '').replace('+', 'T').replace("'", '').replace('`', '')\
    .replace('1', 'L').replace('e', 'Q').replace('3', 'S').replace('€', 'C').replace('{', '').replace('-', '')

# 입력할 텍스트박스 클릭하기.
driver.find_element_by_class_name('validationTxt').click()

# 추출된 문자열 텍스트박스에 입력하기.
chapchaText = driver.find_element_by_id('txtCaptcha')
chapchaText.send_keys(capchaValue)

#chapchaText.send_keys(Keys.ENTER)