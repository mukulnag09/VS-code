from multiprocessing.connection import wait
from sys import builtin_module_names
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import openpyxl
import time

options = Options()
options.add_argument("user-data-dir=C:/VS code/chromedriver_win32/")
options.add_argument("chrome.switches" )
options.add_argument("--disable-extensions" )
options.add_argument('profile-directory=Profile 2')
options.add_argument('ignore-certificate-errors')
options.add_argument("--disable-notifications")
path4="C:\\Users\\mukulnag\\Downloads\\Wrongly classified nodes.xlsx"
website1="https://www.adamchoi.co.uk/overs/detailed"
website2="https://netmon/Orion/Report.aspx?ReportID=1043&ReturnTo=aHR0cHM6Ly9uZXRtb24vb3Jpb24vcmVwb3J0cy92aWV3cmVwb3J0cy5hc3B4&showid=9eb81515721e448e9ab5381ac71c17c0"
website3="https://netmon/Orion/Login.aspx?ReturnUrl=%2fOrion%2fReport.aspx%3fReportID%3d1041%26ReturnTo%3daHR0cHM6Ly9uZXRtb24vT3Jpb24vUmVwb3J0cy9EZWZhdWx0LmFzcHg%3d&ReportID=1041&ReturnTo=aHR0cHM6Ly9uZXRtb24vT3Jpb24vUmVwb3J0cy9EZWZhdWx0LmFzcHg="
website4="https://netmon/Orion/Login.aspx?ReturnUrl=%2fOrion%2fReport.aspx%3fReportID%3d1039%26ReturnTo%3daHR0cHM6Ly9uZXRtb24vT3Jpb24vUmVwb3J0cy9EZWZhdWx0LmFzcHg%3d&ReportID=1039&ReturnTo=aHR0cHM6Ly9uZXRtb24vT3Jpb24vUmVwb3J0cy9EZWZhdWx0LmFzcHg="
website5="https://amd.service-now.com/navpage.do"
website6="https://atlsolmpp01.amd.com/ui/perfstack/PSTK-E3AA90C1D53CCAAADC65C1BD51FE5CC046C3048D"
website7="https://outlook.office365.com/mail/"
website8="https://netmon/Orion/SummaryView.aspx?ViewID=1"
website9="https://atl-ipam/ui/"

path="C:\\VS code\\chromedriver_win32\\chromedriver.exe"
s=Service(path)
driver=webdriver.Chrome(service=s,options=options)

driver.get(website9)

"""driver.find_element("xpath",'//input[@name="username"]').send_keys("mukulnag")
driver.find_element("xpath",'//input[@name="password"]').send_keys("45Amd!1101")
b=driver.find_element("xpath",'//input[@class="ib-login-button"]')
b.click()
time.sleep(25)
b=driver.find_element("xpath",'//span[contains(text(),"Add A Record")]')
b.click()
time.sleep(25)
b=driver.find_element("xpath",'//button[@class="button-container"]/span/span[contains(text(),"Select Zone")]')
b.click()
b=driver.find_element("xpath",'//a[contains(text(),"amd.com")]')
b.click()

driver.find_element("xpath",'//input[@name="view:arecord:address"]').send_keys("abcmi=")

driver.find_element("xpath",'//input[@name="view:arecord:fqdn_fake"]').send_keys("1.1.1.1")"""


"""b=driver.find_element("xpath",'//button[@id="ext-gen349"]')
b.click()"""




















"""sp=wbp.active

driver.get(website8)
driver.find_element("xpath",'//input[@name="ctl00$BodyContent$Username"]').send_keys("amd\mukulnag")
driver.find_element("xpath",'//input[@name="ctl00$BodyContent$Password"]').send_keys("45Amd!1101")
b=driver.find_element("xpath",'//a[@id="ctl00_BodyContent_LoginButton"]')
b.click()
for x in range(23,34):
    element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(("xpath", '//i[@icon="search"]'))
        )
    b=driver.find_element("xpath",'//i[@icon="search"]')
    b.click()
    k=sp.cell(x,4).value
    b=driver.find_element("xpath",'//input').send_keys(k)
    if x==23:
        b=driver.find_element("xpath",'//button[@id="button_01R"]')
        b.click()
    else:
        b=driver.find_element("xpath",'//button[@id="button_03D"]')
        b.click()

    element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(("xpath", '//span[@class="xui-highlighted"]'))
        )
    b=driver.find_element("xpath",'//span[@class="xui-highlighted"]')
    b.click()
    element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(("xpath", '//a[@id="Resource16_ctl00_ctl01_Wrapper_customPropertyList_ManageLink"]'))
        )
    b=driver.find_element("xpath",'//a[@id="Resource16_ctl00_ctl01_Wrapper_customPropertyList_ManageLink"]')
    b.click()
    b=Select(driver.find_element("xpath",'//select[@name="ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$adminContentPlaceholder$ctl01$repCustomProperties$ctl05$PropertyValue$RestrictedValues"]'))
    b.select_by_visible_text('Aruba-Switch')
    b=driver.find_element("xpath",'//a[@id="ctl00_ctl00_ctl00_BodyContent_ContentPlaceHolder1_adminContentPlaceholder_imbtnSubmit"]')
    b.click()
    time.sleep(15)

"""














"""

driver.get(website7)
element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(("xpath", '//input[@type="email"]'))
    )
driver.find_element("xpath",'//input[@type="email"]').send_keys("mukulnag@amd.com")

b=driver.find_element("xpath",'//input[@type="submit"]')

b.click()
element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(("xpath", '//button[@aria-label="New mail"]'))
    )
b=driver.find_element("xpath",'//button[@aria-label="New mail"]')
b.click()
"""





    
"""driver.get(website5)
driver.implicitly_wait(20) 
g=driver.find_elements("tag name",'tr')
print(g)
i=0
for p in g:
    if i<10:
        continue
    k=g.find_element("xpath",'./td[0]')
    print(k.text)
"""
















"""all_button = driver.find_element("xpath",'//label[@analytics-event="All matches"]')
all_button.click()
match=driver.find_elements("tag name",'tr')
print(match)
d=[]
h=[]
s=[]
a=[]

for m in match:
    k=m.find_element(By.XPATH,'./td[1]')
    print(k.text)
"""
