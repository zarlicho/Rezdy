from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pickle
import time
import requests # request img from web
import shutil # save img locally
import urllib.request
from lxml import html
import xlsxwriter
import pandas as pd

outWorkbook = xlsxwriter.Workbook("data.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1","Title")
outSheet.write("B1","Company name")
outSheet.write("C1","Location")
outSheet.write("D1","Company address")
outSheet.write("E1","Contact person")
outSheet.write("F1","Phone")
outSheet.write("G1","Company website")
outSheet.write("H1","availability")
outSheet.write("I1","Company profile")
opt = webdriver.ChromeOptions() 
opt.add_argument("user-data-dir=C:\\chromeprofile")
driver = webdriver.Chrome(executable_path="D:\\Games\\chromedriver.exe", chrome_options=opt)
actions = ActionChains(driver)


def portal2(Clnk):
    driver.get(Clnk)
    response = driver.page_source
    byte_data = response
    uname = ""
    address = ""
    source_code = html.fromstring(byte_data)
    try:
        ards = source_code.xpath("//*[@id='overview']/section[3]/div/div[1]/div[2]/div[2]/a/p")
        adres = ards[0].text_content()
        address=adres
    except:
        ards = source_code.xpath("//*[@id='overview']/section[3]/div/div[1]/div[1]/div[2]/a/p")
        adres = ards[0].text_content()
        address=adres
    cntk = source_code.xpath("//p[@class='truncate']")
    phone = cntk[0].text_content()
    ph = phone.split('Phone: ',1)[1] #Phone 
    print(ph[0:]) #phone
    try:
        usname = source_code.xpath("/html/body/div[1]/div[3]/div[2]/div/div[3]/div[1]/section[3]/div/div[1]/div[2]/div[1]/p[1]")
        uname=usname[0].text_content()
    except:
        usname = source_code.xpath("/html/body/div[1]/div[3]/div[2]/div/div[3]/div[1]/section[3]/div/div[1]/div[1]/div[1]/p[1]")
        uname=usname[0].text_content()
    comEmail = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/div[2]/div/div[3]/div[1]/section[3]/div/div[2]/h4[4]/span/a")))
    compe = comEmail.get_attribute("href")
    driver.back()
    return uname,address,ph[0:], compe
def portal1():
    nilaia=0
    for x in range(1,3):
        ur = "https://app.rezdy.com/marketplace/search?allowNewSearch=1&Product_page={pg}".format(pg=x)
        driver.get(ur)
        fl = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/div/div/div/div/div[3]/div[1]/span/input[2]")))
        fl.send_keys("Australia, South Pacific, Worldwide")
        xa = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/div/div/div/div/strong/strong/button")))
        xa.click()
        time.sleep(4)
        for y in range(1,3):
            nilaia+=1
            title = WebDriverWait(driver, 90).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/strong/strong/div[3]/div/div[3]/section/div/article[{tnum}]/div[1]/h3".format(tnum=y))))
            loc = WebDriverWait(driver, 90).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/strong/strong/div[3]/div/div[3]/section/div/article[{nu}]/div[2]/div[2]/h4/small".format(nu=y))))
            tl=title.text
            compannyName = WebDriverWait(driver, 90).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/strong/strong/div[3]/div/div[3]/section/div/article[{tnum}]/div[2]/div[2]/h4/a".format(tnum=y))))
            name=compannyName.text
            
            lc=loc.text
            # logic here
            state = ""
            response = driver.page_source
            byte_data = response
            source_code = html.fromstring(byte_data)
            try:
                ards = source_code.xpath("/html/body/div[1]/div[3]/form/strong/strong/div[3]/div/div[3]/section/div/article[{indx}]/div[3]/div[2]/div[1]".format(indx=y))
                adres = ards[0].text_content()
                if "Request Negotiated Rate" in adres:
                    state="Request Negotiated Rate"
                elif "CALL" and "BOOK" in adres:
                    state="BOOK"
                elif "BOOK" in adres and "CALL" not in adres:
                    state="BOOK"
                elif "CALL" in adres and "BOOK" not in adres:
                    state="CALL"
            except:
                print("None")
                state="None"
            print(state)
            username,almt,phone,compe = portal2(compannyName.get_attribute("href"))
            outSheet.write(nilaia+1,0,tl)
            outSheet.write(nilaia+1,1,name)
            outSheet.write(nilaia+1,2,lc)
            outSheet.write(nilaia+1,3,almt)
            outSheet.write(nilaia+1,4,username)
            outSheet.write(nilaia+1,5,phone)
            outSheet.write(nilaia+1,6,compe)
            outSheet.write(nilaia+1,7,state)
            outSheet.write(nilaia+1,8,compannyName.get_attribute("href"))
            print(y)
    outWorkbook.close()

portal1()
