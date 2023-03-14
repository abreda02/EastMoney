#selenium:
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
#for wait times:
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from datetime import date
import numpy as np
import pandas as pd
import re
import glob
import os
import os.path
import requests
import pdfplumber
import sys
import time

today = datetime.today().strftime('%Y-%m-%d') #string 
today_d=date.today()

#setting up paths:
def set_paths():
    #absolute path:
    abs_path=os.path.abspath(os.curdir)
    #folder with the scanned files:
    Scanned_Excels=abs_path+'//EaseMoneyScraped//'
    backup=abs_path+'//BackupTranslations//'
    #folder with updated db:
    DataBase=abs_path+'//DataBase//'
    #folder destinatoin for pdf download (in order to take company name in English)
    testpdf=r"C:\Users\AmirBreda\Downloads\testpdf"
    CompanyName=abs_path+'//CompanyName//'
    return Scanned_Excels,backup,DataBase,testpdf,CompanyName

Scanned_Excels,backup,DataBase,testpdf,CompanyName=set_paths()

#########trying To run chrome-headless: STILL NEED TO WORK ON THIS !!!!

# from selenium.webdriver.chrome.options import Options
# chrome_options = Options()
# #chrome_options.add_argument("--disable-extensions")
# #chrome_options.add_argument("--disable-gpu")
# #chrome_options.add_argument("--no-sandbox") # linux only
# chrome_options.add_argument("--headless")
# chrome_options.add_argument("--start-maximized");
# # chrome_options.headless = True # also works
# driver = webdriver.Chrome(options=chrome_options)

# url='https://data.eastmoney.com/xg/ipo'
# driver.get(url)

# driver.maximize_window()
# driver.implicitly_wait(2)

#setting up driver:
def get_driver():
    options=webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
        "download.default_directory":testpdf,
        "download.prompt_for_download":False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally":True
    })

    driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
    url='https://data.eastmoney.com/xg/ipo'
    driver.get(url)
    driver.maximize_window()
    driver.implicitly_wait(2)
    return driver

#wait till local folder is done with download process:
def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 1000:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

#create empty df with column names:
def get_df_column_names():
    thead='//*[@id="dataview_qb"]/div[2]/div[2]/table/thead'
    thead=driver.find_element(By.XPATH, thead)
    col_str=thead.text
    col_list=list(col_str.split('\n'))
    col_list=col_list[:-1]
    col_list[-1]='CompanyName'
    #add company name here:
    df_empty=pd.DataFrame(columns=col_list)
    return df_empty

def clean_df(df):
    df = df.replace('-', np.NaN)
    df['序号']=df['序号'].astype(int)
    #let's see if now I will have the datetime without stamp:
    df['更新日期']=pd.to_datetime(df['更新日期'])#,format='%Y-%m-%d')
    df['受理日期']=pd.to_datetime(df['受理日期'])#,format='%Y-%m-%d')
    #add column, try with the second line:
    df['AddedOn']=today_d
    df['AddedOn']=pd.to_datetime(df['AddedOn'])
    #remove dups
    col=df.columns[1:-1].to_list()
    df=df.drop_duplicates(subset=col,keep=False)
    return df


def is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

a_l=['Co., Ltd.','CO.LTD','Ltd.','Limited'
          ,'Corp.','Corporation'
          ,'Inc.',',Inc.','Inc','Incorporated'
          ,'INTERNATIONAL'
          ,'Group','Bank'
          ,'Company','Co.,','AG']
a='|'.join(a_l)


#translate chinese company name to english using the PDF:
def get_company_name(url):
    with pdfplumber.open(url) as pdf:
        pages=pdf.pages[:100]
        for page in pages:
            # print (page)
            text=page.extract_text()
            for row in text.split('\n'):
                try:#英文名称/英文全称 
                    # if '英文' in row and re.search('[a-zA-Z]', row):
                    if '英文' in row and re.search(a, row,flags=re.IGNORECASE):
                            company_name=re.sub(r'[^A-Za-z 0-9&?!(),.-]','',row)
                            company_name=company_name.lstrip()
                            pdf.close()
                            return company_name
                except:
                    pdf.close()
                    return "not found"

#loop through the pages,translate company name to english (With pdf) and add to company name table:
driver=get_driver()
last_page_str=driver.find_element(By.XPATH,'//*[@id="dataview_qb"]/div[3]/div[1]/a[7]')
last_page_int=int(last_page_str.text)+1


def scrape_webpage():
    try:
        last_page_str=driver.find_element(By.XPATH,'//*[@id="dataview_qb"]/div[3]/div[1]/a[7]')
        last_page_int=int(last_page_str.text)+1
        box = driver.find_element(By.XPATH,'//*[@id="gotopageindex"]')
        df1=get_df_column_names()
        for i in range(1,last_page_int):
            box.clear()
            box.send_keys(i)
            box.send_keys(Keys.ENTER)
            print (f'moved to page {i}')
            time.sleep(10)
            tbody = WebDriverWait(driver, 90).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="dataview_qb"]/div[2]/div[2]/table/tbody')))
            for t in tbody.find_elements(By.TAG_NAME, "tr" ):
                row_data=t.find_elements(By.TAG_NAME, "td" )
                row = [tr.text for tr in row_data]
                print (row[0])
                length = len(df1)
                df1.loc[length] = row
                print ('added row to df')
    except Exception as e:
        print (f'Exception as {e}')
    finally:
        df2=clean_df(df1)      
        #add remove dups function !  
        df2.to_excel(Scanned_Excels+'EastMoneyTable_updated'+today+'.xlsx'
                    ,index=False)
        return df2

    ##first run (15/2): 40m 25.7s(about 34s per page)
    ##second run (24/2)-38 m



def companyname_from_ch_to_Eng(firstpage=1,lastpage=last_page_int
                               ,xlsx_to_read=r"C:\Users\AmirBreda\OneDrive\Desktop\Boldor\companyname_to_eng_db2023-03-09.xlsx"
                               ,xlsx_to_write=r"C:\Users\AmirBreda\OneDrive\Desktop\Boldor\companyname_to_eng_db"+today+"_1.xlsx"):
    try:
        companyname_df=pd.read_excel(xlsx_to_read)
        last_page_str=driver.find_element(By.XPATH,'//*[@id="dataview_qb"]/div[3]/div[1]/a[7]')
        last_page_int=int(last_page_str.text)+1
        box = driver.find_element(By.XPATH,'//*[@id="gotopageindex"]')

        for i in range(firstpage,lastpage):
            pdf_list=[]
            box.clear()
            box.send_keys(i)
            box.send_keys(Keys.ENTER)
            print (f'move to page {i}')
            time.sleep(15)
            tbody = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="dataview_qb"]/div[2]/div[2]/table/tbody')))
            try:
                for t in tbody.find_elements(By.TAG_NAME, "tr" ):
                        row_data=t.find_elements(By.TAG_NAME, "td" )
                        companyname_from_web=row_data[1].text
                        if companyname_from_web not in companyname_df['企业名称'].unique():
                            html_pdf_name=row_data[-1].get_attribute('innerHTML')
                            html_pdf_splited=html_pdf_name.split('<a href="//pdf.dfcfw.com/pdf/')
                            pdfname=html_pdf_splited[-1].split('"><span class')[0]
                            file_path = os.path.join(testpdf, pdfname)
                            pdf_list.append([companyname_from_web,file_path])
                            # WebDriverWait(driver, 500).until(
                            #                 EC.element_to_be_clickable(
                            #                 (By.XPATH, '//*[@id="dataview_qb"]/div[2]/div[2]/table/tbody/tr[50]/td[12]/a/span')))
                            row_data[-1].click()
                            print (f'clicked index: {row_data[0].text} to add {companyname_from_web}')
                            time.sleep(15)
                        else:
                            print (f'you have this company already: {row_data[0].text, companyname_from_web}')
                            time.sleep(2)
            except Exception as e:
                print (f'see exception: {e}')
                time.sleep(3)
                continue
            w=download_wait(testpdf)
            print (f'waited {w}s ')  
            print (pdf_list)
            for i in pdf_list:
                print (i)
                try:
                    com=get_company_name(i[1])
                    print (f'found eng company: {com}')
                    ##add here:
                    #if com contains ..
                    companyname_df.loc[len(companyname_df)]=list([i[0],com])
                except FileNotFoundError as fnf:
                    print (f'Comapny {i[0]} and file {i[1]} not found as: {fnf}')
                    time.sleep(5)
                    continue
                except Exception as e:
                    print (e)
                    continue
            for f in os.listdir(testpdf):
                os.remove(os.path.join(testpdf, f))
            print ('files removed from dir')
    except Exception as e:
        print (f'this Exception was raised: {e}')
    finally:
        companyname_df.to_excel(xlsx_to_write
                                ,index=False)
        driver.close()
        return companyname_df     
    
def get_categorical_col():
    status=pd.read_excel(backup+"chinese_to_eng_db_09032023.xlsx"
                ,sheet_name='Status')
    por=pd.read_excel(backup+"chinese_to_eng_db_09032023.xlsx"
                ,sheet_name='Place_of_Registration')
    ind=pd.read_excel(backup+"chinese_to_eng_db_09032023.xlsx"
                ,sheet_name='Industry')
    se=pd.read_excel(backup+"chinese_to_eng_db_09032023.xlsx"
                ,sheet_name='StockExchange')
    return status,por,ind,se
    