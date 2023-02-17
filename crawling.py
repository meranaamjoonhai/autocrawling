# %%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
from collections import defaultdict 

import csv
from tqdm import tqdm

# %%
import os
 
f = open("category.txt",encoding='utf-8')
line = f.readlines()
category_list = line
f.close

for i in range(len(category_list)):
    category_list[i] = category_list[i].strip()

# %%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
from tqdm import tqdm

column = ['업무','공고번호-차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','입찰마감일시','링크']
df = pd.DataFrame(columns=column)
try:
    driver = webdriver.Chrome('./chromedriver')
    driver.implicitly_wait(2)

    for query in category_list:

        
        
        driver.get('https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do')
        task_dict = {'용역': 'taskClCds5'}
        for task in task_dict.values():
            checkbox = driver.find_element(By.ID, task)
            checkbox.click()
        # ID값이 bidNm인 태그 가져오기
        bidNm = driver.find_element(By.ID, 'bidNm')
        # 내용을 삭제
        bidNm.clear()
        # 검색어에 문자열 전달
        bidNm.send_keys(query)
        bidNm.send_keys(Keys.RETURN)
        
        option_dict = {'검색기간 1달': 'setMonth1_1', '입찰마감건 제외': 'exceptEnd', '검색건수 표시': 'useTotalCount'}
        for option in option_dict.values():
            checkbox = driver.find_element(By.ID, option)
            checkbox.click()

        # 목록수 100건 선택 (드롭다운)
        print(1)
        recordcountperpage = driver.find_element(By.NAME,'recordCountPerPage')
        selector = Select(recordcountperpage)
        selector.select_by_value('100')

        search_button = driver.find_element(By.CLASS_NAME, 'btn_mdl')
        search_button.click()
        print(2)
        # 검색 결과 확인
        elem = driver.find_element(By.CLASS_NAME,'results')
        div_list = elem.find_elements(By.TAG_NAME, 'tr')

        # 리스트 형태로 검색결과 저장
        
        print(3)
        results = []
        links = []
        for div in tqdm(div_list):
            results.append(div.text)
            row_data = [div.text.split('\n')]
            a_tags = div.find_elements(By.TAG_NAME,'a')
            if a_tags:
                a_tag = a_tags[0]
                link = a_tag.get_attribute('href')
                row_data.append([link])
            # create a new DataFrame with the data from this row
            
            df_row = pd.DataFrame(row_data)
            # append it to the overall df using pd.concat()
            df = pd.concat([df, df_row], ignore_index=True)
                #df = df.assign(링크 =links)

except Exception as e:
    print(e)
finally:  
    driver.quit()

# %%
df = df[~df[0].isin(['업무 공고번호-차수 분류 공고명 공고기관 수요기관 계약방법 입력일시','검색된 데이터가 없습니다.'])]
df

# %%
column = ['업무','공고번호-차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','입찰마감일시']
new_df= df.iloc[:,10:19]
new_df.columns = column


# %%
aa = new_df.iloc[::2,:]
with_link = pd.DataFrame(new_df.iloc[1::2,:]['업무'])
with_link = list(with_link['업무'])
aa['링크'] = with_link

aa['입력일시'] = pd.to_datetime(aa['입력일시'])
aa = aa.sort_values('입력일시',ascending=False)
aa = aa.drop_duplicates(['공고번호-차수'])
aa = aa.reset_index(drop=True) 


# %%
from datetime import datetime 
today = datetime.today().strftime('%Y-%m-%d')

# %%
aa.to_excel(f'RA_Operational_공고_{today}.xlsx')

# %%
aa

# %%
import win32com.client as win32
from datetime import date
outlook = win32.gencache.EnsureDispatch('Outlook.Application')

# 새 메일 쓰기창 열기
new_mail = outlook.CreateItem(0)
new_mail.Subject = f"{date.today().strftime('%Y-%m-%d 나라장터 입찰 리스트')}"

#수신자
to_list = ['josuh@deloitte.com']
new_mail.To = ";".join(to_list)
cc_list = ['yosohn@deloitte.com']
new_mail.HTMLBody = "this is for test"
new_mail.CC = ";".join(cc_list)
attachment1 = rf'C:\Users\josuh\Desktop\나라장터 크롤링\autocrawling\RA_Operational_공고_{today}.xlsx'
new_mail.Attachments.Add(Source=attachment1)


new_mail.Send()


# 아웃룩 종료
outlook.Quit()

# %%


# %%
