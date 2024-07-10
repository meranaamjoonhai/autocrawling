#!/usr/bin/env python
# coding: utf-8

# In[17]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd

from tqdm import tqdm
from pathlib import Path 


# In[18]:


path = Path.cwd()
file = pd.read_excel(path/'keyword.xlsx',engine='openpyxl')
category_list = file.values.tolist()


# In[21]:


column = ['업무','공고번호-차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','입찰마감일시']
df = pd.DataFrame(columns=column)
keywords = []
try:
    driver = webdriver.Chrome()
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
                keywords.append(query)
            # create a new DataFrame with the data from this row
            
            df_row = pd.DataFrame(row_data)
            # append it to the overall df using pd.concat()
            df = pd.concat([df, df_row], ignore_index=True)
                #df = df.assign(링크 =links)

except Exception as e:
    print(e)
finally:  
    driver.quit()


# In[22]:


df = df[~df[0].isin(['업무 공고번호-차수 분류 공고명 공고기관 수요기관 계약방법 입력일시','검색된 데이터가 없습니다.'])]
df


# In[23]:


column = ['업무','공고번호-차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','입찰마감일시','링크']
new_df = df.iloc[:,9:19]
new_df.columns = column


# In[24]:


flatten_keyword = [item for sublist in keywords for item in sublist]


# In[25]:


aa = new_df.iloc[::2,:]
with_link = pd.DataFrame(new_df.iloc[1::2,:]['업무'])
with_link = list(with_link['업무'])
aa['링크'] = with_link
aa.insert(0,'키워드',flatten_keyword)


# In[26]:


aa['입력일시'] = pd.to_datetime(aa['입력일시'])
aa = aa.sort_values('입력일시',ascending=False)
aa = aa.drop_duplicates(['공고번호-차수'])
aa = aa.reset_index(drop=True) 


# In[27]:


from datetime import datetime 
today = datetime.today().strftime('%Y-%m-%d')


# In[28]:

aa.to_excel(f'입찰공고/EYC_RC_{today}.xlsx')

