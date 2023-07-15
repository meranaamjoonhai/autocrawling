from flask import Flask, request, render_template, send_file, session,make_response
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from tqdm import tqdm
from datetime import datetime
import io

app = Flask(__name__)
app.secret_key = 'SECRET_KEY'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 키워드 엑셀 파일 처리
        keyword_file = request.files['keyword_file']
        keyword_df = pd.read_excel(keyword_file, engine='openpyxl')
        category_list = keyword_df.values.tolist()

        # 결과 데이터프레임 초기화
        column = ['업무', '공고번호-차수', '분류', '공고명', '공고기관', '수요기관', '계약방법', '입력일시', '입찰마감일시','링크']
        df = pd.DataFrame(columns=column)
        keywords = []

        # WebDriver 초기화
        driver = webdriver.Chrome()
        driver.implicitly_wait(2)

        try:
            for query in category_list:
                driver.get('https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do')
                task_dict = {'용역': 'taskClCds5'}
                for task in task_dict.values():
                    checkbox = driver.find_element(By.ID, task)
                    checkbox.click()

                bidNm = driver.find_element(By.ID, 'bidNm')
                bidNm.clear()
                bidNm.send_keys(query)
                bidNm.send_keys(Keys.RETURN)

                option_dict = {'검색기간 1달': 'setMonth1_1', '입찰마감건 제외': 'exceptEnd', '검색건수 표시': 'useTotalCount'}
                for option in option_dict.values():
                    checkbox = driver.find_element(By.ID, option)
                    checkbox.click()

                recordcountperpage = driver.find_element(By.NAME, 'recordCountPerPage')
                selector = Select(recordcountperpage)
                selector.select_by_value('100')

                search_button = driver.find_element(By.CLASS_NAME, 'btn_mdl')
                search_button.click()

                elem = driver.find_element(By.CLASS_NAME, 'results')
                div_list = elem.find_elements(By.TAG_NAME, 'tr')

                for div in tqdm(div_list):
                    row_data = [div.text.split('\n')]
                    a_tags = div.find_elements(By.TAG_NAME, 'a')
                    if a_tags:
                        a_tag = a_tags[0]
                        link = a_tag.get_attribute('href')
                        row_data.append([link])
                        keywords.append(query)
                    df_row = pd.DataFrame(row_data)
                    df = pd.concat([df, df_row], ignore_index=True)

        except Exception as e:
            print(e)

        finally:
            driver.quit()

        # 데이터 전처리
        df = df[~df[0].isin(['업무 공고번호-차수 분류 공고명 공고기관 수요기관 계약방법 입력일시','검색된 데이터가 없습니다.'])]
        new_df = df.iloc[:,10:19]
        new_df.columns = ['업무','공고번호-차수','분류','공고명','공고기관','수요기관','계약방법','입력일시','입찰마감일시']

        flatten_keyword = [item for sublist in keywords for item in sublist]

        aa = new_df.iloc[::2,:]
        with_link = pd.DataFrame(new_df.iloc[1::2,:]['업무'])
        with_link = list(with_link['업무'])
        aa['링크'] = with_link
        aa.insert(0,'키워드',flatten_keyword)

        aa['입력일시'] = pd.to_datetime(aa['입력일시'])
        aa = aa.sort_values('입력일시',ascending=False)
        aa = aa.drop_duplicates(['공고번호-차수'])
        aa = aa.reset_index(drop=True)
        
        today = datetime.today().strftime('%Y-%m-%d')

        session['aa'] = aa.to_json()

        # 결과를 웹 페이지에 표시
        return render_template('result.html', data=aa.to_dict('records'),keywords=category_list)

    return render_template('index.html')

@app.route('/download')
def download():
    aa = pd.read_json(session['aa'])

    # Generate the Excel file
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    aa.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    # Send the file as a response
    output.seek(0)
    response = make_response(output.getvalue())
    response.headers['Content-Disposition'] = 'attachment; filename= bid_list.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response

if __name__ == '__main__':
    app.run(debug=True)
