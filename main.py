import datetime
import random

import openpyxl
import requests
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
from bs4 import BeautifulSoup
import json
import re
import smtplib  # SMTP 사용을 위한 모듈
from email.mime.multipart import MIMEMultipart  # 메일의 Data 영역의 메시지를 만드는 모듈
from email.mime.text import MIMEText  # 메일의 본문 내용을 만드는 모듈
from email.mime.base import MIMEBase
from email import encoders
import threading
import math


def GetData():
    YOUR_USERNAME="mike102jiro"
    YOUR_PASSWORD="Spc240220==="

    # Define proxy dict. Don't forget to put your real user and pass here as well.
    proxies = {
        'http': 'http://{}:{}@unblock.oxylabs.io:60000'.format(YOUR_USERNAME,YOUR_PASSWORD),
        'https': 'http://{}:{}@unblock.oxylabs.io:60000'.format(YOUR_USERNAME,YOUR_PASSWORD),
    }

    response = requests.request(
        'GET',
        'https://www.okmall.com/products/view?no=689846&item_type=&cate=20008664&uni=M',
        verify=False,  # Ignore the certificate
        proxies=proxies,
    )

    # Print result page to stdout
    print(response.text)


def is_tue_thu_sun():
    # 오늘 날짜 가져오기
    today = datetime.datetime.today().weekday()

    # 화요일(1), 목요일(3), 일요일(6) 중 하나인지 확인
    if today in (1, 3, 6):
        return True
    else:
        return False


def GetGoogleSpreadSheet():
    scope = 'https://spreadsheets.google.com/feeds'
    json = 'credential.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json, scope)
    gc = gspread.authorize(credentials)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1dPaIOsUUKbCOYUy0MjB3lg0UeOsY6yDmbysn52rvz2I/edit#gid=0'
    doc = gc.open_by_url(sheet_url)
    worksheet = doc.worksheet('Sheet1')

    #=================전체정보가져오기
    all_data=worksheet.get_all_records()
    #==================맨 밑행에 데이타 넣기
    # pprint.pprint(all_data)
    dataList=[]
    for data in all_data:
        productNo=data['네이버상품코드']
        productName=data['상품명']
        url = data['상품 링크']
        data={'productNo':productNo,'productName':productName,'url':url}
        dataList.append(data)
    # pprint.pprint(dataList)
    return dataList

def GetInfo(url):
    # cookies = {
    #     'WMONID': 'wCYqcx43ces',
    #     '_fbp': 'fb.1.1704289034744.1442023299',
    #     '_ga': 'GA1.1.192084949.1704289035',
    #     'MyKeyword': '%5B%22%25EB%25A7%2590%25EB%25B3%25B8%22%2C%22FKC334A%2520Y1014%2520415%22%5D',
    #     'imitation_data': '31%2C268%2C270',
    #     'list_ban': '%7B%22on_yn%22%3A1%2C%22on_type%22%3A%22B1%22%7D',
    #     'LastestProduct': '689846%7C693865%7C696164%7C693007%7C717001%7C457689%7C544516%7C743580%7C743205%7C742653%7C725890%7C725879%7C423500%7C446796%7C688429%7C702651%7C724203%7C689177%7C731614%7C744934%7C744920%7C744962%7C690634%7C691955%7C689435%7C688571%7C446792%7C',
    #     '_ga_CW9NG23BGD': 'GS1.1.1708392375.37.1.1708393760.60.0.0',
    #     '_ga_Y9HS705BSQ': 'GS1.1.1708392375.37.1.1708393760.0.0.0',
    #     '_ga_4D8KD9470S': 'GS1.1.1708392375.37.1.1708393760.0.0.0',
    #     'XSRF-TOKEN': 'eyJpdiI6IkRKdWVpQXhoWThzU2xRNTYzeU5mS1E9PSIsInZhbHVlIjoiTFliQlNURkhPb2NQOVM3WGFIclwvRncyaWlRd0lJWXJrRlFcL2hYZTEyVW9LNXBOOWFNd2R6ZHdGMHg2U1VXUnIzIiwibWFjIjoiMWMxMjdkZmQyNDEyZjRmNWMwNDRkODliZGM3MjgxYWY0YThjZjAzYmI5MGUyNjA0ZDViYTNjNTJjZDRmNDhhNiJ9',
    #     'nextokmallweb_session': 'eyJpdiI6InljUUtnZDZSdnBWSXk5ZGx3VEY4b1E9PSIsInZhbHVlIjoiWklyYUNKZ2JFbVZpVlR0V1gxT0Q1Z3JCOWhjc21rT0crd2lmZGNnMHhkcnFmQjhNcjhoaGFaZ21lb3IxZkl0bCIsIm1hYyI6Ijg3ZmE1NDY0ODAwY2Y3NzczMGZlNDRhODQxYjg0YjhmMTNkODJmZDFhOTJhYzlhZDNhNDQwZmVhOTIwMzkyZDIifQ%3D%3D',
    #     'SESSION_GUEST_ID': 'eyJpdiI6InE3b0h1cnpPTk1lMlFTT0lhYUd2NlE9PSIsInZhbHVlIjoiZFJmZUh6dlAxVG5VRXVBc21UaVBiXC81c1oydEFwTHRBSUlzWTNVaWtlbVwvTVgxeXI5V2IrUExKQmZlM1NZK3JkIiwibWFjIjoiMWY4YzcyNWI1NTg1MzY3YTQzMzk4YjYzYjdjZWFjMzFmZTdjN2Y5NGEzYzZhY2FiZjJlODk3MTMwMGYwYTk3ZSJ9',
    # }
    #
    # headers = {
    #     'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    #     'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    #     'Cache-Control': 'max-age=0',
    #     'Connection': 'keep-alive',
    #     # 'Cookie': 'WMONID=wCYqcx43ces; _fbp=fb.1.1704289034744.1442023299; _ga=GA1.1.192084949.1704289035; MyKeyword=%5B%22%25EB%25A7%2590%25EB%25B3%25B8%22%2C%22FKC334A%2520Y1014%2520415%22%5D; imitation_data=31%2C268%2C270; list_ban=%7B%22on_yn%22%3A1%2C%22on_type%22%3A%22B1%22%7D; LastestProduct=689846%7C693865%7C696164%7C693007%7C717001%7C457689%7C544516%7C743580%7C743205%7C742653%7C725890%7C725879%7C423500%7C446796%7C688429%7C702651%7C724203%7C689177%7C731614%7C744934%7C744920%7C744962%7C690634%7C691955%7C689435%7C688571%7C446792%7C; _ga_CW9NG23BGD=GS1.1.1708392375.37.1.1708393760.60.0.0; _ga_Y9HS705BSQ=GS1.1.1708392375.37.1.1708393760.0.0.0; _ga_4D8KD9470S=GS1.1.1708392375.37.1.1708393760.0.0.0; XSRF-TOKEN=eyJpdiI6IkRKdWVpQXhoWThzU2xRNTYzeU5mS1E9PSIsInZhbHVlIjoiTFliQlNURkhPb2NQOVM3WGFIclwvRncyaWlRd0lJWXJrRlFcL2hYZTEyVW9LNXBOOWFNd2R6ZHdGMHg2U1VXUnIzIiwibWFjIjoiMWMxMjdkZmQyNDEyZjRmNWMwNDRkODliZGM3MjgxYWY0YThjZjAzYmI5MGUyNjA0ZDViYTNjNTJjZDRmNDhhNiJ9; nextokmallweb_session=eyJpdiI6InljUUtnZDZSdnBWSXk5ZGx3VEY4b1E9PSIsInZhbHVlIjoiWklyYUNKZ2JFbVZpVlR0V1gxT0Q1Z3JCOWhjc21rT0crd2lmZGNnMHhkcnFmQjhNcjhoaGFaZ21lb3IxZkl0bCIsIm1hYyI6Ijg3ZmE1NDY0ODAwY2Y3NzczMGZlNDRhODQxYjg0YjhmMTNkODJmZDFhOTJhYzlhZDNhNDQwZmVhOTIwMzkyZDIifQ%3D%3D; SESSION_GUEST_ID=eyJpdiI6InE3b0h1cnpPTk1lMlFTT0lhYUd2NlE9PSIsInZhbHVlIjoiZFJmZUh6dlAxVG5VRXVBc21UaVBiXC81c1oydEFwTHRBSUlzWTNVaWtlbVwvTVgxeXI5V2IrUExKQmZlM1NZK3JkIiwibWFjIjoiMWY4YzcyNWI1NTg1MzY3YTQzMzk4YjYzYjdjZWFjMzFmZTdjN2Y5NGEzYzZhY2FiZjJlODk3MTMwMGYwYTk3ZSJ9',
    #     'Sec-Fetch-Dest': 'document',
    #     'Sec-Fetch-Mode': 'navigate',
    #     'Sec-Fetch-Site': 'none',
    #     'Sec-Fetch-User': '?1',
    #     'Upgrade-Insecure-Requests': '1',
    #     'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
    #     'sec-ch-ua': '"Not A(Brand";v="99", "Google Chrome";v="121", "Chromium";v="121"',
    #     'sec-ch-ua-mobile': '?0',
    #     'sec-ch-ua-platform': '"Windows"',
    # }
    #
    # response = requests.get(url, cookies=cookies, headers=headers)

    YOUR_USERNAME="mike102jiro"
    YOUR_PASSWORD="Spc240220==="

    # Define proxy dict. Don't forget to put your real user and pass here as well.
    proxies = {
        'http': 'http://{}:{}@unblock.oxylabs.io:60000'.format(YOUR_USERNAME,YOUR_PASSWORD),
        'https': 'http://{}:{}@unblock.oxylabs.io:60000'.format(YOUR_USERNAME,YOUR_PASSWORD),
    }
    errorCount=0
    while True:
        try:
            response = requests.request(
                'GET',
                url,
                verify=False,  # Ignore the certificate
                proxies=proxies,
                timeout=60
            )
            if response.status_code==200:
                print("수신완료!")
                break
        except:
            print("에러")
            errorCount+=1
        if errorCount>=5:
            break
    soup=BeautifulSoup(response.text,'lxml')
    # print(soup.prettify())
    print("url:",url,"/ url_TYPE:",type(url))
    images=soup.find_all("img")
    isSoldOut="재고있음"
    for image in images:
        if image['src'].find("bx_soldout_rb2.jpg")>=0:
            isSoldOut="품절"
            break
    print("isSoldOut:",isSoldOut,"/ isSoldOut_TYPE:",type(isSoldOut))
    options=soup.find_all("tr",attrs={'name':'selectOption'})

    try:
        price=soup.find("div",attrs={'class':'last_price'}).find('span',attrs={'class':'price'}).get_text().replace(",","")
        # 쉼표를 제외한 연속된 숫자 찾기
        numbers = re.findall(r'\d+', price)[0]
        price=numbers
    except:
        price=""
    print("price:",price)

    optionPriceList=[]
    sizeList=[]
    colorList=[]
    for option in options:
        color=option.find_all('td')[0].get_text()
        colorList.append(color)
        # print("color:",color,"/ color_TYPE:",type(color))
        size = option.find_all('td')[1].get_text()
        # print("size:",size,"/ size_TYPE:",type(size))
        size2 = option.find_all('td')[2].get_text()
        sizeList.append(size+"("+size2+")")
        optionPrice = option.find_all('td')[3].get_text().replace(",","")
        print("optionPrice:",optionPrice,"/ optionPrice_TYPE:",type(optionPrice))
        regex=re.compile("\d+")
        numbers=regex.findall(optionPrice)[-1]
        optionPrice=int(numbers)-int(price)
        print("optionPrice:",optionPrice,"/ optionPrice_TYPE:",type(optionPrice))
        optionPriceList.append(optionPrice)
        # print("optionPrice:",optionPrice,"/ optionPrice_TYPE:",type(optionPrice))
        # print("----------------------------------")
    if len(options)>=1:
        originPrice=int(price)-min(optionPriceList)
    else:
        originPrice = price
    print("originPrice:",originPrice,"/ originPrice_TYPE:",type(originPrice))

    balanceList=[]
    for optionPrice in optionPriceList:
        balanceList.append('10')

    for index,optionPrice in enumerate(optionPriceList):
        if optionPrice!=0:
            optionPriceList[index]=str(optionPriceList[index])

    totalBalance=10*len(balanceList)
    print("totalBalance:",totalBalance,"/ totalBalance_TYPE:",type(totalBalance))

    optionListBalances="\n".join(str(num) for num in balanceList)
    print("optionListBalances:",optionListBalances,"/ optionListBalances_TYPE:",type(optionListBalances))

    optionListPrices = "\n".join(str(num) for num in optionPriceList)
    print("optionListPrices:", optionListPrices, "/ optionListPrices_TYPE:", type(optionListPrices))

    optionListSizes = "\n".join(str(num) for num in sizeList)
    print("optionListSizes:",optionListSizes,"/ optionListSizes_TYPE:",type(optionListSizes))

    optionListColors = "\n".join(str(num) for num in colorList)
    print("optionListColors:",optionListColors,"/ optionListColors_TYPE:",type(optionListColors))

    result=[isSoldOut,optionListColors,optionListSizes,originPrice,optionListPrices,optionListBalances,totalBalance]
    return result

def SendMail(filepath):

    smtp_server = 'smtp.naver.com'
    smtp_port = 587

    # 네이버 이메일 계정 정보
    username = 'sam57892@naver.com'  # 클라이언트 정보 입력
    password = 'Jan240109$$$$'  # 클라이언트 정보 입력

    # receiver='wsgt17@naver.com'
    receiver='mike102jiro@naver.com'
    # receiver=email

    # username = 'hellfir2@naver.com'  # 클라이언트 정보 입력
    # password = 'dlwndwo1!'  # 클라이언트 정보 입력
    # =================커스터마이징
    try:
        to_mail = receiver
    except:
        print("메일주소없음")
        return

    # =================

    # 메일 수신자 정보
    to_email = receiver

    # 참조자 정보
    cc_email = 'ljj3347@naver.com'

    # 메일 본문 및 제목 설정
    contentList=[]

    content="\n".join(contentList)


    # MIMEMultipart 객체 생성
    timeNow=datetime.datetime.now().strftime("%Y년%m월%d일 %H시%M분%S초")
    msg = MIMEMultipart('alternative')
    msg["Subject"] = "[결과]OKMALL 상품 크롤링 (현재시각:{})".format(timeNow)  # 메일 제목
    msg['From'] = username
    msg['To'] = to_email
    msg['Cc'] = cc_email  # 참조 이메일 주소 추가
    msg.attach(MIMEText(content, 'plain'))

    # 파일 첨부
    part = MIMEBase('application', 'octet-stream')
    with open(filepath, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filepath}')
    msg.attach(part)

    # SMTP 서버 연결 및 로그인
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)
    # 이메일 전송 (수신자와 참조자 모두에게 전송)
    to_and_cc_emails = [to_email] + [cc_email]
    server.sendmail(username, to_and_cc_emails, msg.as_string())
    # SMTP 서버 연결 종료
    server.quit()
    print("전송완료")

# filepath='idpw.txt'
# SendMail(filepath)

def process_chunk(chunk, timeNow, result_lock):
    wb = openpyxl.Workbook()
    ws = wb.active
    columnName = ['네이버상품번호', '상품명', '링크', '재고여부(재고있음/품절)', '구매가능색상', '구매가능사이즈', '기본판매가격', '옵션별 가격', '옵션별 재고', '총재고']
    ws.append(columnName)

    for index, inputElem in enumerate(chunk):
        text = f"{index+1}/{len(chunk)}번째 확인중..."
        if len(inputElem['url']) == 0:
            print("없어서스킵")
            time.sleep(1)
            continue
        print(text)
        try:
            infos = GetInfo(inputElem['url'])
        except:
            print("에러로넘어감")
            time.sleep(1)
            continue

        data = [inputElem['productNo'], inputElem['productName'], inputElem['url']]
        data.extend(infos)

        print("data:", data, "/ data_TYPE:", type(data))
        ws.append(data)
        print("====================")
        time.sleep(random.randint(10, 20) * 0.1)

    with result_lock:
        filepath = f'result_{timeNow}_thread_{threading.current_thread().name}.xlsx'
        wb.save(filepath)
        return filepath

count=0
firstFlag=True ## 테스트를 위한 firstFlag
while True:
    timeNow=datetime.datetime.now().strftime("%H%M%S")
    print("현재시각:{}".format(timeNow),"화목일?:",is_tue_thu_sun())
    count+=1
    time.sleep(0.9)
    # if firstFlag==True or (timeNow=="010000" and is_tue_thu_sun()):
    if (timeNow=="010000" and is_tue_thu_sun()) or firstFlag==True:
        # ==========리스트가져오기
        while True:
          try:
            inputList=GetGoogleSpreadSheet()
            break
          except:
            print("가져오기에러1")
          time.sleep(5)
        
        with open('inputList.json', 'w',encoding='utf-8-sig') as f:
            json.dump(inputList, f, indent=2,ensure_ascii=False)

        # ===========디테일가져오기
        with open('inputList.json', "r", encoding='utf-8-sig') as f:
                inputList = json.load(f)

        # Define the number of threads
        num_threads =4  # You can adjust this number based on your needs
        chunk_size = math.ceil(len(inputList) / num_threads)
        chunks = [inputList[i:i + chunk_size] for i in range(0, len(inputList), chunk_size)]

        threads = []
        result_lock = threading.Lock()
        timeNow = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filepaths = []

        for i, chunk in enumerate(chunks):
            thread = threading.Thread(target=process_chunk, args=(chunk, timeNow, result_lock), name=f"Thread-{i+1}")
            threads.append(thread)
            thread.start()

        # Wait for all threads to complete
        for thread in threads:
            thread.join()

        # Collect all filepaths
        for thread in threads:
            filepath = f'result_{timeNow}_thread_{thread.name}.xlsx'
            filepaths.append(filepath)

        # Send emails with all generated files
        for filepath in filepaths:
            SendMail(filepath)

        firstFlag = False