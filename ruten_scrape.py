import requests
from bs4 import BeautifulSoup
import json
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

def ruten():
    datalink = []
    goodsname = input('input goodsname:')
    page = 400
    for num in range(1,page+1,80):
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.193 Safari/537.36',
                   'referer': 'https://find.ruten.com.tw'}
        payloads = {'q': goodsname,
                    'sort': 'rnk/dc',
                    'offset': num,
                    'limit': '80'
                    }
        url = 'https://rtapi.ruten.com.tw/api/search/v3/index.php/core/prod'
        
        res = requests.get(url,headers=headers,params=payloads)
        data = res.json()
        
        for item in data['Rows']:
            link = 'https://www.ruten.com.tw/item/show?' + item['Id']
            datalink.append(link)
            
    return datalink

def parse_content():
    all_data = []
    driverPath = 'C:\\Program Files\Google\Chrome\Application\chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument('--no-sandbox')
    browser = webdriver.Chrome(driverPath,options=chrome_options)
    
    for link in ruten():
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.193 Safari/537.36',
                    'referer': 'https://find.ruten.com.tw/s/?p=3&q=%E5%B0%8F%E8%B1%86%E6%A2%93',
                    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                    'upgrade-insecure-requests': '1'}
        browser.get(link)
        time.sleep(1.5)
        soup = BeautifulSoup(browser.page_source,'lxml')
        
        #if page is R18
        try:
            browser.find_element(By.CSS_SELECTOR, "#checkAdult > img").click()
            browser.switch_to.alert.text == "您即將進入成人專區，內容包含裸露圖文與限制級商品，請再次確認是否繼續瀏覽 ?"
            browser.switch_to.alert.accept() 
            time.sleep(3)
        except:
            pass
        
        try:
            name = soup.find('h1','item-title').text.strip() #goods name
        except:
            name = ''
            
        try:
            price = soup.find('div','item-purchase-stack').text #price
        except:
            price = ''
            
        try:  
            payway = ''
            payways = soup.select('.item-detail-table')[0].find_all('span','vmiddle')
            for pay in payways:
                payway += pay.text + ' '
            print(payway)
        except:
            payway = ''
            
        try:    
            delivery = ''
            delivery_way = soup.select('.item-detail-table')[1].find_all('span','vmiddle')
            for deli in delivery_way:
                delivery += deli.text + ' '
            print(delivery)
        except:
            delivery = ''
            
        try:
            stock = soup.find('span','rt-ml-2x rt-text-label').text #stock
            print(stock)
        except:
            stock = ''
        try:   
            seller_board = soup.find('section','seller-board-body').text #seller_board
            print(seller_board)
        except:
            seller_board = ''
        
        try:
            seller_name = soup.find('div','seller-board').h3.text.strip() #seller_name
            print(seller_name)
        except:
            seller_name = ''
        
        df = pd.DataFrame([{'link':link,
                            'name':name,
                            'price':price,
                            'payway':payway,
                            'delivery_way':delivery,
                            'stock':stock,
                            'seller_board':seller_board,
                            'seller_name':seller_name}])
        all_data.append(df)  
    return all_data

def outputexcel():
    file = 'ruten' + time.strftime('%Y%m%d') + '.xlsx'
    df1 = pd.concat(parse_content(),ignore_index=True)
    df1.to_excel(file,index=0)
    
outputexcel()
        