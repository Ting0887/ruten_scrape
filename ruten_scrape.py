import requests
from bs4 import BeautifulSoup
import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter.ttk import *
from PIL import Image,ImageTk
from functools import partial

def SearchProducts(search_name,total_page,file_name,t,progress,window):
    All_ProductsLink = []
    page  = int(total_page.get())
    search_name = search_name.get()
    file_name = file_name.get()
    
    """1 81 161 241"""
    for num in range(1,page+1):
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.193 Safari/537.36',
                   'referer': 'https://find.ruten.com.tw'}
        
        offset = (num-1)*80+1
               
        url = 'https://rtapi.ruten.com.tw/api/search/v3/index.php/core/prod?q='+str(search_name)+'&type=direct&sort=rnk%2Fdc&offset='+str(offset)+'&limit=80'
        print(url)
        res = requests.get(url,headers=headers)
        data = res.json()
        
        for item in data['Rows']:
            link = 'https://www.ruten.com.tw/item/show?' + item['Id']
            All_ProductsLink.append(link)
        print(All_ProductsLink)   
    
    #確認是否有商品資料
    if len(All_ProductsLink)!=0:        
        parse_info(All_ProductsLink,file_name,t,progress,window)
    else:
        msg = '沒有此商品'
        t.insert('insert',msg)

def parse_info(All_ProductsLink,file_name,t,progress,window):
    driverPath = 'D:/chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_experimental_option("detach", True)
    
    service = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=service, options=chrome_options)
    
    #建立檔案
    with open(file_name+'.csv','w',encoding='utf-8-sig',newline='') as csvf:
        w = csv.writer(csvf)
        w.writerow(['link','name','price','sold_num','payway','shipway','stock','seller_board','seller_name'])

    index = 1
    all_progress = 100/len(All_ProductsLink)
    for link in All_ProductsLink:
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
        
        name = ProductName(soup)
        price = ProductPrice(soup)
        sold_num = ProductSoldCount(soup)
        payway = ProductPayway(soup,sold_num)
        shipway = ProductShipway(soup,sold_num)
        stock = ProductStock(soup)
        seller_board = ProductBoard(soup)
        seller_name = ProductSeller(soup)
        pro_img = ProductImage(soup)
        
        result = [link,name,price,sold_num,payway,shipway,stock,seller_board,seller_name,pro_img]
        print(result)

        progress['value']+= all_progress
        window.update_idletasks()        
        
        msg = '第{}筆資料完成!'.format(index)
        t.insert('insert',msg+'\n')
        t.update()
        index += 1
        
        OutputData(result,file_name)
        
    t.insert('insert','已經完成資料爬取!')
    t.update()

def ProductName(soup):
    try:
        name = soup.find('h1','item-title').text.strip() #goods name
    except:
        name = ''
    return name

def ProductPrice(soup):
    try:
        price = soup.find('div','item-purchase-stack').text.replace('直購價：$','') #price
    except:
        price = ''
    return price
    
def ProductSoldCount(soup):
    try:
        sold_num = soup.find('strong','rt-text-x-large number').text #sold num
    except:
        sold_num = 0
    return sold_num

def ProductPayway(soup,sold_num):
    payway = ''
    if sold_num != 0:
        payways = soup.select('table.item-detail-table')[1].find_all('li')
        for p in payways:
            payway += p.text + '、'   
    else:
        payways = soup.select('table.item-detail-table')[0].find_all('li')
        for p in payways:
            payway += p.text + '、'   
    return payway[:-1].replace('\n','').replace('  ','')
        
def ProductShipway(soup,sold_num):
    shipway = ''
    if sold_num !=0:
        shipways = soup.select('table.item-detail-table')[2].find_all('li')
        for s in shipways:
            shipway += s.text + '、'
    else:
        shipways = soup.select('table.item-detail-table')[1].find_all('li')
        for s in shipways:
            shipway += s.text + '、'
    return shipway[:-1].replace('\n','').replace('  ','')

def ProductStock(soup):
    try:
        stock = soup.find('span','rt-ml-2x rt-text-label').text #stock
    except:
        stock = ''
    return stock

def ProductBoard(soup):
    try:   
        seller_board = soup.find('section','seller-board-body').text #seller_board
        #print(seller_board)
    except:
        seller_board = ''
    return seller_board
        
def ProductSeller(soup):
    try:
        seller_name = soup.find('div','seller-board').h3.text.strip() #seller_name
    except:
        seller_name = ''
    return seller_name

def ProductImage(soup):
    try:
        pro_img = soup.find('div','item-gallery-main-image-wrap').img['src']
    except:
        pro_img = ''
    return pro_img

def OutputData(result,file_name):
    with open(file_name+'.csv','a+',newline='',encoding='utf-8-sig') as csvf:
        w = csv.writer(csvf)
        w.writerow(result)

def get_image(file_name,width, height):
    im = Image.open(file_name).resize((width, height))
    return ImageTk.PhotoImage(im)

def main():
    window = tk.Tk()
    window.title("露天拍賣爬蟲")
    window.geometry("1200x800")
    window.configure(background='black')
    window.resizable(width=False,height=False)
    
    """
    canvas = tk.Canvas(window,height=1000,width=1500,bd=0, highlightthickness=0)
    im_root = get_image("D:/black.jpg",width=1500,height=1000)
    canvas.create_image(500, 500, image=im_root)
    canvas.pack()
    """
    #使用者輸入資訊
    L1 = tk.Label(window, bg='yellow', text='商品名稱',font=("SimHei",15))
    L2 = tk.Label(window, bg='yellow', text='頁數',font=("SimHei",15))
    L3 = tk.Label(window, bg='yellow', text='輸出檔名',font=("SimHei",15))
    
    L1.place(x=65,y=100)
    L2.place(x=102,y=150)
    L3.place(x=65,y=200)
    
    b1 = tk.Entry(window, font=("SimHei", 15), show=None, width=35)
    b2 = tk.Entry(window, font=("SimHei", 15), show=None, width=35)
    b3 = tk.Entry(window, font=("SimHei", 15), show=None, width=35)

    b1.place(x=175, y=100)
    b2.place(x=175, y=150)
    b3.place(x=175, y=200)
    
    #進度條
    progress = Progressbar(window,style="green.Horizontal.TProgressbar",orient=tk.HORIZONTAL,length=500,mode='determinate')
    L4 = tk.Label(window, bg='green', text='完成進度',font=("SimHei",15))
    L4.place(x=240,y=295)
    progress.grid(column=0, row=0)
    progress.pack(padx=0,pady=300)
    
    #輸出文字結果
    t = tk.Text(window, width=60, height=10, font=("SimHei", 18), selectforeground='red')  # 顯示多行文字
    t.place(x=100, y=350)
    scrollbar = tk.Scrollbar(window,command=t.yview)
    t.configure(yscrollcommand=scrollbar.set)
    
    button = tk.Button(window,bg='red',text="開始爬取", width=10, height=2, command=partial(SearchProducts,b1,b2,b3,t,progress,window),font=("SimHei", 15))
    button.place(x=850,y=165,anchor=tk.CENTER)
       
    window.mainloop()
    
if __name__ == '__main__':
    main()