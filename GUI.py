#!/usr/bin/env python
# -*- coding: utf-8 -*-
from tkinter import *

from selenium.webdriver.common.keys import Keys

from CenterLib import *

from selenium import webdriver
import time
from bs4 import BeautifulSoup

from google.cloud import translate
import os

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "Key.json"

# Google Cloud client library
from google.cloud import language
from google.cloud.language import enums
from google.cloud.language import types

# Ghi file Excel
import xlsxwriter
import operator
# ---------------------------Khởi tạo giao diện form-----------------------------#
# Khởi tạo giao diện
root = Tk()
root.tk.call('encoding', 'system', 'utf-8')
root.title("GOM NHÓM VÀ PHÂN TÍCH BÌNH LUẬN THEO CHỦ ĐỀ TRÊN FACEBOOK")
root.resizable(height=True, width=True)
root.minsize(height=600, width=460)

# Khởi tạo biến cục bộ
txtContent = StringVar()
txtLinkPost = StringVar()
txtQuantity = StringVar()
txtKeyword = StringVar()
txtLink = StringVar()
# Tạo mảng lưu
lstDone = []
listLinkPost = []
arrayAnalysis = []

# ---------------------------Hàm xử lý------------------------------#
# Function
#Get comments post facebook with link
def getCommentsWithLink(driver,link):
    # Tối ưu link
    arr = link.split("?")
    temp = str(arr[0])
    if temp[len(temp) - 9:len(temp)] == "photo.php":
        arrT = link.split("&")
        temp = str(arrT[0])
    elif temp[len(temp) - 8:len(temp)] == "link.php":
        arrT = link.split("&")
        temp = str(arrT[0] + "&" + arrT[1])
    else:
        temp = link
    if temp != "":
        x = temp.replace("www.facebook", "m.facebook")
        url = x
    i = 0
    j = 0
    listLink = []
    listComments = []
    listSubComments = []
    dem = -1
    #print("*"*200)
    while True:
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, "lxml")
        for a in soup.find_all("div", attrs={'data-sigil': 'comment-body'}):
            try:
                #print(str(j) + " - " + str(i) + " - " + a.text)
                listComments.append(a.text)
            except:
                pass
            j += 1

        #Sub comments post page
        for div in soup.find_all("div",attrs={'class':'_2b1h'}):
            for a in div.find_all("a"):
                listSubComments.clear()
                listSubComments = getCommentsWithLink(driver,"https://m.facebook.com"+a['data-ajaxify-href'])
                for item in listSubComments:
                    listComments.append(str(item))

        for div in soup.find_all("div", attrs={'class': 'async_elem'}):
            try:
                if str(div['id'])[0:8] == "see_next":
                    for a in div.find_all("a", attrs={'class': '_108_'}):
                        url = "https://m.facebook.com" + str(a['href'])
                        listLink.append(url)
            except:
                pass
        dem += 1
        if len(listLink) == dem:
            break
    i += 1
    return set(listComments)
#Search with content
def SearchAction():
    # ---------------------------Lấy link bài post từ Facebook-----------------------------#
    # Cho các gia trị về lại ban đầu
    lstDone.clear()
    listLinkPost.clear()
    listBoxComments.delete(0, END)
    arrayAnalysis.clear()

    #Thiết lập
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path='chromedriver.exe',options=chrome_options)
    #driver.minimize_window()
    url = "https://www.facebook.com/"
    # Điền username và pass Facebook
    username = "trangiahuysn1998@gmail.com"
    password = "trangiahuysn1998"
    driver.get(url)
    # Điền tên đăng nhập và pass
    emailelement = driver.find_element_by_xpath("""//*[@id="email"]""")
    emailelement.send_keys(username)
    passelement = driver.find_element_by_xpath("""//*[@id="pass"]""")
    passelement.send_keys(password)
    # Đăng nhập Facebook
    driver.find_element_by_xpath("""//*[@id="loginbutton"]""").click()
    # Nội dung cần tìm
    content = txtContent.get();

    # Chuyển đến trang Post theo nội dung biến content bên trên
    driver.get("https://www.facebook.com/search/posts/?q=" + content+"&epa=FILTERS&filters=eyJycF9hdXRob3IiOiJ7XCJuYW1lXCI6XCJtZXJnZWRfcHVibGljX3Bvc3RzXCIsXCJhcmdzXCI6XCJcIn0ifQ%3D%3D")

    #Kiểm tra số lượnng
    if txtQuantity.get() == "":
        txtQuantity.set("1")
    #Khởi tạo biến
    listT = []
    count = 0
    countBreak = 0
    while True:
        time.sleep(5)
        elm = driver.find_element_by_tag_name('html')
        # Đợi facebook load các bài đăng lên
        time.sleep(5)
        # Lấy ra đường dẫn các bài đăng
        soup = BeautifulSoup(driver.page_source, "html.parser")
        for span in soup.find_all("span", attrs={'class': '_6-cm'}):
            for a in span.find_all("a"):
                linkPOST = ""
                if a['href'] != "#":
                    if a['href'][0:1] == "/":
                        linkPOST = "https://www.facebook.com" + a['href']
                    else:
                        linkPOST = a['href']
                # Tối ưu link
                arr = linkPOST.split("?")
                temp = str(arr[0])
                if temp[len(temp) - 9:len(temp)] == "photo.php":
                    arrT = linkPOST.split("&")
                    temp = str(arrT[0])
                elif temp[len(temp) - 8:len(temp)] == "link.php":
                    arrT = linkPOST.split("&")
                    temp = str(arrT[0] + "&" + arrT[1])
                if temp != "":
                    x = temp.replace("www.facebook", "m.facebook")
                    listLinkPost.append(str(x))
        if count != 0:
            listT = listLinkPost[count:len(listLinkPost)]
        else:
            countBreak += 1
        countLink = len(listT)
        if countBreak == 3:
            listBoxComments.insert(END,"Không tin thấy")
            break
        if countLink > int(txtQuantity.get()):
            print("*" * 100)
            for item in listT:
                print(item)
            i = 0
            j = 0
            for item in listT:
                # print(item)
                listTemp = getCommentsWithLink(driver, str(item))
                for com in listTemp:
                    try:
                        comTemp = str(j)+ " - "+str(i)+" - "+str(com)
                        lstDone.append(comTemp)
                        listBoxComments.insert(END,comTemp)
                    except: pass
                    j+=1
                i += 1
                if i == (int(txtQuantity.get())):
                    driver.close()
                    break
            break
        # Kéo xuống lấy tiếp bài đăng
        elm.send_keys(Keys.END)
        count = len(listLinkPost)
#Filter keyword in list
def filterKeyword(keyword,list):
    listT = []
    for item in list:
        if str(item).find(keyword) > 0:
            listT.append(item)
    return listT

# ---------------------------Thiết kế giao diện------------------------------#
# Title
Label(root, text="Get Comments Facebook", font=("Tahoma", 20), fg="blue", justify=CENTER).grid(row=0, columnspan=3)

# Search
frameHead = Frame(root)
Label(frameHead, text="Content: ", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Entry(frameHead, width=48, textvariable=txtContent).pack(side=LEFT)
Label(frameHead, text=" Quantity", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Entry(frameHead, width=5, textvariable=txtQuantity).pack(side=LEFT)
Label(frameHead, text=" ", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Button(frameHead, text="Search", command=SearchAction).pack(side=LEFT)
frameHead.grid(row=1, columnspan=3)

# Posts
framePost = Frame(root)
Label(framePost, text="Link: ", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Entry(framePost, width=68, textvariable=txtLinkPost).pack(side=LEFT)
Label(framePost, text=" ", font=("Tahoma", 10), fg="blue").pack(side=LEFT)
def GoLinkAction():
    driver = webdriver.Chrome('chromedriver.exe')
    url = "https://www.facebook.com/"
    # Điền username và pass Facebook
    username = "trangiahuysn1998@gmail.com"
    password = "trangiahuysn1998"
    driver.get(url)
    # Điền tên đăng nhập và pass
    emailelement = driver.find_element_by_xpath("""//*[@id="email"]""")
    emailelement.send_keys(username)
    passelement = driver.find_element_by_xpath("""//*[@id="pass"]""")
    passelement.send_keys(password)
    # Đăng nhập Facebook
    driver.find_element_by_xpath("""//*[@id="loginbutton"]""").click()
    # Chuyển đến trang Post theo nội dung
    driver.get(txtLinkPost.get())
Button(framePost, text="Go to link", command=GoLinkAction).pack(side=LEFT)
framePost.grid(row=2, columnspan=3)

# Filter
frameFilter = Frame(root)
Label(frameFilter, text="Keyword: ", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Entry(frameFilter, width=67, textvariable=txtKeyword).pack(side=LEFT)
Label(frameFilter, text=" ", font=("Tahoma", 5), fg="blue").pack(side=LEFT)
def FilterAction():
    lstT = []
    if txtKeyword.get() == "":
        for item in lstDone:
            lstT.append(item)
    else:
        lstT = filterKeyword(txtKeyword.get(),lstDone)
    listBoxComments.delete(0,END)
    for link in lstT:
        try:
            listBoxComments.insert(END,link)
        except: pass
Button(frameFilter, text="Filter",command=FilterAction).pack(side=LEFT)
frameFilter.grid(row=3, columnspan=3)

#Search with link
frameSearch = Frame(root)
Label(frameSearch, text="Search with link: ", font=("Tahoma", 13), fg="blue").pack(side=LEFT)
Entry(frameSearch, width=57, textvariable=txtLink).pack(side=LEFT)
Label(frameSearch, text=" ", font=("Tahoma", 5), fg="blue").pack(side=LEFT)
def SearchWithLinkAction():
    # Thiết lập
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path='chromedriver.exe', options=chrome_options)
    url = "https://www.facebook.com/"
    # Điền username và pass Facebook
    username = "trangiahuysn1998@gmail.com"
    password = "trangiahuysn1998"
    driver.get(url)
    # Điền tên đăng nhập và pass
    emailelement = driver.find_element_by_xpath("""//*[@id="email"]""")
    emailelement.send_keys(username)
    passelement = driver.find_element_by_xpath("""//*[@id="pass"]""")
    passelement.send_keys(password)
    # Đăng nhập Facebook
    driver.find_element_by_xpath("""//*[@id="loginbutton"]""").click()
    lstT = getCommentsWithLink(driver,txtLink.get())
    lstDone.clear()
    listLinkPost.clear()
    listLinkPost.append(txtLink.get())
    listBoxComments.delete(0,END)
    i = 0
    for item in lstT:
        try:
            lstDone.append(str(i)+" - 0"+ " - "+item)
            listBoxComments.insert(END,str(i)+" - 0"+ " - "+item)
            i+=1
        except: pass
    driver.close()
Button(frameSearch, text="Search",command=SearchWithLinkAction).pack(side=LEFT)
frameSearch.grid(row=4, columnspan=3)

def CurSelet(evt):
    try:
        value = str((listBoxComments.curselection()))
        index = int(value[1:len(value) - 2])
        link = str(listBoxComments.get(index))
        arr = link.split("-")
        content = arr[1].strip()
        item = str(listLinkPost[int(content)]).replace("m.facebook.com","www.facebook.com")
        txtLinkPost.set(item)
    except:
        pass
# ListBox
listBoxComments = Listbox(root, width=90, height=28)
listBoxComments.grid(row=5, columnspan=2, column=0)
listBoxComments.bind('<<ListboxSelect>>', CurSelet)

# Export to Excel
frameFooter = Frame()
def ExportAndAnalysisAcTion():
    # Khởi tạo google translate
    translate_client = translate.Client()
    # Khởi tạo google natural language
    client = language.LanguageServiceClient()

    # Ngôn ngữ muốn dịch
    target = 'en'

    # Tiến hành dịch và Phân tích comments
    # Nếu điểm số từ -1 => - 0.25: tiêu cực
    # Nếu điểm số từ -0.25 => 0.25: có thể tiêu cực hoặc không
    # Nếu điểm số từ -1 => - 0.25: tích cực
    lst = listBoxComments.get(0,END)
    for com in lst:
        try:
            arr = str(com).split("-")
            mess = str(arr[2])
            translation = translate_client.translate(mess.strip(), target_language=target)
            temp = translation['translatedText']
            document = types.Document(content=temp, type=enums.Document.Type.PLAIN_TEXT)
            sentiment = client.analyze_sentiment(document=document).document_sentiment
            rate = ""
            if (float(sentiment.score) <= -0.25 and float(sentiment.score) > -1):
                rate= "Negative"
            elif float(sentiment.score) <= 0.25 and float(sentiment.score) > -0.25:
                rate = "Neutral"
            else:
                rate = "Positive"
            item = mess, sentiment.score, rate
            print(item)
            arrayAnalysis.append(item)
        except:
            pass

    i = 0
    for mess in arrayAnalysis:
        print(str(i) + "-" + str(mess[0]) + "-" + str(mess[1]) + "-" + str(mess[2]))
        i += 1

    # Khởi tạo
    workbook = xlsxwriter.Workbook('KetQua.xlsx')

    # Khởi tạo sheet
    worksheet = workbook.add_worksheet("My sheet")

    # Tạo format
    bold = workbook.add_format({'bold': 1})

    # Ghi title cho cột
    topic = ['Topic: ' + txtContent.get()]
    headings = ['ID', 'Contents', 'Scores', 'Rate']
    worksheet.write_row('A1', topic, bold)
    worksheet.write_row('A2', headings, bold)

    def sort_table(table, col=0):
        return sorted(table, key=operator.itemgetter(col))

    # Ghi dữ liệu
    d = 2
    c = 0
    for mess in sort_table(arrayAnalysis,2):
        worksheet.write(d, c, d - 1)
        c += 1
        worksheet.write(d, c, mess[0])
        c += 1
        worksheet.write(d, c, mess[1])
        c += 1
        worksheet.write(d, c, mess[2])
        c = 0
        d += 1

    workbook.close()

    # Mở file kết quả
    os.startfile("KetQua.xlsx")
Button(frameFooter, text="Export and analysis to Excel file (*.xlsx)", command=ExportAndAnalysisAcTion).pack(
    side=BOTTOM)
frameFooter.grid(row=6, columnspan=3)

Center(root)
root.mainloop()