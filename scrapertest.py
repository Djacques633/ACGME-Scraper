import xlsxwriter
import requests
import time
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import pyautogui

browser = webdriver.Chrome("C:/Users/dj214316/Desktop/Newfolder/Newfolder/chromedriver.exe")
browser.get("https://apps.acgme.org/ads/Public/Programs/Search")
button2 = browser.find_element_by_link_text("Search by Specialty")
button2.click()
specialties = browser.find_elements_by_class_name("select2-result-label")
specialty = []
for x in range(1,len(specialties)):
    specialty.append(str(specialties[x].text))
browser.close()
f = xlsxwriter.Workbook('master6.xlsx')
sheet1 = f.add_worksheet()
sheet1.write('A1', 'Specialty')
sheet1.write('B1', 'Title')
sheet1.write('C1', 'Address')
sheet1.write('D1', 'Website')
sheet1.write('E1', 'Phone')
sheet1.write('F1', 'Email')
sheet1.write('G1', 'Director')
sheet1.write('H1', 'Director Appointment Date')
sheet1.write('I1', 'Cordinator')
sheet1.write('J1', 'Cordinator Phone Number')
y = 2
numOfFails = 0
for x in range(0,len(specialty)):
    browser = webdriver.Chrome("C:/Users/dj214316/Desktop/Newfolder/Newfolder/chromedriver.exe")
    sheet1.write('A' + str(y), specialty[x])
    y += 1
    browser.get("https://apps.acgme.org/ads/Public/Programs/Search")
    browser.maximize_window()
    time.sleep(3)
    button = browser.find_element_by_link_text("Search by State")
    button.click()
    browser.find_element_by_link_text("Search by State").send_keys("Ohio")
    browser.find_element_by_link_text("Search by State").send_keys(Keys.ENTER)

    button2 = browser.find_element_by_link_text("Search by Specialty")
    button2.click()
    browser.find_element_by_link_text("Search by Specialty").send_keys(specialty[x])    #Search for the next specialty
    browser.find_element_by_link_text("Search by Specialty").send_keys(Keys.ENTER)
    time.sleep(1)
    button3 = browser.find_elements_by_class_name("listview-filter-accept-button")
    button3[1].click()
    time.sleep(1)
   # pyautogui.scroll(-500)
    time.sleep(1)       #1100, 602
    for i in range(0,100):
        odd_element = browser.find_elements_by_class_name("odd")
        even_element = browser.find_elements_by_class_name("even")
        if i % 2 == 1:
            try:
                hover = ActionChains(browser).move_to_element(odd_element[int(i/2)])
            except:
                continue
        else:
            try:
                hover = ActionChains(browser).move_to_element(even_element[int(i/2)])
            except:
                continue
        #pyautogui.moveTo(100,100)
        #pyautogui.moveTo(1150,815 + i * 70)  #The first row mouse coordinates to make "View Program" visible
        hover.perform()
        data = browser.find_elements_by_link_text("View Program")
        try:
            data[0].click() #If data[0].click works, then that means there was another program to view
        except:
            #Try scrolling to find the next link
            
            
            
            #pyautogui.moveTo(100,100)
            #pyautogui.moveTo(1150,815 + (i - numOfFails) * 70)
            time.sleep(1)
           # pyautogui.scroll(-63)
            time.sleep(1)
            data = browser.find_elements_by_link_text("View Program")
            try:
                data[0].click()
            except:
                numOfFails = 0
                break





        
        html_source = browser.page_source
        title = str(browser.find_element_by_tag_name('h1').text)
        try:
            address = str(browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[1]/div/div[3]/address").text)
        except:
            address = "Could not find"
        try:
            website = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[3]/a").text)
        except:
            website = "Could not find"
        try:
            phone = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[3]/dl[3]/dd[1]").text)
        except:
            phone = "Could not find"
        try:
            email = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[3]/dl[3]/dd[2]/a").text)
        except:
            email = "Could not find"
        try:
            director = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[4]/ul/li[1]").text)
        except:
            director = "Could not find"
        try:
            directapp = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[4]/dl/dd").text)
        except:
            directapp = "Could not find"
        try:
            cord = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[5]/ul/li[1]").text)
        except:
            cord = "Could not find"
        try:
            cord_phone = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[5]/dl/dd[1]").text)
        except:
            cord_phone = "Could not find"
        try:
            cord_email = str(browser.find_element_by_xpath("//*[@id='content-panel']/div[5]/dl/dd[2]/a").text)
        except:
            cord_email = "Could not find"

        
        title = " ".join(title.split())
        address = " ".join(address.split())
        website = " ".join(website.split())
        phone = " ".join(phone.split())
        email = " ".join(email.split())
        director = " ".join(director.split())
        directapp = " ".join(directapp.split())
        cord = " ".join(cord.split())
        cord_phone = " ".join(cord_phone.split())
        cord_email = " ".join(cord_email.split())

        sheet1.write('A' + str(y), specialty[x])
        sheet1.write('B' + str(y), title)
        sheet1.write('C' + str(y), address)
        sheet1.write('D' + str(y), website)
        sheet1.write('E' + str(y), phone)
        sheet1.write('F' + str(y), email)
        sheet1.write('G' + str(y), director)
        sheet1.write('H' + str(y), directapp)
        sheet1.write('I' + str(y), cord)
        sheet1.write('J' + str(y), cord_phone)
        sheet1.write('K' + str(y), cord_email)     
        y += 1
        browser.back()
        time.sleep(3)
    browser.close()
f.close()
