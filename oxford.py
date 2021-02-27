import pyautogui
import tkinter as tk
import requests
import bs4
import re
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import os

options = webdriver.ChromeOptions()
options.add_argument("--incognito")

chrome_driver_path = "C:\P\python\chromedriver.exe"
driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)

driver.get("https://home.oxfordowl.co.uk/reading/free-ebooks/")
links = driver.find_elements_by_css_selector(".et_pb_text_inner ul li a")


levelfolder = 'Oxford Level 16'
filepath = "H:\\oxfordowl\\" + levelfolder

if not os.path.isdir(filepath):
    os.mkdir(filepath)


def takes(foldername, filename):
    subfolder = filepath +"\\" + foldername
    if not os.path.isdir(subfolder):
        os.mkdir(subfolder)
    img = pyautogui.screenshot(imageFilename=subfolder + "\\" + filename + ".png", region=(8, 120, 1250, 840)) #80, 120, 1120, 800

def loginpage():
    try:
        username = driver.find_element_by_css_selector(
        "#parent-eacloginbox #eacLoginWidgetContainer #eacLoginBody #eacBasicLoginForm #eacLoginUsername #eacUsername")
        username.send_keys("username")
        password = driver.find_element_by_css_selector(
        "#parent-eacloginbox #eacLoginWidgetContainer #eacLoginBody #eacBasicLoginForm #eacLoginPassword #eacPassword")
        password.send_keys("password")
        loginbutton = driver.find_element_by_css_selector(
        "#parent-eacloginbox #eacLoginWidgetContainer #eacLoginBody #eacBasicLoginForm #eacLoginButton #eacLoginSubmit")
        loginbutton.click()
    except:
        sleep(2000)
        loginpage()

def turnpage_takescreenshot(foldername):
    for i in range(19): #36
        sleep(10)
        takes(foldername, str(i))
        try:
            nextbutton = driver.find_element_by_css_selector(".pagination .go-next div") #div
        except NoSuchElementException:
            try:
                nextbutton = driver.find_element_by_css_selector(".pagination .go-next")
            except NoSuchElementException:
                nextbutton = driver.find_element_by_css_selector(".center .group6 .go-next div")  # div
        nextbutton.click()

def selectbooks():
    counter = 2
    sleep(5)
    books = driver.find_elements_by_css_selector(".bookshelf li")
    #for book in books:
    for i in range(1):
        tagclass = ".pos-" + str(counter) + "-e"
        bk = driver.find_element_by_css_selector(".bookshelf li "+tagclass)
        bk.click()
        #driver.switch_to.window(driver.window_handles[1])
        ## e1 = driver.find_element_by_css_selector(".read-ebook-link")
        sleep(2)
        e1 = driver.execute_script("document.getElementsByClassName('read-ebook-link')[0].click()")

        # inside book
        driver.switch_to.window(driver.window_handles[2])
        driver.fullscreen_window()
        sleep(5)

        #nextbutton = driver.find_element_by_xpath('//*[@id="content-0"]/div/div/div/div[3]/div/div[2]/div[6]/button[2]/div')
        turnpage_takescreenshot(str(counter))
        driver.close()
        driver.switch_to.window(driver.window_handles[1])
        closebutton = driver.find_element_by_css_selector(".popover-close")
        closebutton.click()
        sleep(2)
        counter += 1


for link in links:
    if (link.text == levelfolder):
        driver.execute_script("arguments[0].click();", link)
        driver.switch_to.window(driver.window_handles[1])
        login = driver.find_element_by_link_text("Log in")
        login.click()
        parents = driver.find_element_by_link_text("Parents")
        parents.click()

        # login page
        # if error, just wait 5 minutes
        loginpage()


        #select books
        driver.switch_to.window(driver.window_handles[1])
        selectbooks()


        # inside book
        # driver.switch_to.window(driver.window_handles[2])
        # driver.fullscreen_window()
        # sleep(5)
        # nextbutton = driver.find_element_by_css_selector(".go-next div")
        # #nextbutton = driver.find_element_by_xpath('//*[@id="content-0"]/div/div/div/div[3]/div/div[2]/div[6]/button[2]/div')
        # nextbutton.click()
        # turnpage_takescreenshot()
        break

#driver.close()
#driver.quit()

# le = tk.StringVar()
# top = tk.StringVar()
# width = tk.StringVar()
# height = tk.StringVar()
# p = tk.StringVar()
# url = tk.StringVar()
#
# tk.Label(root, text = "left").grid(row = 0, column=0, sticky='E')
# tk.Entry(root, textvariable=le).grid(row = 0, column = 1, sticky='W')
#
#
# tk.Label(root, text = "top").grid(row = 1, column=0, sticky='E')
# tk.Entry(root, textvariable=top).grid(row = 1, column = 1, sticky='W')
#
#
# tk.Label(root, text = "width").grid(row = 2, column=0, sticky='E')
# tk.Entry(root,textvariable=width).grid(row = 2, column = 1, sticky='W')
#
#
# tk.Label(root, text = "height").grid(row = 3, column=0, sticky='E')
# tk.Entry(root,textvariable=height).grid(row = 3, column = 1, sticky='W')
#
# tk.Label(root, text = "File Path").grid(row = 5, ipady=23)
# tk.Entry(root,textvariable=p, width="65").grid(row = 5, column = 1)
#
#
# tk.Label(root, text = "MP3 URL").grid(row = 9, ipady=40)
# tk.Entry(root,textvariable=url, width="65").grid(row = 9, column = 1)
#
# def takes():
#     lv = le.get()
#     tv = top.get()
#     wv = width.get()
#     hv = height.get()
#     pv = p.get()
#     img=pyautogui.screenshot(imageFilename=r""+pv+".png", region=(float(lv),float(tv), float(wv), float(hv)))
#
#
# def quitApp():
#     root.destroy()
#
#
# def download_mp3():
#     pv = p.get()
#     pv = pv[:-2]
#     uv = url.get()
#
#     for x in range(1, 12):
#         new_url = uv.replace('2', str(x))
#         print(new_url)
#         doc = requests.get(new_url)
#         with open('' + pv + str(x) + '.mp3', 'wb') as f:
#             f.write(doc.content)
#
#
# button1=tk.Button(root, text="Take ScreenShot", command=takes)
# button1.place(x=60,y=130)
#
# button3=tk.Button(root, text="Download mp3", command=download_mp3)
# button3.place(x=60,y=215)
#
# button2=tk.Button(root, text="Quit", command=quitApp)
# button2.place(x=410,y=270)
#
# root.mainloop()