import pyautogui
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


levelfolder = 'Oxford Level 1'
filepath = "H:\\oxfordowl\\" + levelfolder

if not os.path.isdir(filepath):
    os.mkdir(filepath)


def takes(foldername, filename):
    subfolder = filepath +"\\" + foldername
    if not os.path.isdir(subfolder):
        os.mkdir(subfolder)
    img = pyautogui.screenshot(imageFilename=subfolder + "\\" + filename + ".png", region=(10, 120, 1250, 800)) #80, 120, 1120, 805
                                                                                                               #9, 120, 1250, 820
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
        sleep(900)
        loginpage()

def turnpage_takescreenshot(foldername):
    for i in range(9): # pages: 36
        sleep(10)
        takes(foldername, str(i))
        try:
            nextbutton = driver.find_element_by_css_selector(".pagination .go-next div") #div
        except NoSuchElementException:
            nextbutton = driver.find_element_by_css_selector(".pagination .go-next")
        except NoSuchElementException:
            nextbutton = driver.find_element_by_css_selector(".center .group6 .go-next div")  # div
        nextbutton.click()

def selectbooks():
    counter = 0 #0   6
    sleep(5)
    books = driver.find_elements_by_css_selector(".bookshelf li")
    booktitles = ( "Family Poems")

    for book in books:
    #for i in booktitles:
    #for i in range(1):
        tagclass = ".pos-" + str(counter) + "-e"
        bk = driver.find_element_by_css_selector(".bookshelf li "+tagclass)
        #find by title
        #bk = driver.find_element_by_css_selector(f"[title^='{booktitles[counter-6]}']")
        bk.click()
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
