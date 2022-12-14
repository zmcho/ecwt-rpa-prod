from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import easygui

import calendar
import time
from datetime import datetime
from tkinter import *

# TO EDIT
computer_name = 'Snow'
chrome_profile = 'Profile 5'

month = []
year = []
date = datetime.now()

# Variables (DO NOT EDIT!!)
error = 0
start = 0
custom = 0
loop = 0


def byquarter1():
    global month
    month = ['January', 'February', 'March']
    qwindow.destroy()
    window.destroy()


def byquarter2():
    global month
    month = ['April', 'May', 'June']
    qwindow.destroy()
    window.destroy()


def byquarter3():
    global month
    month = ['July', 'August', 'September']
    qwindow.destroy()
    window.destroy()


def byquarter4():
    global month
    month = ['October', 'November', 'December']
    qwindow.destroy()
    window.destroy()


def byquarter():
    global qwindow
    qwindow = Tk()
    q1 = Button(qwindow, text="Quarter 1", command=byquarter1)
    q1.place(x=125, y=0)
    q2 = Button(qwindow, text="Quarter 2", command=byquarter2)
    q2.place(x=125, y=50)
    q3 = Button(qwindow, text="Quarter 3", command=byquarter3)
    q3.place(x=125, y=100)
    q4 = Button(qwindow, text="Quarter 4", command=byquarter4)
    q4.place(x=125, y=150)
    qwindow.title('Quarter')
    qwindow.eval('tk::PlaceWindow . center')
    qwindow.geometry("300x200")


def bymonth():
    mwindow = Tk()
    listbox = Listbox(mwindow, width=12, height=12, exportselection=FALSE, selectmode=MULTIPLE)

    listbox.insert(1, "January")
    listbox.insert(2, "February")
    listbox.insert(3, "March")
    listbox.insert(4, "April")
    listbox.insert(5, "May")
    listbox.insert(6, "June")
    listbox.insert(7, "July")
    listbox.insert(8, "August")
    listbox.insert(9, "September")
    listbox.insert(10, "October")
    listbox.insert(11, "November")
    listbox.insert(12, "December")
    listbox.place(x=20, y=50)

    listboxyear = Listbox(mwindow, width=12, height=3, exportselection=FALSE, selectmode=SINGLE)

    listboxyear.insert(1, date.year - 2)
    listboxyear.insert(2, date.year - 1)
    listboxyear.insert(3, date.year)
    listboxyear.place(x=100, y=50)

    def selected_item():
        global custom
        custom = 1
        for i in listbox.curselection():
            global month
            month.append(listbox.get(i))
        for i in listboxyear.curselection():
            global year
            year.append(listboxyear.get(i))
        mwindow.destroy()
        window.destroy()

    custom = 1
    ok = Button(mwindow, text="OK", command=selected_item)
    ok.place(x=200, y=220)
    mwindow.title('Month')
    mwindow.eval('tk::PlaceWindow . center')
    mwindow.geometry("250x300")


##DIALOG BOX MAIN##
global window
window = Tk()
btn = Button(window, text="By Quarter", command=byquarter)
btn.place(x=50, y=100)
btn = Button(window, text="By Month", command=bymonth)
btn.place(x=170, y=100)
lbl = Label(window, text="How do you wish to run the bot?", fg='white', font=("Helvetica", 14), bg='#1C3D80')
lbl.place(x=20, y=50)

window.title('eCWT')
window.geometry("300x200")
window.configure(bg='#1C3D80')
window.eval('tk::PlaceWindow . center')
window.mainloop()

##WEB AUTOMATION##
print("START: " + str(datetime.now()))
# url = 'https://globetelecominc-tst.outsystemsenterprise.com/eCWT_Portal/Login_SSO.aspx'
url = 'https://ecwt.globe.com.ph/portal/Login_SSO.aspx' #PROD

options = Options()
options.add_experimental_option("detach", True)
options.add_argument(r'user-data-dir=C:\Users\Snow\AppData\Local\Google\Chrome\User Data')
options.add_argument('--profile-directory=Profile 5')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get(url)

# LOGIN
time.sleep(1)
driver.implicitly_wait(60)
driver.find_element(By.XPATH, '//*[@id="wt1"]/div/div[2]').click()

# Navigation
time.sleep(1)
driver.implicitly_wait(60)
# driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt207_block_wtMenu_wt109_RichWidgets_wt55_block_wtMenuItem_wt132"]').click()
driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt207_block_wtMenu_wt109_RichWidgets_wt56_block_wtMenuItem_wt124"]').click()  # PRODa
time.sleep(1)
# driver.find_element(By.XPATH, '//*[@href="PDFUploading.aspx?(Not.Licensed.For.Production)="]').click()
driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt207_block_wtMenu_wt109_RichWidgets_wt56_block_wtMenuSubItems_wt87"]').click()  # PROD

driver.find_element(By.XPATH, '//*[@data-identifier="zmcho@globe.com.ph"]').click()
driver.find_element(By.XPATH, '//*[@id="submit_approve_access"]/div/button').click()

# PDF Uploading
while loop != len(month):
    # for i in range(len(month)):
    if error == 1:
        driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_WebPatterns_wt13_block_RichWidgets_wt9_block_wt33"]').click()
    mnum = datetime.strptime(month[loop], '%B').month
    driver.implicitly_wait(60)
    driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn1_wtDateStart"]').clear()
    driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn2_wtDateEnd"]').clear()
    if custom == 1:
        driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn1_wtDateStart"]').send_keys(str(year[0]) + "-" + str(mnum) + "-01")
        driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn2_wtDateEnd"]').send_keys(str(year[0]) + "-" + str(mnum) + "-" + str(calendar.monthrange(year[0], mnum)[1]))
    else:
        driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn1_wtDateStart"]').send_keys(str(date.year) + "-" + str(mnum) + "-01")
        driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt94_block_wtMainContent_WebPatterns_wt104_block_wtColumn2_wtDateEnd"]').send_keys(str(date.year) + "-" + str(mnum) + "-" + str(calendar.monthrange(date.year, mnum)[1]))
    driver.implicitly_wait(60)
    driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt65"]').click()
    time.sleep(5)
    if start == 0:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt68"]').click()
        time.sleep(5)
    driver.implicitly_wait(60)
    driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').clear()
    if mnum == 1:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1y1g2H2XHSjhkFFGpQaNpTQ_vwBb9blrI")
    elif mnum == 2:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1fA7egtwsI4ldtIOuPxZ_H9P7NzSOZ1jS")
    elif mnum == 3:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1A7oYNgBP2BSaPBWXnG4qRpAf1B49fNhm")
    elif mnum == 4:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1FTpLiSi7OC5l-IvhDBwwMY60RUdT3yRW")
    elif mnum == 5:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1iAZEbg_ynGKB6jch8iRvKaMBgrEY0jF7")
    elif mnum == 6:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1EH06vdzs3VeLPmEbZUMakYl2sGT8EE9z")
    elif mnum == 7:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1DLvEG8Uz1yP56SEUi28rDzeDPGaaUdtg")
    elif mnum == 8:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1uU3XKt05whmSY0OgTtS6Z1dAGMLCtLID")
    elif mnum == 9:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1VnBupVcr2-WCa4O9noxiMCFKeuA3F5_q")
    elif mnum == 10:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("17UR_zrbVYuROj1CAbpRvg8WD9AuMC26C")
    elif mnum == 11:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1fekjbT25drRuRy_2F0nUZbIjwWGZsYtJ")
    elif mnum == 12:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys("1yJ_5BGZRq1cw3Ihx48TxpiX77w--yFkO")
    driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt4"]').click()
    # try:
    #     element = WebDriverWait(driver, 3600).until(
    #         EC.visibility_of_element_located((By.XPATH, '//*[text()="No ecwtpdf uploads to show..."]'))
    #     )
    # finally:
    #     loop=loop-1
    #     pass
    try:
        # element = WebDriverWait(driver, 3600).until(
        #     #EC.visibility_of_element_located((By.XPATH, '//*/span[text()="An exception occurred in the client script. Error: The connection to the server was reset. Server returned status Internal Server Error"]'))
        #     EC.visibility_of_element_located((By.XPATH, '//*[@id="LisbonTheme_wt94_block_WebPatterns_wt13_block_RichWidgets_wt9_block_wtSanitizedHtml2"]'))
        # )
        element = WebDriverWait(driver, 3600).until_not(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="LisbonTheme_wt94_block_WebPatterns_wt13_block_wt13_wtdivWait"]'))
        )

    finally:
        try:
            element = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, '//*[text()="No ecwtpdf uploads to show..."]'))
            )
        except:
            start = 1
        else:
            loop = loop + 1
        finally:
            pass
print("END: " + str(datetime.now()))
driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMenu_wt76_wt9_wtLogoutLink"]').click()
driver.quit()
easygui.msgbox("Upload Done.")
