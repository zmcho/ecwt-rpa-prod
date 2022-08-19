from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import easygui

import ctypes
import sys

import calendar
import time
from datetime import datetime
from tkinter import *

######################TO EDIT###########################################
chrome_path = r'user-data-dir=C:\Users\default.DESKTOP-N05I6I4\AppData\Local\Google\Chrome\User Data'
chrome_profile = '--profile-directory=Profile 1'
##GLOBE Google Drive###
januaryid = '16NagReBAcnNI0pQZP6QVQBzhSJtXMDoV'
februaryid = '1hajPgONDdwHhzOHE-x9LAIxtReRsfH77'
marchid = '1YsXs92EE3-4gKhkx8SSiR1hXrxraaFuo'
aprilid = '1727PKtqIFr9RkBMMD-D1pyV1HMBR8TnT'
mayid = '1PQIlIAWnoCNSpnjP0h3h9Tr6DnbD8Wak'
juneid = '1BXYtOOXi2UYvFk9CDtvXCZiI0vvP30D_'
julyid = '1gs4AR2ToEXD8v0GwxplX6ezlpqvxIS0w'
augustid = '14lBggDdPPDImqpWIQWqzJ5LUOh8kioct'
septemberid = '1PisqaYqnb_9CgkhbN8e3mhi0DM0yrG_k'
octoberid = '1McJGs3dPX1--4nT1HbuKDu9ol9Vk6pnq'
novemberid = '1TzzCF_iQPH0dTqj49oiDOYTEh4nRycux'
decemberid = '133i0_5M3RxBBczb0ehcS7aRDi_SV6FsW'
##INNOVE Google Drive###
januaryidinnove = '1n0wXk4eicfYEXiOU_BAbkBvtQJ62HH5B'
februaryidinnove = '13cyfZKyBl0DM8xYz5Bq08KNunCl1rm-q'
marchidinnove = '1mIQKjZ0HcduHW6M2gKFrnQdkTPh7IWIx'
aprilidinnove = '1tHj1if18DUkHMb-R-wMX-DJ8AHpv6q94'
mayidinnove = '1CSsRbXAu4sygeg8wpR2nGYA6XULhxitJ'
juneidinnove = '14rAkmDMPU9ge346sjHgAXXHSjb0YDg_R'
julyidinnove = '1dWOETprh1MNfVQHjbN7dj8kHlF132I6J'
augustidinnove = '1NpnaS0RhgPnOQkvdQl82BEL24WjiGSur'
septemberidinnove = '1WWSurH5la4BQtZ17UbkwCO5OsuIjVfKS'
octoberidinnove = '14hcoUg_59lfKpLyBwsAIelzwyXN9KsYT'
novemberidinnove = '1LKvSbqBtee144luas2lI22DjH6X-gHy4'
decemberidinnove = '1USptGsjcj_SSkRNV5IHO8_OvzSaTnqik'
##BAYAN Google Drive###
januaryidbayan = '1JpyJ6v7lPQQG3cd7pur8fHcp8HDXlXv5'
februaryidbayan = '1E_MkoVqEr3jFjGVdApUrHvphcaAynxen'
marchidbayan = '1ONgPaE2x10VvGa9FwwF0ecFZflg9LxG1'
aprilidbayan = '1XNSCYJ6UpyxGNPKVJD2cv55XqIDphNiR'
mayidbayan = '1l8a-B_W13956TSV2dRoiQzw1uc3r6-K2'
juneidbayan = '1Es3tyWWiblWh-ef-CNyf9PSSRojuElWA'
julyidbayan = '1-UHyX_BpoDVL4mw23vC5OZfXWWUzA7cJ'
augustidbayan = '19kj9qSM-TX31zr3udcprvNsp7fNmr5dp'
septemberidbayan = '1E3t0pxdqnHsKtgPuGRQUFzOOuh2O8OLj'
octoberidbayan = '1ks4AAO9QovAgX-3jdXtL-wZQqrCNrUjR'
novemberidbayan = '1YM6Kk3qY8m0NQmw238CW0CMbAYO6Wz1E'
decemberidbayan = '1w95NzN1ChcEN-WREyzUzuOmW4pSIIB8V'
########################################################################
month = []
year = []
date = datetime.now()

# Variables (DO NOT EDIT!!)
error = 0
start = 0
custom = 0
loop = 0
display = 0
globe = 0
innove = 0
bayan = 0

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
    qwindow.configure(bg='#1C3D80')

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
        def globe():
            global globe
            globe = 1
            mwindow.destroy()
            window.destroy()
            iwindow.destroy()
        def innove():
            global innove
            innove = 1
            mwindow.destroy()
            window.destroy()
            iwindow.destroy()
        def bayan():
            global bayan
            bayan = 1
            mwindow.destroy()
            window.destroy()
            iwindow.destroy()

        iwindow = Tk()
        btnglobe = Button(iwindow, text="Globe", command=globe)
        btnglobe.place(x=50, y=100)
        btninnove = Button(iwindow, text="Innove", command=innove)
        btninnove.place(x=120, y=100)
        btninnove = Button(iwindow, text="Bayan", command=bayan)
        btninnove.place(x=190, y=100)
        lbl2 = Label(iwindow, text="Select Payee:", fg='white', font=("Helvetica", 14), bg='#1C3D80')
        lbl2.place(x=20, y=50)
        iwindow.title('eCWT')
        iwindow.geometry("300x200")
        iwindow.configure(bg='#1C3D80')
        iwindow.eval('tk::PlaceWindow . center')

    ok = Button(mwindow, text="OK", command=selected_item)
    ok.place(x=200, y=220)
    mwindow.title('Month')
    mwindow.eval('tk::PlaceWindow . center')
    mwindow.geometry("250x300")
    mwindow.configure(bg='#1C3D80')

def display_on():
    global root
    print("Always On")
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)
    root.iconify()

def display_reset():
    ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
    sys.exit(0)

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
print(custom)
print("START: " + str(datetime.now()))
# url = 'https://globetelecominc-tst.outsystemsenterprise.com/eCWT_Portal/Login_SSO.aspx'
url = 'https://ecwt.globe.com.ph/portal/Login_SSO.aspx' #PROD

options = Options()
options.add_experimental_option("detach", True)
options.add_argument(chrome_path)
options.add_argument(chrome_profile)
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
driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt207_block_wtMenu_wt109_RichWidgets_wt55_block_wtMenuItem_wt132"]').click()  # PROD
time.sleep(1)
# driver.find_element(By.XPATH, '//*[@href="PDFUploading.aspx?(Not.Licensed.For.Production)="]').click()
driver.find_element(By.XPATH,'//*[@id="LisbonTheme_wt207_block_wtMenu_wt109_RichWidgets_wt55_block_wtMenuSubItems_wt87"]').click()  # PROD
###TO EDIT###
driver.find_element(By.XPATH, '//*[@data-identifier="zmcmorales@globe.com.ph"]').click()
####
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
    time.sleep(5)
    if globe == 1:
        driver.find_element(By.XPATH, '//*[@value="Globe"]').click()
        time.sleep(3)
        driver.implicitly_wait(60)
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').clear()
        if mnum == 1:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(januaryid)
        elif mnum == 2:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(februaryid)
        elif mnum == 3:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(marchid)
        elif mnum == 4:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(aprilid)
        elif mnum == 5:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(mayid)
        elif mnum == 6:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(juneid)
        elif mnum == 7:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(julyid)
        elif mnum == 8:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(augustid)
        elif mnum == 9:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(septemberid)
        elif mnum == 10:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(octoberid)
        elif mnum == 11:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(novemberid)
        elif mnum == 12:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(decemberid)
    if innove == 1:
        driver.find_element(By.XPATH, '//*[@value="Innove"]').click()
        time.sleep(3)
        driver.implicitly_wait(60)
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').clear()
        if mnum == 1:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(januaryidinnove)
        elif mnum == 2:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(februaryidinnove)
        elif mnum == 3:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(marchidinnove)
        elif mnum == 4:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(aprilidinnove)
        elif mnum == 5:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(mayidinnove)
        elif mnum == 6:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(juneidinnove)
        elif mnum == 7:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(julyidinnove)
        elif mnum == 8:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(augustidinnove)
        elif mnum == 9:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(septemberidinnove)
        elif mnum == 10:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(octoberidinnove)
        elif mnum == 11:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(novemberidinnove)
        elif mnum == 12:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(decemberidinnove)
    if bayan == 1:
        driver.find_element(By.XPATH, '//*[@value="Bayan"]').click()
        time.sleep(1)
        driver.implicitly_wait(60)
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').clear()
        if mnum == 1:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(januaryidbayan)
        elif mnum == 2:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(februaryidbayan)
        elif mnum == 3:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(marchidbayan)
        elif mnum == 4:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(aprilidbayan)
        elif mnum == 5:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(mayidbayan)
        elif mnum == 6:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(juneidbayan)
        elif mnum == 7:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(julyidbayan)
        elif mnum == 8:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(augustidbayan)
        elif mnum == 9:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(septemberidbayan)
        elif mnum == 10:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(octoberidbayan)
        elif mnum == 11:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(novemberidbayan)
        elif mnum == 12:
            driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt46"]').send_keys(decemberidbayan)
    driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt65"]').click()
    time.sleep(5)
    if start == 0:
        driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt68"]').click()
        time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMainContent_wt4"]').click()
    # try:
    #     element = WebDriverWait(driver, 3600).until(
    #         EC.visibility_of_element_located((By.XPATH, '//*[text()="No ecwtpdf uploads to show..."]'))
    #     )
    # finally:
    #     loop=loop-1
    #     pass
    if display == 0:
        print("Display on")
        root = Tk()
        root.geometry("200x60")
        root.title("Display App")
        root.configure(bg='#1C3D80')
        root.eval('tk::PlaceWindow . center')
        display_on()
    else:
        display_reset()
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
display = 1
print("END: " + str(datetime.now()))
driver.find_element(By.XPATH, '//*[@id="LisbonTheme_wt94_block_wtMenu_wt76_wt9_wtLogoutLink"]').click()
driver.quit()
easygui.msgbox("Upload Done.")
