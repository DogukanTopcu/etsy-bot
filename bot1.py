from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl
from openpyxl.drawing.image import Image
import urllib.request
import datetime
import os
from selenium.webdriver.common.keys import Keys


class Order:
    name = ""
    street = ""
    apartment = ""
    city = ""
    country = ""
    image = []
    final = []
    date = ""
    isStarred = False


# New Director Path

today = datetime.datetime.now()
stringToday = str(today.month) + "-" + str(today.day) + "-" + str(today.year)
print(stringToday)

if not os.path.exists(stringToday):
    os.mkdir(stringToday)
    folderPath = stringToday
    os.mkdir(f"./{folderPath}/{stringToday} fotograflar")
    os.mkdir(f"./{folderPath}/{stringToday} excel")
    os.mkdir(f"./{folderPath}/{stringToday} excel/fotograflar")
else:
    z = 1
    while True:
        if not os.path.exists(stringToday + " (" + str(z) + ")"):
            os.mkdir(stringToday + " (" + str(z) + ")")
            folderPath = stringToday + " (" + str(z) + ")"
            os.mkdir(f"./{folderPath}/{stringToday} fotograflar")
            os.mkdir(f"./{folderPath}/{stringToday} excel")
            os.mkdir(f"./{folderPath}/{stringToday} excel/fotograflar")
            
            break
        z += 1
# ************************************************


PATH = "C:\Program Files (x86)\msedgedriver.exe"
driver = webdriver.Edge()

workbook = openpyxl.Workbook()
sheet = workbook.active
workbook.save(f"./{folderPath}/{stringToday} excel/sample.xlsx")


driver.get("https://www.etsy.com/")
driver.maximize_window()
time.sleep(0.1)

signIn = driver.find_element(
    By.XPATH, "/html/body/div[2]/header/div[4]/nav/ul/li[1]/button")
signIn.click()
time.sleep(0.5)

emailInput = driver.find_element(By.XPATH, "//*[@id='join_neu_email_field']")
emailInput.send_keys("*******************")
time.sleep(0.3)
passwordInput = driver.find_element(
    By.XPATH, "//*[@id='join_neu_password_field']")
passwordInput.send_keys("*****************")
time.sleep(0.3)

signInButton = driver.find_element(
    By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[2]/form/div[1]/div/div[7]/div/button")
signInButton.click()
time.sleep(25)


shopManager = driver.find_element(
    By.XPATH, "//*[@id='gnav-header-inner']/div[4]/nav/ul/li[3]")
shopManager.click()
time.sleep(3)


od = driver.find_element(
    By.XPATH, "//*[@id='root']/div/div[1]/div[3]/div/div[1]/div[2]/ul/li[5]")
od.click()
time.sleep(10)


k = 2
while True:
    dispatchXP = "/html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[" + str(
        k) + "]"

    try:
        dispatch = driver.find_element(By.XPATH, dispatchXP)
    except:
        break

    i = 1
    while True:
        xpath = dispatchXP + "/div[2]/div[" + str(i) + "]"

        try:
            data = driver.find_element(By.XPATH, xpath)
            order = Order()
            order.date = stringToday
        except:
            break


        # TASK 2*************************************************************************************************
        data.click()
        time.sleep(2)
        a = 2
        original_window = driver.current_window_handle
        while True:
            linkPath = "/html/body/div[3]/div/div[1]/div/main/div/div[3]/div[2]/div/div/div[9]/div/div/div[2]/table/tbody/tr[" + str(a)  + "]/td[1]/div[2]/div[2]/a"

            try:
                link = driver.find_element(By.XPATH, linkPath)
            except:
                try:
                    linkPath = "/html/body/div[3]/div/div[1]/div/main/div/div[3]/div[2]/div/div/div[10]/div/div/div[2]/table/tbody/tr[" + str(a)  + "]/td[1]/div[2]/div[2]/a"
                    link = driver.find_element(By.XPATH, linkPath)
                except:
                    break
            link.click()
            time.sleep(0.15)
            driver.switch_to.window(driver.window_handles[1])

            time.sleep(1)
            imgSrc = driver.find_element(By.XPATH, "/html/body/main/div[1]/div[2]/div/div/div[1]/div[1]/div/div/div/div/div[1]/ul/li[1]/img").get_attribute("src")
            fileName = driver.find_element(By.XPATH, "/html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[4]/h1").text

            try:
                urllib.request.urlretrieve(imgSrc, f"./{folderPath}/{stringToday} fotograflar/{fileName}.jpg")
            except:
                print("")
                print(fileName, "indirilemedi")
                print("")

            time.sleep(0.5)

            driver.close()
            driver.switch_to.window(original_window)
            a += 1
        
        # driver.find_element(By.XPATH, "/html/body/div[3]/div/div[1]/div/main/div/div[3]/div[2]/button").click()
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.2)
        # *************************************************************************************************

        # time.sleep(1)
        delivertoBtn = driver.find_element(
            By.XPATH, xpath + "/div/div[2]/div[1]/div[2]/div[4]/div/div[1]/div/button")
        delivertoBtn.click()

        obj = driver.find_element(
            By.XPATH, xpath + "/div/div[2]/div/div[2]/div[4]/div/div[2]/div/div/p[1]")
        id = obj.text.split("\n")
        order.name = id[0]
        order.street = id[1]
        order.apartment = id[2]
        order.city = id[3]
        try:
            order.country = id[4]
        except:
            order.country = ""
        
        print("*******************************")
        print(id[0])
        print(id[1])
        print(id[2])
        print(id[3])
        print("*******************************")

        j = 1
        while True:
            isExistPath = xpath + "/div/div[2]/div[2]"
            try:
                commentLine = driver.find_element(By.XPATH, isExistPath)
                productPath = xpath + \
                    "/div/div[2]/div[1]/div[1]/div[3]/div/div[" + str(j) + "]"
            except:
                productPath = xpath + \
                    "/div/div[2]/div/div[1]/div[2]/div/div[" + str(j) + "]"

            try:
                pro = driver.find_element(By.XPATH, productPath)
            except:
                break

            try:
                proNamePath = productPath + \
                    "/div/div[2]/ul/ul/li[3]/div[1]/span[2]"
                proName = driver.find_element(By.XPATH, proNamePath)
            except:
                proNamePath = productPath + \
                    "/div/div[2]/ul/ul/li[2]/div[1]/span[2]"
                proName = driver.find_element(By.XPATH, proNamePath)

            order.final.append(proName.text)

            try:
                imgPath = productPath + \
                    "/div/div[2]/ul/ul/li[3]/div[1]/span[2]"
                proName = driver.find_element(By.XPATH, proNamePath)
            except:
                imgPath = productPath + \
                    "/div/div[2]/ul/ul/li[2]/div[1]/span[2]"
                proName = driver.find_element(By.XPATH, proNamePath)

            img = driver.find_element(
                By.XPATH, productPath + "/div/div[1]/a/div/div[1]/img")
            src = img.get_attribute("src")
            

            order.image.append(src)

            print()

            j += 1

        i += 1

    k += 1

time.sleep(1)

driver.quit()


# /html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[1]/div[1]/div/span[2]/span/span/span/svg
# /html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[3]/div/div[2]/div/div[1]/div[1]/div[1]/div[1]/div/span[2]/span/span/svg
# /html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[7]/div/div[2]/div/div[1]/div[1]/div[1]/div[1]/div/span[2]/span/span/svg
# /html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[4]/div[2]/div[4]/div/div[2]/div/div[1]/div[1]/div[1]/div[1]/div/span[2]/span/span/svg

# /html/body/div[3]/div/div[1]/div/main/div/div[2]/div/div/div[1]/div[3]/div[2]/div[2]/div[3]/div/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div/span[2]/span/span[1]/svg



# /html/body/main/div[1]/div[2]/div/div/div[1]/div[1]/div/div/div/div/div[1]/ul/li[1]/img
# /html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[4]/h1
