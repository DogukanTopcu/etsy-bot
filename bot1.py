from selenium import webdriver
from selenium.webdriver.common.by import By
import time

PATH = "C:\Program Files (x86)\msedgedriver.exe"
driver = webdriver.Edge()

driver.get("https://www.etsy.com/")
time.sleep(0.1)

signIn = driver.find_element(
    By.XPATH, "/html/body/div[2]/header/div[4]/nav/ul/li[1]/button")
signIn.click()
time.sleep(0.5)

emailInput = driver.find_element(By.XPATH, "//*[@id='join_neu_email_field']")
emailInput.send_keys("*********************")
time.sleep(0.3)
passwordInput = driver.find_element(
    By.XPATH, "//*[@id='join_neu_password_field']")
passwordInput.send_keys("******************")
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
        except:
            break

        titlePath = xpath + \
            "/div/div[2]/div/div[1]/div[1]/div[1]/div[1]/div/span[1]/span/span/div/span"
        name = data.find_element(By.XPATH, titlePath)
        print(name.text)

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

            proNamePath = productPath + "/div/div[2]/div[1]/div/div/span"
            proName = driver.find_element(By.XPATH, proNamePath)
            print(proName.text)
            print()
            j += 1

        i += 1

    k += 1

time.sleep(1)
driver.quit()
