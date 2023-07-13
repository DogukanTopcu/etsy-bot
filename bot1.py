from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side
import urllib.request
import datetime
import os
from selenium.webdriver.common.keys import Keys


class Order:
    def __init__(self, name, street, apartment, city, country, image, final, date, isStarted):
        self.name = name
        self.street = street
        self.apartment = apartment
        self.city = city
        self.country = country
        self.image = image
        self.final = final
        self.date = date
        self.isStarred = isStarted

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
sheet.column_dimensions['A'].width = 11
sheet.column_dimensions['B'].width = 11
sheet.column_dimensions['C'].width = 11
sheet.column_dimensions['D'].width = 11
sheet.column_dimensions['E'].width = 40
sheet.column_dimensions['F'].width = 30
workbook.save(f"./{folderPath}/{stringToday} excel/{stringToday}.xlsx")
cell_row = 1
alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


driver.get("https://www.etsy.com/")
driver.maximize_window()
time.sleep(0.1)

signIn = driver.find_element(
    By.XPATH, "/html/body/div[2]/header/div[4]/nav/ul/li[1]/button")
signIn.click()
time.sleep(0.5)

emailInput = driver.find_element(By.XPATH, "//*[@id='join_neu_email_field']")
emailInput.send_keys("estveryin15@gmail.com")
time.sleep(0.3)
passwordInput = driver.find_element(
    By.XPATH, "//*[@id='join_neu_password_field']")
passwordInput.send_keys("2022Th3005")
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
            order = Order("", "", "", "", "", [], [], str(today.month) + "/" + str(today.day) + "/" + str(today.year), False)
            workbook = openpyxl.load_workbook(f"./{folderPath}/{stringToday} excel/{stringToday}.xlsx")
            sheet = workbook.active
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
            
            try:
                imgSrc = driver.find_element(By.XPATH, "/html/body/main/div[1]/div[2]/div/div/div[1]/div[1]/div/div/div/div/div[1]/ul/li[1]/img").get_attribute("src")
            except:
                imgSrc = driver.find_element(By.XPATH, "//*[@id='content']/div/div[1]/div/div/div[2]/div[1]/div/img")
            
            try:
                fileName = driver.find_element(By.XPATH, "/html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[4]/h1").text
            except:
                fileName = driver.find_element(By.XPATH, "/html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[3]/h1").text

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

        print(len(order.image))
        l = 0
        for img in order.image:
            imgUrl = f"./{folderPath}/{stringToday} excel/fotograflar/{order.name}.jpg"
            c = 1
            while True:
                if not os.path.exists(imgUrl):
                    urllib.request.urlretrieve(img, imgUrl)
                    break
                else:
                    imgUrl = f"./{folderPath}/{stringToday} excel/fotograflar/{order.name}({str(c)}).jpg"
                    c += 1
            
            image = Image(imgUrl)
            # img_cell = sheet.cell(row=cell_row+1, column=l)
            sheet.add_image(image, f"{alphabet[l]}{str(cell_row + 1)}")

            sheet[f"{alphabet[l]}{str(cell_row + 5)}"] = order.final[l]
            sheet[f"{alphabet[l]}{str(cell_row + 5)}"].font = Font(size=8, name="Arial Tur")
            l += 1
        
        sheet[f"A{cell_row}"] = order.date
        sheet[f"C{cell_row}"] = order.date
        sheet[f"A{cell_row}"].font = Font(size=8, name="Arial Tur")
        sheet[f"C{cell_row}"].font = Font(size=8, name="Arial Tur")

        sheet[f"F{cell_row + 1}"] = "Tugce SONUC"
        sheet[f"F{cell_row + 1}"].font = Font(size=11, name="Arial Tur")

        sheet[f"F{cell_row + 2}"] = "Gaziosmanpasa Bulv."
        sheet[f"F{cell_row + 2}"].font = Font(size=11, name="Arial Tur")

        sheet[f"F{cell_row + 3}"] = "No:64 D:601 K:6 Yüncü İşhanı"
        sheet[f"F{cell_row + 3}"].font = Font(size=10, name="Arial Tur")

        sheet[f"F{cell_row + 4}"] = "Cankaya - İzmir"
        sheet[f"F{cell_row + 4}"].font = Font(size=11, name="Arial Tur")



        sheet[f"E{cell_row + 1}"] = "Ship To"
        sheet[f"E{cell_row + 1}"].font = Font(size=11, name="Arial Tur")

        sheet[f"E{cell_row + 2}"] = order.name
        sheet[f"E{cell_row + 2}"].font = Font(size=9, name="Arial Tur", bold=True)

        sheet[f"E{cell_row + 3}"] = order.street
        sheet[f"E{cell_row + 3}"].font = Font(size=9, name="Tahoma")

        sheet[f"E{cell_row + 4}"] = order.apartment
        sheet[f"E{cell_row + 4}"].font = Font(size=9, name="Arial Tur")

        sheet[f"E{cell_row + 5}"] = order.city
        sheet[f"E{cell_row + 5}"].font = Font(size=9, name="Arial Tur")
        
        border = Border(
            left=Side(border_style='thin', color='000000'),
            right=Side(border_style='thin', color='000000'),
            top=Side(border_style='thin', color='000000'),
            bottom=Side(border_style='thin', color='000000')
        )

        # sheet[f"A{cell_row}:F{cell_row}"].border = border

        if len(order.country) != 0:
            sheet[f"E{cell_row + 6}"] = order.country
            sheet[f"E{cell_row + 6}"].font = Font(size=9, name="Arial Tur")

            sheet[f"F{cell_row + 6}"] = "35240   Turkey"
            sheet[f"F{cell_row + 6}"].font = Font(size=11, name="Arial Tur")

            # sheet[f"A{cell_row}:D{cell_row + 6}"].border = border
            # sheet[f"E{cell_row}:E{cell_row + 6}"].border = border
            # sheet[f"F{cell_row}:F{cell_row + 6}"].border = border

            cell_row += 7
        else:
            sheet[f"F{cell_row + 5}"] = "35240   Turkey"
            sheet[f"F{cell_row + 5}"].font = Font(size=11, name="Arial Tur")

            # sheet[f"A{cell_row}:D{cell_row + 5}"].border = border
            # sheet[f"E{cell_row}:E{cell_row + 5}"].border = border
            # sheet[f"F{cell_row}:F{cell_row + 5}"].border = border

            cell_row += 6


        
        i += 1

        workbook.save(f"./{folderPath}/{stringToday} excel/{stringToday}.xlsx")
    k += 1

time.sleep(1)

driver.quit()

# /html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[4]/h1
# /html/body/main/div[1]/div[2]/div/div/div[1]/div[2]/div/div[4]/h1

# //*[@id="content"]/div/div[1]/div/div/div[2]/div[1]/div/img