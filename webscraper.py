from selenium import webdriver
from selenium.common import exceptions as seleniumException
import openpyxl
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep

## IMPORTANT! BREAKPOINT AT LINE 173 TO PAUSE FOR LOGIN ##

# Selenium CSS classes
cardCssClass = "index__ExhibitorCard--uaYJ5"
cardTitleClass = "index__title--PQWpm"
companyInfoClass = "index__name--_7RES"
companyContactClass = "index__name--_7RES"

# XLSX / Data Storage
# Reference File stores the company numbers independently of the cantonfair.xlsx progress
# cantonfair.xlsx is the main data storage, content & columns are divided per worksheetColumns
worksheetColumns = {
    "CN": {
        "name": "A",
        "province": "B",
        "address": "C",
        "founded": "D",
        "employees": "E",
        "type": "F",
        "clients": "G",
        "products": "H",
        "website": "I",
        "contactPerson": "J",
        "officePhone": "K",
        "privatePhone": "L",
        "email": "M",
    },
    "EN": {
        "name": "N",
        "province": "O",
        "address": "P",
        "founded": "Q",
        "employees": "R",
        "type": "S",
        "clients": "T",
        "products": "U",
        "website": "V",
        "contactPerson": "W",
        "officePhone": "X",
        "privatePhone": "Y",
        "email": "Z",
    },
}
filename = "cantonfair"
extFilename = f"{filename}.xlsx"
referenceFilename = "ref"
ReferenceExtFilename = f"{referenceFilename}.xlsx"
sheetName = "Sheet1"
refWB = openpyxl.load_workbook(filename=ReferenceExtFilename)
refSheet = refWB[sheetName]
wb = openpyxl.load_workbook(filename=extFilename)
sheet = wb[sheetName]


def setPauseForLogin(val: bool):
    pauseForLogin = val


def getCantonFairURL(page: int, size=60):
    return f"https://www.cantonfair.org.cn/zh-CN/detailed?category=461147067079729152%2C461151997756706816%2C461147173489221632%2C461147340519006208%2C461147775145353216%2C461147662834475008%2C461148599963619328%2C461147404528259072%2C461148890532417536%2C461147860176490496%2C461149133265178624%2C461147666248638464%2C461147574737326080%2C461147830828945408%2C461147080727994368%2C461147862630162432&scategory=461147067079729152%2C461151997756706816%2C461147173489221632%2C461147340519006208%2C461147775145353216%2C461147662834475008%2C461148599963619328%2C461147404528259072%2C461148890532417536%2C461147860176490496%2C461149133265178624%2C461147666248638464%2C461147574737326080%2C461147830828945408%2C461147080727994368%2C461147862630162432&type=1&keyword=&page={page}&size={size}&offline=N&tab=exhibitor&sort=relate%20desc&filter=186dfe854e9-17244"


def getSearchExhibitorURL(companyName: str, lang: str):
    match lang:
        case "CN":
            return f"https://www.cantonfair.org.cn/zh-CN/detailed?category=&scategory=&type=2&keyword={companyName}&page=1&size=60&offline=N&tab=exhibitor&sort=relate%20desc&filter=186e4ef248c-d93b"
        case "EN":
            return f"https://www.cantonfair.org.cn/en-US/detailed?category=&scategory=&type=2&keyword={companyName}&page=1&size=60&offline=N&tab=exhibitor&sort=relate%20desc&filter=186e4ef248c-d93b"


def getSearchProductURL(searchParams: str):
    return f"https://www.cantonfair.org.cn/zh-CN/detailed?category=&scategory=&type=1&keyword={searchParams}page=1&size=60&offline=N&tab=product&sort=relate%20desc&filter=186e4ef248c-d93b"


def goToTab(num: int, driver):
    driver.switch_to.window(driver.window_handles[num - 1])


def saveMainFile():
    wb.save(filename=extFilename)


def saveRefFile():
    refWB.save(filename=ReferenceExtFilename)


def getReference():
    # Get references and store them in a seperate file
    driver = webdriver.Firefox()
    maxPages = 34

    for companyIndex in range(1, maxPages + 1):
        print(f"Current index: {companyIndex}")  # Logging

        # Open URL
        driver.get(getCantonFairURL(page=companyIndex))
        sleep(15)

        titleDivs = driver.find_elements(By.CLASS_NAME, cardTitleClass)

        # Write Company Names to A Column in XLSX
        for title in titleDivs:
            for row in range(2, 10_000):
                if refSheet[f"A{row}"].value != None:
                    continue
                refSheet[f"A{row}"].value = title.text

                print(f"Row #{row}: {title.text}")  # Logging
                saveRefFile()
                break
        print()
    # except (AttributeError, seleniumException.NoSuchElementException):
    #     continue


def getAllCompaniesInfo():
    # Get all companies' information and store is in an XLSX file

    # Init Selenium Browser
    driver = webdriver.Firefox()

    # Search for every company name
    for row in range(328, 10_000):
        goToTab(num=1, driver=driver)
        companyName = refSheet[f"A{row}"].value
        print(f"Getting info for row #{row}")
        if companyName == None:
            break
        getCompanyInfo(
            lang="EN",
            row=row,
            companyName=companyName,
            driver=driver,
        )
        getCompanyInfo(
            lang="CN",
            row=row,
            companyName=companyName,
            driver=driver,
        )

    # Quit all browser tabs/windows
    driver.quit()


def getCompanyInfo(lang: str, row: int, companyName: str, driver):
    # Get a single company's information in the language given

    goToTab(num=1, driver=driver)

    try:
        driver.get(getSearchExhibitorURL(companyName, lang=lang))
    except (seleniumException.WebDriverException):
        print("WebDriverException. Trying again in 10s")
        sleep(10)
        driver.get(getSearchExhibitorURL(companyName, lang=lang))

    sleep(3)

    # Breakpoint here to manually Sign In
    # after login wait for popup, click the X
    # after login disable breakpoint
    try:
        card = driver.find_elements(By.CLASS_NAME, "index__title--PQWpm")[0]
    except (IndexError):
        try:
            print("Couldn't find card title, trying something else...")
            card = driver.find_elements(By.CLASS_NAME, "index__title--PQWpm")[0]
        except:
            print("Could not find card title, skipping")
            return
    card.click()
    sleep(5)
    goToTab(num=2, driver=driver)
    companyBaseURL = driver.current_url

    # Contact Page
    driver.get(f"{companyBaseURL}contact")
    sleep(3)
    try:
        gatedSection = driver.find_element(By.CLASS_NAME, "index__gate--EcyhK")
        btn = gatedSection.find_element(By.TAG_NAME, "button")
        btn.click()
    except (
        seleniumException.NoSuchElementException,
        seleniumException.ElementClickInterceptedException,
    ):
        try:
            print("Couldn't find gated section, trying something else...")
            sleep(5)
            gatedSection = driver.find_element(
                By.XPATH, "/html/body/div[2]/div/div[5]/div/div[2]"
            )
            btn = gatedSection.find_element(By.TAG_NAME, "button")
            btn.click()
        except:
            print("Could not find gated section button")

    sleep(1)
    contactElements = driver.find_elements(By.CLASS_NAME, "index__item--vuNk7")

    # Loop through all information fields
    for element in contactElements:
        categoryClass = "index__name--KiZnD"
        contentClass = "index__content--HCLQC"

        category = element.find_element(By.CLASS_NAME, categoryClass).text
        content = element.find_element(By.CLASS_NAME, contentClass).text

        match category:
            case "企业名称":
                sheet[f"{worksheetColumns['CN']['name']}{row}"].value = content
            case "企业网站":
                sheet[f"{worksheetColumns['CN']['website']}{row}"].value = content
            case "国家/地区":
                sheet[f"{worksheetColumns['CN']['province']}{row}"].value = content
            case "地址":
                sheet[f"{worksheetColumns['CN']['address']}{row}"].value = content
            case "业务联系人":
                sheet[f"{worksheetColumns['CN']['contactPerson']}{row}"].value = content
            case "办公电话":
                sheet[f"{worksheetColumns['CN']['officePhone']}{row}"].value = content
            case "手机":
                sheet[f"{worksheetColumns['CN']['privatePhone']}{row}"].value = content
            case "邮箱":
                sheet[f"{worksheetColumns['CN']['email']}{row}"].value = content
            case "Company Name":
                sheet[f"{worksheetColumns['EN']['name']}{row}"].value = content
            case "Company website":
                sheet[f"{worksheetColumns['EN']['website']}{row}"].value = content
            case "Country/Region":
                sheet[f"{worksheetColumns['EN']['province']}{row}"].value = content
            case "Address":
                sheet[f"{worksheetColumns['EN']['address']}{row}"].value = content
            case "Contact Person":
                sheet[f"{worksheetColumns['EN']['contactPerson']}{row}"].value = content
            case "Telephone":
                sheet[f"{worksheetColumns['EN']['officePhone']}{row}"].value = content
            case "Mobile Phone":
                sheet[f"{worksheetColumns['EN']['privatePhone']}{row}"].value = content
            case "Email":
                sheet[f"{worksheetColumns['EN']['email']}{row}"].value = content
            case other:
                continue
        saveMainFile()

    # Introduction Page
    driver.get(f"{companyBaseURL}introduction")
    sleep(2)
    contactElements = driver.find_elements(By.CLASS_NAME, "index__item--vuNk7")

    # Loop through all information fields
    for element in contactElements:
        categoryClass = "index__name--KiZnD"
        contentClass = "index__content--HCLQC"

        category = element.find_element(By.CLASS_NAME, categoryClass).text
        content = element.find_element(By.CLASS_NAME, contentClass).text

        match category:
            case "企业类型":
                sheet[f"{worksheetColumns['CN']['type']}{row}"].value = content
            case "成立日期":
                sheet[f"{worksheetColumns['CN']['founded']}{row}"].value = content
            case "企业规模":
                sheet[f"{worksheetColumns['CN']['employees']}{row}"].value = content
            case "主要目标客户":
                sheet[f"{worksheetColumns['CN']['clients']}{row}"].value = content
            case "主营展品":
                sheet[f"{worksheetColumns['CN']['products']}{row}"].value = content
            case "Company type":
                sheet[f"{worksheetColumns['EN']['type']}{row}"].value = content
            case "Register Date":
                sheet[f"{worksheetColumns['EN']['founded']}{row}"].value = content
            case "Enterprise Scale":
                sheet[f"{worksheetColumns['EN']['employees']}{row}"].value = content
            case "Main Target Customers":
                sheet[f"{worksheetColumns['EN']['clients']}{row}"].value = content
            case "Main Products":
                sheet[f"{worksheetColumns['EN']['products']}{row}"].value = content
            case other:
                continue

        saveMainFile()
    driver.close()


if __name__ == "__main__":
    # getReference()
    getAllCompaniesInfo()
