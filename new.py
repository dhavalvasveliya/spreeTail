from numpy import append
import pandas as pd
import time
import datetime
from datetime import date,timedelta
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
import requests
# import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

website_url = 'http://partner.spreetail.com/'
username = 'automations@globecommerce.co'
pw = '7k*hXbz565H$Fp'


kingcamp_sales_url = 'https://partner.spreetail.com/kingcamp-outdoor-products-co-ltd/sales'
xindao_sales_url = 'https://partner.spreetail.com/xindao-us-ltd/sales'

xindao_start_date = '2021-05-13'
kingcamp_start_date = '2021-08-12'
# Xindao starts at 5/13/21
# KingCamp starts at 8/12/21


def start_driver(headless_or_not):
    user_agent_pc = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36'
    # if headless_or_not=='headless':
    options = Options()
    # options.headless=True
    options.add_argument('--headless')
    options.add_argument('--user-agent={user_agent}'.format(user_agent=user_agent_pc))
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    # driver = uc.Chrome(executable_path="/usr/bin/chromedriver",options=options)
    # driver = uc.Chrome(options=options)
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    return driver
    # else:
    #     options = uc.ChromeOptions()
    #     options.add_argument('--user-agent={user_agent}'.format(user_agent=user_agent_pc))
    #     driver = uc.Chrome(options=options)
    #     return driver

def login(driver,url,username,pw):
    driver.get(url)
    while True:
        try:
            print('Trying to login')
            driver.find_element(By.ID, 'okta-signin-username').clear()
            driver.find_element(By.ID, 'okta-signin-username').send_keys(username)
            driver.find_element(By.ID, 'okta-signin-password').clear()
            driver.find_element(By.ID, 'okta-signin-password').send_keys(pw)
            driver.find_element(By.ID, 'okta-signin-submit').click()
            time.sleep(3)
            while True:
                try:
                    driver.find_element(By.CSS_SELECTOR, '[data-testid=doneButton]').click()
                    break
                except:
                    pass
            print('Logged in')
            break
        except Exception as e:
            print(e)
            time.sleep(2)


def fix_negative_values(x):
    if x<0:
        return int(x/86400000)*-1
    else:
        return x


def fix_percentage(x):
    try:
        return int(x.replace('%','').replace(' ',''))
    except:
        return x



def select_month(driver,monthYear,dateinput): #e.g monthYear=December 2021 #dateinput=either start or end
    while True:
        target = monthYear
        if dateinput=='start':
            check = driver.find_element(By.CLASS_NAME,'rdtSwitch').text
        else:
            check = driver.find_elements(By.CLASS_NAME,'rdtSwitch')[1].text
        if check==target:
            break
        else:
            if dateinput=='start':
                driver.find_element(By.CLASS_NAME,'rdtNext').click()
            else:
                driver.find_elements(By.CLASS_NAME,'rdtNext')[1].click()

def select_day(driver,day,dateinput):
    if dateinput=='start':
        if int(day) < 15:
            [td for td in driver.find_element(By.CLASS_NAME,'rdtDays').find_elements(By.TAG_NAME,'td') if td.text==day][0].click()
        else:
            [td for td in driver.find_element(By.CLASS_NAME,'rdtDays').find_elements(By.TAG_NAME,'td') if td.text==day][-1].click()
    else:
        if int(day) < 15:
            [td for td in driver.find_elements(By.CLASS_NAME,'rdtDays')[1].find_elements(By.TAG_NAME,'td') if td.text==day][0].click()
        else:
            [td for td in driver.find_elements(By.CLASS_NAME,'rdtDays')[1].find_elements(By.TAG_NAME,'td') if td.text==day][-1].click()
        



def change_start_date(driver,date):#format15 December 2021
    day = date.split(' ')[0]
    if day.startswith('0'):
        day=day[1:]
    monthYear = ' '.join(date.split(' ')[1:])
    while True:
        try:
            driver.find_element(By.ID, 'startDateRange').click()
            break
        except:
            time.sleep(2)
    select_month(driver,monthYear,'start')
    select_day(driver,day,'start')


def change_end_date(driver,date):
    day = date.split(' ')[0]
    if day.startswith('0'):
        day=day[1:]
    monthYear = ' '.join(date.split(' ')[1:])
    driver.find_element(By.ID, 'endDateRange').click()
    select_month(driver,monthYear,'end')
    select_day(driver,day,'end')

def upload_file_sharepoint(site_name,filename):
    authcookie = Office365('https://netorgft3300690.sharepoint.com/',username='automations@globecommerce.co', password='u#Fz6cf&jPU&U3').GetCookies()
    site = Site('https://netorgft3300690.sharepoint.com/sites/GlobeCommerceInternal/', version=Version.v365, authcookie=authcookie)
    if site_name=='kingcamp':
        folder = site.Folder('Shared Documents/Order Operations Team/PO Data for Dashboards/Sell Thru Report/Spreetail Daily Sales/KingCamp')
    elif site_name=='xd':
        folder = site.Folder('Shared Documents/Order Operations Team/PO Data for Dashboards/Sell Thru Report/Spreetail Daily Sales/XD Design')
    with open(filename, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, filename)
    print('File uploaded')


def get_data(driver,site):
    date_today = datetime.datetime.now()
    date_yesterday = date_today-datetime.timedelta(days=1)
    searchDate = date_yesterday.strftime('%d %B %Y')
    print(searchDate)
    time.sleep(5)
    if site=='kc':
        driver.get(kingcamp_sales_url)
        filename = 'sales_data_{}_{}.csv'.format(searchDate,site)
    if site=='xd':
        driver.get(xindao_sales_url)
        filename = 'sales_data_{}_{}.csv'.format(searchDate,site)
    change_start_date(driver,searchDate)
    time.sleep(3)
    change_end_date(driver,searchDate)
    time.sleep(3)
    change_start_date(driver,searchDate)
    time.sleep(5)
    column_names = ['itemId','partNumber','title','quantityOnHand','unitSales','unitSalesPerDay','vsPreviousPeriod','vsPreviousYear','amazon','bestBuy','ebay','facebook','googleExpress','homedepot','kohls','kroger','lowes','macys','newegg','overstock','sears','target','vminnovations','vminnvovation','walmart','wayfair','wish','Date']
    driver.find_element(By.ID, 'react-select-2-input').send_keys('100\n')
    time.sleep(5)
    obj = {}
    obj['itemId'] = []
    obj['partNumber'] = []
    obj['title'] = []
    obj['quantityOnHand'] = []
    obj['unitSales'] = []
    obj['unitSalesPerDay'] = []
    obj['vsPreviousPeriod'] = []
    obj['vsPreviousYear'] = []
    obj['amazon'] = []
    obj['bestBuy'] = []
    obj['ebay'] = []
    obj['facebook'] = []
    obj['googleExpress'] = []
    obj['homedepot'] = []
    obj['kohls'] = []
    obj['kroger'] = []
    obj['lowes'] = []
    obj['macys'] = []
    obj['newegg'] = []
    obj['overstock'] = []
    obj['sears'] = []
    obj['target'] = []
    obj['vminnovations'] = []
    obj['vminnvovations'] = []
    obj['walmart'] = []
    obj['wayfair'] = []
    obj['wish'] = []
    obj['Date'] = []
    table = driver.find_element(By.CLASS_NAME,'rt-table')
    rows = table.find_elements(By.CLASS_NAME,'rt-tr-group')
    for row in rows:
        itemId = row.find_element(By.CLASS_NAME,'rt-td').find_elements(By.TAG_NAME,'p')[1].get_attribute('data-for').split('-')[-1]
        row_data = row.text.split('\n')
        row_data_fixed = []
        for data in row_data:
            if data=='-':
                row_data_fixed.append(0)
            
            else:
                row_data_fixed.append(data)     
        obj['itemId'].append(itemId)
        obj['partNumber'].append(row_data_fixed[0])
        try:
            obj['title'].append(driver.execute_script("return document.getElementById('tooltip-{}').innerHTML".format(itemId)))
        except:
            obj['title'].append(row_data_fixed[0])
        obj['quantityOnHand'].append(row_data_fixed[2])
        obj['unitSales'].append(row_data_fixed[3])
        unitSales = int(row_data_fixed[3])
        obj['unitSalesPerDay'].append(row_data_fixed[4])
        obj['vsPreviousPeriod'].append(row_data_fixed[5])
        obj['vsPreviousYear'].append(row_data_fixed[6])
        if row_data_fixed[7] == 0:
            obj['amazon'].append(row_data_fixed[7])
        else:
            percentageUnit = fix_percentage(row_data_fixed[7])
            amazonUnit = ((percentageUnit)*(unitSales))/100
            obj['amazon'].append(round(amazonUnit))
        if row_data_fixed[8] == 0:
            obj['bestBuy'].append(row_data_fixed[8])
        else:
            percentageUnit = fix_percentage(row_data_fixed[8])
            bestbuyUnit = ((percentageUnit)*(unitSales))/100
            obj['bestBuy'].append(round(bestbuyUnit))
        if row_data_fixed[9] == 0:
            obj['ebay'].append(row_data_fixed[9])
        else:
            percentageUnit = fix_percentage(row_data_fixed[9])
            ebayUnit = ((percentageUnit)*(unitSales))/100
            obj['ebay'].append(round(ebayUnit))
        if row_data_fixed[10] == 0:
            obj['facebook'].append(row_data_fixed[10])
        else:
            percentageUnit = fix_percentage(row_data_fixed[10])
            facebookUnit = ((percentageUnit)*(unitSales))/100
            obj['facebook'].append(round(facebookUnit))
        if row_data_fixed[11] == 0:
            obj['googleExpress'].append(row_data_fixed[11])
        else:
            percentageUnit = fix_percentage(row_data_fixed[11])
            googleexpressUnit = ((percentageUnit)*(unitSales))/100
            obj['googleExpress'].append(round(googleexpressUnit))
        if row_data_fixed[12] == 0:
            obj['homedepot'].append(row_data_fixed[12])
        else:
            percentageUnit = fix_percentage(row_data_fixed[12])
            homedepotUnit = ((percentageUnit)*(unitSales))/100
            obj['homedepot'].append(round(homedepotUnit))
        if row_data_fixed[13] == 0:
            obj['kohls'].append(row_data_fixed[13])
        else:
            percentageUnit = fix_percentage(row_data_fixed[13])
            kohlsUnit = ((percentageUnit)*(unitSales))/100
            obj['kohls'].append(round(kohlsUnit))
        if row_data_fixed[14] == 0:
            obj['kroger'].append(row_data_fixed[14])
        else:
            percentageUnit = fix_percentage(row_data_fixed[14])
            krogerUnit = ((percentageUnit)*(unitSales))/100
            obj['kroger'].append(round(krogerUnit))
        if row_data_fixed[15] == 0:
            obj['lowes'].append(row_data_fixed[15])
        else:
            percentageUnit = fix_percentage(row_data_fixed[15])
            lowesUnit = ((percentageUnit)*(unitSales))/100
            obj['lowes'].append(round(lowesUnit))
        if row_data_fixed[16] == 0:
            obj['macys'].append(row_data_fixed[16])
        else:
            percentageUnit = fix_percentage(row_data_fixed[16])
            macysUnit = ((percentageUnit)*(unitSales))/100
            obj['macys'].append(round(macysUnit))
        if row_data_fixed[17] == 0:
            obj['newegg'].append(row_data_fixed[17])
        else:
            percentageUnit = fix_percentage(row_data_fixed[17])
            neweggUnit = ((percentageUnit)*(unitSales))/100
            obj['newegg'].append(round(neweggUnit))
        if row_data_fixed[18] == 0:
            obj['overstock'].append(row_data_fixed[18])
        else:
            percentageUnit = fix_percentage(row_data_fixed[18])
            overstockUnit = ((percentageUnit)*(unitSales))/100
            obj['overstock'].append(round(overstockUnit))
        if row_data_fixed[19] == 0:
            obj['sears'].append(row_data_fixed[19])
        else:
            percentageUnit = fix_percentage(row_data_fixed[19])
            searsUnit = ((percentageUnit)*(unitSales))/100
            obj['sears'].append(round(searsUnit))
        if row_data_fixed[20] == 0:
            obj['target'].append(row_data_fixed[20])
        else:
            percentageUnit = fix_percentage(row_data_fixed[20])
            targetUnit = ((percentageUnit)*(unitSales))/100
            obj['target'].append(round(targetUnit))
        if row_data_fixed[21] == 0:
            obj['vminnovations'].append(row_data_fixed[21])
        else:
            percentageUnit = fix_percentage(row_data_fixed[21])
            vminnovationsUnit = ((percentageUnit)*(unitSales))/100
            obj['vminnovations'].append(round(vminnovationsUnit))
        if row_data_fixed[22] == 0:
            obj['vminnvovations'].append(row_data_fixed[22])
        else:
            percentageUnit = fix_percentage(row_data_fixed[22])
            vminnvovationsUnit = ((percentageUnit)*(unitSales))/100
            obj['vminnvovations'].append(round(vminnvovationsUnit))
        if row_data_fixed[23] == 0:
            obj['walmart'].append(row_data_fixed[23])
        else:
            percentageUnit = fix_percentage(row_data_fixed[23])
            walmartUnit = ((percentageUnit)*(unitSales))/100
            obj['walmart'].append(round(walmartUnit))
        if row_data_fixed[24] == 0:
            obj['wayfair'].append(row_data_fixed[24])
        else:
            percentageUnit = fix_percentage(row_data_fixed[24])
            wayfairUnit = ((percentageUnit)*(unitSales))/100
            obj['wayfair'].append(round(wayfairUnit))
        if row_data_fixed[25] == 0:
            obj['wish'].append(row_data_fixed[25])
        else:
            percentageUnit = fix_percentage(row_data_fixed[25])
            wishUnit = ((percentageUnit)*(unitSales))/100
            obj['wish'].append(round(wishUnit))
        obj['Date'].append(searchDate)
    df=pd.DataFrame(obj)
    df['vsPreviousPeriod'] = df['vsPreviousPeriod'].apply(fix_percentage)
    df.to_csv(filename)
    return filename



def main():
    print('##################################### Logging in ##############################################')
    driver = start_driver('No')
    login(driver,website_url,username,pw)
    print('##################################### Getting king camp sales data ##############################################')
    kc_filename = get_data(driver,'kc')
    print('##################################### Getting xd sales data ##############################################')
    xd_filename = get_data(driver,'xd')
    print('##################################### Uploading king camp sales data ##############################################')
    upload_file_sharepoint('kingcamp',kc_filename)
    print('##################################### Uploading xd sales data ##############################################')
    upload_file_sharepoint('xd',xd_filename)
    driver.close()


main()