import pandas as pd
import time
import datetime
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
    # options.add_argument('--headless')
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
            print('Logged in')
            break
        except Exception as e:
            print(e)
            time.sleep(2)


def get_second_sunday_date_march(year):
    date = '{}-03-01'.format(year)
    date_formatted = datetime.datetime.strptime(date,'%Y-%m-%d')
    all_sundays = []
    while True:
        dayName = date_formatted.strftime('%A')
        if dayName=='Sunday':
            all_sundays.append(date_formatted)
        date_formatted = date_formatted+datetime.timedelta(days=1)
        if date_formatted.strftime('%m')=='04':
            break
    return all_sundays[1]


def get_first_sunday_date_november(year):
    date = '{}-11-01'.format(year)
    date_formatted = datetime.datetime.strptime(date,'%Y-%m-%d')
    all_sundays = []
    while True:
        dayName = date_formatted.strftime('%A')
        if dayName=='Sunday':
            all_sundays.append(date_formatted)
        date_formatted = date_formatted+datetime.timedelta(days=1)
        if date_formatted.strftime('%m')=='04':
            break
    return all_sundays[0]



def get_tz_string_for_start_date(date):
    date_year = date.split('-')[0]
    march_date,nov_date = get_second_sunday_date_march(date_year),get_first_sunday_date_november(date_year)
    date_formatted = datetime.datetime.strptime(date,'%Y-%m-%d')
    if date_formatted>march_date and date_formatted<=nov_date:
        tz_string = 'T05:00:00.000'
    else:
        tz_string = 'T06:00:00.000'
    return tz_string



def get_tz_string_for_end_date(date):
    date_year = date.split('-')[0]
    march_date,nov_date = get_second_sunday_date_march(date_year),get_first_sunday_date_november(date_year)
    date_formatted = datetime.datetime.strptime(date,'%Y-%m-%d')
    if date_formatted>march_date and date_formatted<=nov_date:
        tz_string = 'T04:59:59.999'
    else:
        tz_string = 'T05:59:59.999'
    return tz_string


def generate_api_url_for_date_xindoa(start_date,end_date):#date format: 2021-12-01
    tz_string_start_date = get_tz_string_for_start_date(start_date)
    tz_string_end_date = get_tz_string_for_end_date(end_date)
    url = 'https://sales.partner-api-prod.spreetail.com/api/vendors/xindao-us-ltd/sales-summary?pageNumber=1&pageSize=100&orderBy=unitSales&orderByDesc=true&partNumberFilterText=&periodStart={startDate}{startDateTz}Z&periodEnd={endDate}{endDateTz}Z&withMarketplaceData=true'.format(startDate=start_date,endDate=end_date,startDateTz=tz_string_start_date,endDateTz=tz_string_end_date)
    return url


def generate_api_url_for_date_kingcamp(start_date,end_date):#date format: 2021-12-01
    tz_string_start_date = get_tz_string_for_start_date(start_date)
    tz_string_end_date = get_tz_string_for_end_date(end_date)
    url = 'https://sales.partner-api-prod.spreetail.com/api/vendors/kingcamp-outdoor-products-co-ltd/sales-summary?pageNumber=1&pageSize=100&orderBy=unitSales&orderByDesc=true&partNumberFilterText=&periodStart={startDate}{startDateTz}Z&periodEnd={endDate}{endDateTz}Z&withMarketplaceData=true'.format(startDate=start_date,endDate=end_date,startDateTz=tz_string_start_date,endDateTz=tz_string_end_date)
    return url

# def generate_api_url_for_date_xindoa_dec(start_date,end_date):#date format: 2021-12-01
#     # switch to 06:00 time zone point 8nov
#     #before nov first sunday =>05:00<= after march 2nd sunday
#     #after nov first sunday =>06:00<= before march 2nd sunday
#     url = 'https://sales.partner-api-prod.spreetail.com/api/vendors/xindao-us-ltd/sales-summary?pageNumber=1&pageSize=100&orderBy=unitSales&orderByDesc=true&partNumberFilterText=&periodStart={startDate}{startDateTz}&periodEnd={endDate}{endDateTz}Z&withMarketplaceData=true'.format(startDate=start_date,endDate=end_date,startDateTz=tz_string_start_date,endDateTz=tz_string_end_date)
#     return url

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


def get_sales_data(date,site):# format '2021-05-13'
    user_agent_pc = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36'
    session = requests.Session()
    session.headers.update({'User-Agent': user_agent_pc})
    token = 'Bearer eyJraWQiOiJXaklrWkk1OXlSVHhDaGtkc1hYdWdHdnRPU0NVQllLOE5reFNOd1I5TjFjIiwiYWxnIjoiUlMyNTYifQ.eyJ2ZXIiOjEsImp0aSI6IkFULjZOTUUwMEtOZzdfdGozSEhfcno3U2Q2d0NIc3o5M2hSVW9qRk51WHpyZkUiLCJpc3MiOiJodHRwczovL3NwcmVldGFpbC5va3RhLmNvbS9vYXV0aDIvZGVmYXVsdCIsImF1ZCI6ImFwaTovL2RlZmF1bHQiLCJpYXQiOjE2Mzk1MzgyODYsImV4cCI6MTYzOTYyNDY4NiwiY2lkIjoiMG9hN2lzcXp3cFZKYWpXN1ozNTciLCJ1aWQiOiIwMHVkYjhzanlvM1NwdVJoMzM1NyIsInNjcCI6WyJwcm9maWxlIiwib3BlbmlkIiwiZW1haWwiXSwic3ViIjoiaGFycmlzb25AZ2xvYmVjb21tZXJjZS5jbyIsImZpcnN0TmFtZSI6IkhhcnJpc29uIiwibGFzdE5hbWUiOiJLcmF0dGVyIn0.Vuxh3zp37n4_YGxw4ljHa0nG94FF2OH2oJRLC18x-s_Wjpe49h6YgIRZG1zO-ZvW6MaEPl73U4Kr9BMfwqB7J6VmmYiPxtO_HXUkvpXzP8JlkwLeddew9T5NSfnfEItqgFHK4nJMy75xf3DYqqTJ0FgVzJD4nXJP7tB_ivo1QCuVg0XtDO8sDwiKfYZEuxBJvnfMGZihbSbGdmUOtPCoAzqJYeMJJqsoe0lWcqNn4x1VRn158x2tRmR4Z5CxaIFex9QJ_B2QSViOdYtcNsBBIn_75NTiNv2zPnRbQYPjg4yupZq-KX-6cywEcTonIEfpe5p5cUBktzknb8RHLqaQ7A'
    headers = {'Authorization': token,
            'contenttype':'application/json',
            'sec-fetch-mode':'cors',
            'sec-fetch-site':'same-site',
            'authority':'sales.partner-api-prod.spreetail.com'}
    all_dfs = []

    next_date = (datetime.datetime.strptime(date,'%Y-%m-%d')+datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    if site=='kingcamp':
        url = generate_api_url_for_date_kingcamp(date,next_date)
    elif site=='xd':
        url = generate_api_url_for_date_xindoa(date,next_date)
    resp = session.get(url,headers=headers)
    json_data = resp.json()['page']['nodes']
    df = pd.DataFrame(json_data)
    df = df.fillna(0)
    df['vsPreviousPeriod'] = df['vsPreviousPeriod'].apply(fix_percentage)
    df['Date'] = date
    filename = 'sales_data_{}_{}.xlsx'.format(date,site)
    df.to_excel(filename)
    return filename


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
        element_list = [td for td in driver.find_element(By.CLASS_NAME,'rdtDays').find_elements(By.TAG_NAME,'td') if td.text==day]
        for elements in element_list:
            if elements.get_attribute('class') == 'rdtOld' or 'rdtNew':
                pass
            else:
                elements.click()
    else:
        element_list = [td for td in driver.find_elements(By.CLASS_NAME,'rdtDays')[1].find_elements(By.TAG_NAME,'td') if td.text==day]
        for elements in element_list:
            if elements.get_attribute('class') == 'rdtOld' or 'rdtNew':
                pass
            else:
                elements.click()
        



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


def get_data(driver,site):
    date_today = datetime.datetime.now()
    date_yesterday = date_today-datetime.timedelta(days=1)
    date_before_yesterday = date_today-datetime.timedelta(days=2)
    searchDate = date_yesterday.strftime('%d %B %Y')
    time.sleep(10)
    if site=='kc':
        driver.get(kingcamp_sales_url)
        filename = 'sales_data_{}_{}.csv'.format(searchDate,site)
    if site=='xd':
        driver.get(xindao_sales_url)
        filename = 'sales_data_{}_{}.csv'.format(searchDate,site)

    change_start_date(driver,searchDate)
    time.sleep(5)
    change_end_date(driver,searchDate)
    time.sleep(5)
    change_start_date(driver,searchDate)
    time.sleep(10)
    column_names = ['itemId','partNumber','title','quantityOnHand','unitSales','unitSalesPerDay','vsPreviousPeriod','vsPreviousYear','Date']
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
        obj['unitSalesPerDay'].append(row_data_fixed[4])
        obj['vsPreviousPeriod'].append(row_data_fixed[5])
        obj['vsPreviousYear'].append(row_data_fixed[6])
        obj['Date'].append(searchDate)

    df=pd.DataFrame(obj)
    df['vsPreviousPeriod'] = df['vsPreviousPeriod'].apply(fix_percentage)
    df.to_csv(filename)
    return filename


def find_sleep_time():
    time_now = datetime.datetime.now()
    today_date = time_now.strftime('%d-%m-%Y')
    tomorrow_date = (time_now+datetime.timedelta(days=1)).strftime('%d-%m-%Y')
    target_time = datetime.datetime.strptime('{} 13:00:00'.format(tomorrow_date),'%d-%m-%Y %H:%M:%S')
    difference = target_time-time_now
    print('Sleeping for {} seconds (approx {} hours)'.format(difference.seconds,difference.seconds/3600))
    return difference.seconds

def main():
    print('##################################### Logging in ##############################################')
    driver = start_driver('No')
    login(driver,website_url,username,pw)
    print('##################################### Getting king camp sales data ##############################################')
    kc_filename = get_data(driver,'kc')
    print('##################################### Getting xd sales data ##############################################')
    xd_filename = get_data(driver,'xd')
    print('##################################### Uploading king camp sales data ##############################################')
    # upload_file_sharepoint('kingcamp',kc_filename)
    print('##################################### Uploading xd sales data ##############################################')
    # upload_file_sharepoint('xd',xd_filename)
    driver.close()

# while True:
#     main()
#     time.sleep(find_sleep_time())

main()