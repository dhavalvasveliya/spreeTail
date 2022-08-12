
from fileinput import filename
from regex import F
import requests
import time
import json
import pandas as pd
import datetime
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

# url = 'https://sales.partner-api-prod.spreetail.com/api/vendors/xindao-us-ltd/sales-summary?partNumberFilterText=&periodEnd=2021-11-02T00:59:59.000-04:00&periodStart=2021-11-01T01:00:00.000-04:00&withMarketplaceData=true&pageSize=10&pageNumber=1&orderByDesc=true&orderBy=unitSales&orderByChannelName=&totalPageCount=5'
website_url = 'http://partner.spreetail.com/'
username = 'automations@globecommerce.co'
pw = '7k*hXbz565H$Fp'


# r = requests.get(url, headers=headers)

class LocalStorage:

    def __init__(self, driver) :
        self.driver = driver

    def __len__(self):
        return self.driver.execute_script("return window.localStorage.length;")

    def items(self) :
        return self.driver.execute_script( \
            "var ls = window.localStorage, items = {}; " \
            "for (var i = 0, k; i < ls.length; ++i) " \
            "  items[k = ls.key(i)] = ls.getItem(k); " \
            "return items; ")

    def keys(self) :
        return self.driver.execute_script( \
            "var ls = window.localStorage, keys = []; " \
            "for (var i = 0; i < ls.length; ++i) " \
            "  keys[i] = ls.key(i); " \
            "return keys; ")

    def get(self, key):
        return self.driver.execute_script("return window.localStorage.getItem(arguments[0]);", key)

    def set(self, key, value):
        self.driver.execute_script("window.localStorage.setItem(arguments[0], arguments[1]);", key, value)

    def has(self, key):
        return key in self.keys()

    def remove(self, key):
        self.driver.execute_script("window.localStorage.removeItem(arguments[0]);", key)

    def clear(self):
        self.driver.execute_script("window.localStorage.clear();")

    def __getitem__(self, key) :
        value = self.get(key)
        if value is None :
          raise KeyError(key)
        return value

    def __setitem__(self, key, value):
        self.set(key, value)

    def __contains__(self, key):
        return key in self.keys()

    def __iter__(self):
        return self.items().__iter__()

    def __repr__(self):
        return self.items().__str__()



def start_driver():
    user_agent_pc = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36'
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--user-agent={user_agent}'.format(user_agent=user_agent_pc))
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    return driver

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
            storage = LocalStorage(driver)
            all_key = json.loads(storage.get("okta-token-storage"))
            break
        except Exception as e:
            print(e)
            time.sleep(2)      
    return all_key["accessToken"]["accessToken"]

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

def get_sales_data(site,token):
    dateToday = datetime.datetime.now()
    searchDate = dateToday - datetime.timedelta(days=2)
    date = searchDate.strftime("%Y-%m-%d")
    headers = {
        "accept": "application/json, text/plain, */*",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "en-US,en;q=0.9",
        "authorization": "Bearer " + token,
        "contenttype": "application/json",
        "origin": "https://partner.spreetail.com",
        "referer": "https://partner.spreetail.com/",
        "sec-ch-ua-mobile": "?1",
        "sec-ch-ua-platform": "Android",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-site",
        "user-agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Mobile Safari/537.36"
    }
    next_date = (searchDate + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    if site=='kc':
        url = generate_api_url_for_date_kingcamp(date,next_date)
    elif site=='xd':
        url = generate_api_url_for_date_xindoa(date,next_date)
    
    r = requests.get(url,headers=headers)
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
    json_data = r.json()['nodes']
    for data in json_data:
        unitsales = data['unitSales']
        def convertToUnit(ratioOfSales):
            try:
                return (int(ratioOfSales*unitsales))
            except:
                return ratioOfSales
        obj['itemId'].append(data['itemId'])
        obj['partNumber'].append(data['partNumber'])
        obj['title'].append(data['title'])
        obj['quantityOnHand'].append(data['qohEndingNetwork'])
        obj['unitSales'].append(data['unitSales'])
        obj['unitSalesPerDay'].append(data['unitSalesPerDay'])
        try:
            convertedValue = data['vsPreviousPeriod']
            percentageValue = convertedValue*100
            obj['vsPreviousPeriod'].append("{:.0f}".format(int(round(percentageValue,0))))
        except:
            obj['vsPreviousPeriod'].append(data['vsPreviousPeriod'])
        obj['vsPreviousYear'].append(data['vsPreviousYear'])
        if "channelSalesDistribution" in data:
            try:                
                obj['amazon'].append(convertToUnit(int(data['channelSalesDistribution']['amazon']['ratioOfSales'])))
            except:
                obj['amazon'].append(0)
            try:
                obj['bestBuy'].append(convertToUnit(int(data['channelSalesDistribution']['bestbuy']['ratioOfSales'])))
            except:
                obj['bestBuy'].append(0)
            try:
                obj['ebay'].append(convertToUnit(int(data['channelSalesDistribution']['ebay']['ratioOfSales'])))
            except:
                obj['ebay'].append(0)
            try:
                obj['facebook'].append(convertToUnit(int(data['channelSalesDistribution']['facebook']['ratioOfSales'])))
            except:
                obj['facebook'].append(0)
            try:
                obj['googleExpress'].append(convertToUnit(int(data['channelSalesDistribution']['googleexpress']['ratioOfSales'])))
            except:
                obj['googleExpress'].append(0)
            try:
                obj['homedepot'].append(convertToUnit(int(data['channelSalesDistribution']['homedepot']['ratioOfSales'])))
            except:
                obj['homedepot'].append(0)
            try:
                obj['kohls'].append(convertToUnit(int(data['channelSalesDistribution']['kohls']['ratioOfSales'])))
            except:
                obj['kohls'].append(0)
            try:
                obj['kroger'].append(convertToUnit(int(data['channelSalesDistribution']['kroger']['ratioOfSales'])))
            except:
                obj['kroger'].append(0)
            try:
                obj['lowes'].append(convertToUnit(int(data['channelSalesDistribution']['lowes']['ratioOfSales'])))
            except:
                obj['lowes'].append(0)
            try:
                obj['macys'].append(convertToUnit(int(data['channelSalesDistribution']['macys']['ratioOfSales'])))
            except:
                obj['macys'].append(0)
            try:
                obj['newegg'].append(convertToUnit(int(data['channelSalesDistribution']['newegg']['ratioOfSales'])))
            except:
                obj['newegg'].append(0)
            try:
                obj['overstock'].append(convertToUnit(int(data['channelSalesDistribution']['overstock']['ratioOfSales'])))
            except:
                obj['overstock'].append(0)
            try:
                obj['sears'].append(convertToUnit(int(data['channelSalesDistribution']['sears']['ratioOfSales'])))
            except:
                obj['sears'].append(0)
            try:
                obj['target'].append(convertToUnit(int(data['channelSalesDistribution']['target']['ratioOfSales'])))
            except:
                obj['target'].append(0)
            try:
                obj['vminnovations'].append(convertToUnit(int(data['channelSalesDistribution']['vminnovations']['ratioOfSales'])))
            except:
                obj['vminnovations'].append(0)
            try:
                obj['vminnvovations'].append(convertToUnit(int(data['channelSalesDistribution']['vminnvovations']['ratioOfSales'])))
            except:
                obj['vminnvovations'].append(0)
            try:
                obj['walmart'].append(convertToUnit(int(data['channelSalesDistribution']['walmart']['ratioOfSales'])))
            except:
                obj['walmart'].append(0)
            try:
                obj['wayfair'].append(convertToUnit(int(data['channelSalesDistribution']['wayfair']['ratioOfSales'])))
            except:
                obj['wayfair'].append(0)
            try:
                obj['wish'].append(convertToUnit(int(data['channelSalesDistribution']['wish']['ratioOfSales'])))
            except:
                obj['wish'].append(0)
        else:
            obj['amazon'].append(0)
            obj['bestBuy'].append(0)
            obj['ebay'].append(0)
            obj['facebook'].append(0)
            obj['googleExpress' ].append(0)
            obj['homedepot'].append(0)
            obj['kohls'].append(0)
            obj['kroger'].append(0)
            obj['lowes'].append(0)
            obj['macys'].append(0)
            obj['newegg'].append(0)
            obj['overstock'].append(0)
            obj['sears'].append(0)
            obj['target'].append(0)
            obj['vminnovations'].append(0)
            obj['vminnvovations'].append(0)
            obj['walmart'].append(0)
            obj['wayfair'].append(0)
            obj['wish'].append(0)
        obj['Date'] = searchDate    
    df = pd.DataFrame(obj)
    df.fillna(0, inplace=True)
    filename = 'sales_data_{}_{}.csv'.format(date,site)
    df.to_csv(filename)
    return filename

def upload_file_sharepoint(site_name,filename):
    authcookie = Office365('https://netorgft3300690.sharepoint.com/',username='automations@globecommerce.co', password='u#Fz6cf&jPU&U3').GetCookies()
    site = Site('https://netorgft3300690.sharepoint.com/sites/GlobeCommerceInternal/', version=Version.v365, authcookie=authcookie)
    if site_name=='kc':
        folder = site.Folder('Shared Documents/Order Operations Team/PO Data for Dashboards/Sell Thru Report/Spreetail Daily Sales/KingCamp')
    elif site_name=='xd':
        folder = site.Folder('Shared Documents/Order Operations Team/PO Data for Dashboards/Sell Thru Report/Spreetail Daily Sales/XD Design')
    with open(filename, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, filename)
    print('File uploaded')


driver = start_driver()
jwtToken = login(driver,website_url,username,pw)
driver.close()
kc_filename = get_sales_data('kc',jwtToken)
xd_filename = get_sales_data('xd',jwtToken)
upload_file_sharepoint('kc',kc_filename)
upload_file_sharepoint('xd',xd_filename)
