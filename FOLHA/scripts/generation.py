from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import re
import datetime

def number_format(s):
    if s < 10:
        return '0' + str(s)
    return str(s)

def generate_csv():
    # settings -----------------------
    USER_EMAIL = 'romullooliveira@ppa.com'
    USER_PASSWORD = '123456'
    PAT1 = r'(\d+?) (\d\d\/\d\d\/\d\d\d\d) (.+?) (\d\d:\d\d)'
    PAT2 = r'(\d+?) (\d\d\/\d\d\/\d\d\d\d) (.+?) (\d\d:\d\d) (\d\d:\d\d)'
    PAT3 = r'Página 1 de (\d+?)\.'
    START_DATE = datetime.datetime.strptime(input("Digite data de início (DD/MM/YYYY): "), '%d/%m/%Y')
    END_DATE = datetime.datetime.strptime(input("Digite data de término (DD/MM/YYYY): "), '%d/%m/%Y')
    LOCALE = input("Unidade: ").upper()
    driver = webdriver.Firefox()
    # --------------------------------

    # do login -----------------------
    driver.get("http://webppa.ddns.net/login")
    username = driver.find_element_by_name("email")
    username.clear()
    username.send_keys(USER_EMAIL)

    password = driver.find_element_by_name("password")
    password.clear()
    password.send_keys(USER_PASSWORD)
    password.send_keys(Keys.RETURN)
    time.sleep(2)
    # --------------------------------

    # get folha ----------------------
    done = False
    data = []
    driver.get('http://webppa.ddns.net/dashboard-register-job?grid=list&page=1')
    time.sleep(2)
    pages = int(re.search(PAT3, driver.page_source).group(1))
    for page in range(pages):
        if done: 
            break
        print('Página: {}'.format(page + 1))
        if page != 0: # do not reload first page
            page = page + 1
            driver.get('http://webppa.ddns.net/dashboard-register-job?grid=list&page={}'.format(page))
            time.sleep(2)
        table = driver.find_element_by_class_name('table-responsive')
        tbody = table.find_element_by_tag_name('tbody')
        rows = tbody.find_elements_by_tag_name('tr')
        for row in rows:
            match = re.search(PAT2, row.text)
            rdata = None
            print(repr(row.text))
            try:
                if match:
                    rdata = list(match.groups())
                else:
                    rdata = list(re.search(PAT1, row.text).groups()) + ['00:00']
            except:
                continue
            date = datetime.datetime.strptime(rdata[1], '%d/%m/%Y')
            if date > END_DATE:
                continue
            if date < START_DATE:
                done = True
                break
            data.append(rdata)
    # --------------------------------

    # save folha ---------------------
    data_str = ''
    for line in data:
        data_str += ','.join(line) + '\n'

    outpt = 'inputs/FOLHA_{}_{}-{}.csv'.format(LOCALE, number_format(START_DATE.day), number_format(START_DATE.month))
    with open(outpt, 'w') as f:
        f.write(data_str) 
    # --------------------------------

    driver.close()
    return outpt
