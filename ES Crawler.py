import re
import os
from selenium import webdriver
import time
import pandas as pd

# start=time.time()
# Proxy setting
os.environ['http_proxy'] = 'http://127.0.0.1:1080'
os.environ['https_proxy'] = 'https://127.0.0.1:1080'

# set up the primary element for our needed data
month=int(input("Which month's data do we want? (use num) : "))
year=int(input("Which year's data do we want? (use num) : "))

# supplier name: qty, price
prod={}

# Set up the browser options for better performance
options=webdriver.FirefoxOptions()
# FireFox setting
options.headless = True
options.set_preference('permissions.default.image',2)
options.set_preference("javascript.enabled",False)

# Chrome setting
#options.add_experimental_option('excludeSwitches', ['enable-automation'])
#options.add_experimental_option('useAutomationExtension', False)
#prefs = {'profile.default_content_setting_values': {'javascript': 2}}
#options.add_experimental_option('prefs', prefs)
#options.add_argument('--headless') headless would be detected as crawler
#options.add_argument('--disable-gpu')
#options.page_load_strategy = 'eager'

# apply headers and open the url
driver=webdriver.Firefox(options=options,executable_path="C:\\Users\\inven\\anaconda3\\geckodriver")
driver.get("https://www.esutures.com/account/login/")

# driver.find_element_by_xpath('/html/body/scipt/div[2]/div/button').click()
# input account name and password
driver.find_element_by_id("login_id").send_keys("mdujowich@vrcvet.com")
driver.find_element_by_id("pass").send_keys("1820Monterey")
# submit passcode
driver.find_element_by_id("loginBtn").click()

# to avoid IP block, set time off
time.sleep(2)

# open account history
url='https://www.esutures.com/account/history/'
driver.get(url)

# get the order history
orders=driver.find_elements_by_xpath('//*[@id="mainWrap"]/div/table/tbody/tr')
for order in orders[1:]:

    # decode the each order url and get the order detail
    order_t=order.text
    date=int(order_t.split('/')[0][-2:].replace(' ',''))
    y=int(order_t.split('/')[2][:4])
    if y < year:
        break
    else:
        if y == year:

            if date == month:

                order_num = order_t[:6]

                order_url = 'https://www.esutures.com/account/history/?orderId=' + order_num
                # pull target month and open the detailed order page with new page
                driver.execute_script("window.open();")
                handles=driver.window_handles
                driver.switch_to.window(handles[1])
                driver.get(order_url)
                time.sleep(1)

                # enter the specific order page and read the product information using re
                # 2023/07/05 ES update the website using the new code
                items=driver.find_elements_by_xpath('//*[@id="items"]')
                for item in items:

                    txt=item.get_attribute('innerHTML')
                    pattern_name = re.compile('<div class="d item_description">(.*?)</div>',re.S)
                    pattern_qty = re.compile('<td class="c d item_quantity">(.*?)</td>', re.S)
                    pattern_price = re.compile('<td class="d r item_price">(.*?)</td>', re.S)
                    pattern_unit = re.compile('<td class="c d item_quantity">.*?</td><td>(.*?)</td><td class="d r item_price">', re.S)
                    pattern_sku = re.compile('<td><span style="color:#0000FF;">(.*?)</span>', re.S)
                    name=pattern_name.findall(txt)
                    qty_list=pattern_qty.findall(txt)
                    price_list=pattern_price.findall(txt)
                    unit_list=pattern_unit.findall(txt)
                    sku_list=pattern_sku.findall(txt)

                    # convert the product list to each product
                    for i in enumerate(name):
                        qty=int(qty_list[i[0]])
                        price=float(price_list[i[0]].replace('$','').replace(',',''))
                        unit=unit_list[i[0]]
                        sku=sku_list[i[0]]
                        # record the ES Hist as this format
                        if i[1] in prod:
                            prod[i[1]][1] = prod[i[1]][1] + qty
                        else:
                            prod[i[1]] = [sku, qty, price, unit]
                        # close order page and back to the order history
                        # time.sleep(5)
                        #driver.close()
                        driver.switch_to.window(handles[0])

        else:
            continue

driver.quit()

# convert the data to data frame to export to Excel
df=pd.DataFrame(prod).transpose().reset_index()
df.columns=['name','sku','qty','price','unit']

# output the data to excel, change the file path if needed
writer = pd.ExcelWriter("C:\Inv Data\Purchase History\爬虫数据\Es Hist {}.xlsx".format(month), engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()

# Ending words
print('OK, WHO IS THE NEXT!')


