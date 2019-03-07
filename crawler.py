from selenium import webdriver
import time
import pandas as pd


#前置輸入

#排名1~2000
min_count = 1 
max_count = 2000  
#帳號密碼
username_edit = 'a123456'
password_edit = '0000000'

#載入瀏覽驅動
driver = webdriver.Chrome('C:\selenium_driver_chrome\chromedriver.exe') 

#網頁起始點
driver.get('http://ezproxy.nutc.edu.tw/login?url=http://www.credit.com.tw/Special/')
time.sleep(5)

#帳號密碼欄位
username = driver.find_element_by_xpath("//input[@name='user']")
password = driver.find_element_by_xpath("//input[@name='pass']")

#帳號密碼
username.send_keys(username_edit)
password.send_keys(password_edit)
password.submit()


time.sleep(5) # Let the user actually see something!


#確認按鈕
confirmbtn = driver.find_element_by_class_name('con02a')
confirmbtn.submit()

time.sleep(5)

#直接設定URL 跳過  herf問題
driver.get('https://www.credit.com.tw.ezproxy.nutc.edu.tw/Special/database/newtop5000/index.cfm?Fuseaction=Top5000')
# driver.find_element_by_link_text("TOP5000企業排名").click()

#點Radio Button
driver.find_element_by_xpath("//input[@name='search_type' and @value='4']").click()


#設定抓取筆數
index_from = driver.find_element_by_name('search_condition42')
index_from.send_keys(str(min_count))
index_to = driver.find_element_by_name('search_condition43')
index_to.send_keys(str(max_count))

#確定
index_to.submit()

time.sleep(10)

#取得公司名稱 及 排名

string_list =[]
rank_list = []

def getCompanyName():
    
    #  testString = driver.find_element_by_xpath("//html/body/table[1]/tbody/tr[4]/td/table/tbody/tr[2]/td[3]")
 
    for i in range(2, 7):
       
        companyRank = driver.find_element_by_xpath('//html/body/table[1]/tbody/tr[4]/td/table/tbody/tr['+str(i)+']/td[2]')
        companyNameString = driver.find_element_by_xpath('//html/body/table[1]/tbody/tr[4]/td/table/tbody/tr['+str(i)+']/td[3]')
        print(companyRank.text , companyNameString.text)
        
        rank_list.append(rank_list)
        string_list.append(companyNameString.text)
    
getCompanyName()

# 下一頁
next_btn = driver.find_element_by_xpath("//input[@name='submit1' and @value='下一頁']")


for i in range(1,int(max_count/5)):
    print("------")
    next_btn = driver.find_element_by_xpath("//input[@name='submit1' and @value='下一頁']")
    next_btn.click()
    time.sleep(1)
    getCompanyName()




print("抓取完畢")

#driver關閉
driver.stop_client()
driver.quit()


"""
以下匯出功能
"""

# print("匯出excel開始")

# company_dict = {
#                 "rank": rank_list,
#                 "companyName": string_list
# }

# company_df = pd.DataFrame(company_dict)


# # Create a Pandas Excel writer using XlsxWriter as the engine.
# writer = pd.ExcelWriter('2000_company.xlsx', engine='xlsxwriter')

# # Convert the dataframe to an XlsxWriter Excel object.
# company_df.to_excel(writer, sheet_name='Sheet1')

# # Close the Pandas Excel writer and output the Excel file.
# writer.save()

# print("匯出完成")

