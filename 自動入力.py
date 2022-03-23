import pandas as pd
from selenium import webdriver
from time import sleep
import time
import datetime
from selenium.webdriver.support.select import Select

df = pd.read_excel("研究遂行経費の支出報告書用入力シート.xlsx").dropna()

df.columns = ["date", "item", "value", "is_tax","number", "expenditure", "is_receipt"]
df['date'] = df['date'].dt.strftime("%Y-%m")

def add_tax(x):
    if x.is_tax =="税抜":
        return x.value*1.1
    else:
        return x.value

def convert_date_to_flag(x):
    if x.date == "2021-04":
        return 1
    elif x.date == "2021-05":
        return 2
    elif x.date == "2021-06":
        return 3
    elif x.date == "2021-07":
        return 4
    elif x.date == "2021-08":
        return 5
    elif x.date == "2021-09":
        return 6
    elif x.date == "2021-10":
        return 7
    elif x.date == "2021-11":
        return 8
    elif x.date == "2021-12":
        return 9
    elif x.date == "2022-01":
        return 10
    elif x.date == "2022-02":
        return 11
    elif x.date == "2022-03":
        return 12
    
def is_receipt(x):
    if x.is_receipt == "あり":
        return "etr_needOrNotReceipt_01"
    else:
        return "etr_needOrNotReceipt_02"
    
def assign_expenditure(x):
    if x.expenditure == "学会関係経費":
        return 1
    elif x.expenditure == "各種研究集会等への参加費":
        return 2
    elif x.expenditure =="学術調査にかかる経費":
        return 3
    elif x.expenditure == "自宅での研究に必要な経費":
        return 4
    elif x.expenditure == "所属・関連機関への交通費":
        return 5

df['value'] = df.apply(lambda x: add_tax(x), axis = 1)*df["number"]
df["date"] = df.apply(lambda x: convert_date_to_flag(x), axis = 1)
df["is_receipt"] = df.apply(lambda x: is_receipt(x), axis = 1)
df["expenditure"] = df.apply(lambda x: assign_expenditure(x), axis = 1)

df_l = [list(row) for row in df.itertuples()]
df = df[["date", "item", "expenditure", "is_receipt", "value"]]
df["value"] = pd.Series(df["value"], dtype = 'int64')

driver = webdriver.Chrome(executable_path='./chromedriver.exe')

JSPS_URL = "https://area31.smp.ne.jp/area/servlet/area.MyPageBundle?MyPageID=629871a6_lalj9lenala0oare4"
JSPS_ID = "JSPS_ID"
JSPS_PASS = "JSPS_PASS"

J_SYSTEM_ID = "ID"
J_SYSTEM_PASS = "PASS"

driver.get(JSPS_URL)
sleep(1)
driver.find_element_by_name("SMPID").send_keys(JSPS_ID)
driver.find_element_by_name("SMPPASSWORD").send_keys(JSPS_PASS)
driver.find_element_by_class_name("btn").click()
driver.find_element_by_link_text('⑤ 報告書等提出システム（Jシステム）').click()
sleep(1)

driver.switch_to.window(driver.window_handles[1])
driver.find_element_by_name("login_id").send_keys(J_SYSTEM_ID)
driver.find_element_by_name("login_psw").send_keys(J_SYSTEM_PASS)
driver.find_element_by_name("lg").click()
sleep(1)

driver.find_element_by_id("li_of_F003").click()
sleep(1)

def enter_forum(row):
    dropdown_expense = driver.find_element_by_id('etr_group')
    select_expense = Select(dropdown_expense)
    select_expense.select_by_index(row.expenditure)
    
    dropdown_date = driver.find_element_by_id('etr_dateOfComp')
    select_date = Select(dropdown_date)
    select_date.select_by_index(row.date)
    
    driver.find_element_by_id(str(row.is_receipt)).click()
    
    driver.find_element_by_name("etr_itemName").send_keys(row.item)
    
    driver.find_element_by_name("etr_billAmount").send_keys(row.value)
    
    driver.find_element_by_id("add_to_basetable").click()
    sleep(2)


if driver.find_element_by_id("goto_F003_edit"):
    driver.find_element_by_id("goto_F003_edit").click()
    sleep(2)
    for row in df.itertuples():
        enter_forum(row)
else:
    sleep(2)
    for row in df.itertuples():
        enter_forum(row)

driver.find_element_by_id("additional_function_on_submitting").click()