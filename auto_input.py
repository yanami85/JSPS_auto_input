import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import settings

def add_tax(x):
    if x.is_tax =="税抜":
        return x.value*1.1
    else:
        return x.value

def convert_date_to_flag(x):
    this_year = settings.THIS_YEAR
    date_to_flag_map = {
        f"{this_year}-04": 1, f"{this_year}-05": 2, f"{this_year}-06": 3,
        f"{this_year}-07": 4, f"{this_year}-08": 5, f"{this_year}-09": 6,
        f"{this_year}-10": 7, f"{this_year}-11": 8, f"{this_year}-12": 9,
        f"{this_year + 1}-01": 10, f"{this_year + 1}-02": 11, f"{this_year + 1}-03": 12
    }
    return date_to_flag_map[x.date]

def determine_receipt(row):
    if row.determine_receipt == "あり":
        return "etr_needOrNotReceipt_01"
    else:
        return "etr_needOrNotReceipt_02"

def assign_expenditure(row):
    expenditure_map = {
        "学会関係経費": 1,
        "各種研究集会等への参加費": 2,
        "学術調査にかかる経費": 3,
        "自宅での研究に必要な経費": 4,
        "所属・関連機関への交通費": 5
    }
    return expenditure_map[row.expenditure]


def enter_forum(row):
    try:
        wait.until(EC.presence_of_element_located((By.ID,'etr_group')))
    except TimeoutException:
        wait.until(EC.presence_of_element_located((By.ID,'goto_F003_edit')))
        driver.find_element(By.ID, "goto_F003_edit").click()
        wait.until(EC.presence_of_element_located((By.ID,'etr_group')))
    dropdown_expense = driver.find_element(By.ID,'etr_group')
    select_expense = Select(dropdown_expense)
    select_expense.select_by_index(row.expenditure)
    
    wait.until(EC.presence_of_element_located((By.ID,'etr_dateOfComp')))
    dropdown_date = driver.find_element(By.ID,'etr_dateOfComp')
    select_date = Select(dropdown_date)
    select_date.select_by_index(row.date)
    
    # wait.until(EC.element_to_be_clickable((By.ID, str(row.determine_receipt))))
    # driver.find_element(By.ID,str(row.determine_receipt)).click()
    
    wait.until(EC.presence_of_element_located((By.NAME,"etr_itemName")))
    driver.find_element(By.NAME,"etr_itemName").send_keys(row.item)
    
    wait.until(EC.presence_of_element_located((By.NAME,"etr_billAmount")))
    driver.find_element(By.NAME,"etr_billAmount").send_keys(row.value)

    wait.until(EC.presence_of_element_located((By.NAME,"etr_remark")))
    driver.find_element(By.NAME,"etr_remark").send_keys(row.remark)

    wait.until(EC.element_to_be_clickable((By.ID,"add_to_basetable")))
    driver.find_element(By.ID,"add_to_basetable").click()

if __name__ == '__main__':
    df = pd.read_excel(settings.EXCEL_SHEET_PATH, usecols=[0,1,2,3,4,5,6,7]).dropna(how='all').fillna('')

    # seleniumで入力するためのidやキーを纏めたデータを作成する
    df.columns = ["date", "item", "value", "is_tax","number", "expenditure", "determine_receipt", "remark"]
    df['date'] = df['date'].dt.strftime("%Y-%m")
    df['value'] = df.apply(lambda x: add_tax(x), axis = 1)*df["number"]
    df["date"] = df.apply(lambda x: convert_date_to_flag(x), axis = 1)
    df["determine_receipt"] = df.apply(lambda x: determine_receipt(x), axis = 1)
    df["expenditure"] = df.apply(lambda x: assign_expenditure(x), axis = 1)
    df_l = [list(row) for row in df.itertuples()]
    df = df[["date", "item", "expenditure", "determine_receipt", "value", "remark"]]
    df["value"] = pd.Series(df["value"], dtype = 'int64')

    driver = webdriver.Chrome(executable_path=settings.DRIVER_PATH)

    J_SYSTEM_URL = rf"https://tyousa.jsps.go.jp/stu{str(settings.THIS_YEAR)[-2:]}/"
    J_SYSTEM_ID = settings.J_SYSTEM_ID
    J_SYSTEM_PASS = settings.J_SYSTEM_PASS

    driver.get(J_SYSTEM_URL)
    wait = WebDriverWait(driver, 5)
    
    wait.until(EC.presence_of_element_located((By.NAME, 'login_id')))
    driver.find_element(By.NAME, 'login_id').send_keys(J_SYSTEM_ID)
    driver.find_element(By.NAME,"login_psw").send_keys(J_SYSTEM_PASS)
    driver.find_element(By.NAME,"lg").click()
    
    wait.until(EC.presence_of_element_located((By.ID, "li_of_F003")))
    driver.find_element(By.ID, "li_of_F003").click()

    for row in df.itertuples():
        enter_forum(row)

    driver.find_element(By.ID,"additional_function_on_submitting").click()
    alert = driver.switch_to.alert
    alert.accept()
