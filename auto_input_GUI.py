import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import PySimpleGUI as sg
import auto_input


layout =[
    [sg.Text("ファイル指定", size=(25,1)), sg.InputText(key='EXCEL_SHEET_PATH', enable_events = True, size=(45,1)), sg.FileBrowse('参照',key='EXCEL_SHEET_PATH', file_types=(('Excel file', '*.xlsx'), ('Excel file', '*.xls')))],
    [sg.Text("J-Sytem ID", size=(25,1)), sg.InputText(key='J_SYSTEM_ID', size=(45,1))],
    [sg.Text("J-Sytem PASSWORD", size=(25,1)), sg.InputText(key='J_SYSTEM_PASS', size=(45,1))],
    [sg.Submit(button_text="やめる"), sg.Submit(button_text="次へ")]
    ]

window = sg.Window('Test', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "やめる":
        values["cancel_flag"] = ""
        break
    if event == "次へ":
        print(event, values)
        break
window.close()

if any(string == "" for string in values.values()):
    exit()

if __name__ == '__main__':
    df = pd.read_excel(values['EXCEL_SHEET_PATH'], usecols=[0,1,2,3,4,5,6]).dropna()

    df.columns = ["date", "item", "value", "is_tax","number", "expenditure", "is_receipt"]
    df['date'] = df['date'].dt.strftime("%Y-%m")

    df['value'] = df.apply(lambda x: auto_input.add_tax(x), axis = 1)*df["number"]
    df["date"] = df.apply(lambda x: auto_input.convert_date_to_flag(x), axis = 1)
    df["is_receipt"] = df.apply(lambda x: auto_input.determine_receipt(x), axis = 1)
    df["expenditure"] = df.apply(lambda x: auto_input.assign_expenditure(x), axis = 1)

    df_l = [list(row) for row in df.itertuples()]
    df = df[["date", "item", "expenditure", "is_receipt", "value"]]
    df["value"] = pd.Series(df["value"], dtype = 'int64')

    driver = webdriver.Chrome(executable_path=r"./chromedriver.exe")

    J_SYSTEM_URL = auto_input.J_SYSTEM_URL
    J_SYSTEM_ID = values["J_SYSTEM_ID"]
    J_SYSTEM_PASS = values["J_SYSTEM_PASS"]

    driver.get(J_SYSTEM_URL)
    wait = WebDriverWait(driver, 5)
    
    wait.until(EC.presence_of_element_located((By.NAME, 'login_id')))
    driver.find_element(By.NAME, 'login_id').send_keys(J_SYSTEM_ID)
    driver.find_element(By.NAME,"login_psw").send_keys(J_SYSTEM_PASS)
    driver.find_element(By.NAME,"lg").click()
    
    wait.until(EC.presence_of_element_located((By.ID, "li_of_F003")))
    driver.find_element(By.ID, "li_of_F003").click()

    for row in df.itertuples():
        auto_input.enter_forum(row)

    driver.find_element(By.ID,"additional_function_on_submitting").click()
    alert = driver.switch_to.alert
    alert.accept()