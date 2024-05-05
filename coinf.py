import pandas as pd
from datetime import datetime, time
import time as tm
from selenium import webdriver
from selenium.webdriver.common.by import By

# Set up WebDriver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

def should_run():
    now = datetime.now().time()
    return ((time(5, 0) <= now <= time(8, 0)) or
            (time(11, 0) <= now <= time(12, 0)) or
            (time(18, 0) <= now <= time(19, 0)))

def fetch_data(driver):
    url = 'https://upbit.com/exchange'
    driver.get(url)
    driver.implicitly_wait(10)  # Adjust based on load time

    data = []
    for i in range(1, 11):
        coin_name_xpath = f'//*[@id="UpbitLayout"]/div[3]/div/section[2]/article/span[2]/div/div/div[1]/table/tbody/tr[{i}]/td[3]'
        volume_xpath = f'//*[@id="UpbitLayout"]/div[3]/div/section[2]/article/span[2]/div/div/div[1]/table/tbody/tr[{i}]/td[6]'
        coin_name = driver.find_element(By.XPATH, coin_name_xpath).text
        volume = driver.find_element(By.XPATH, volume_xpath).text
        data.append([coin_name, volume])
    return data

def save_to_excel(data, session_start):
    filename = datetime.now().strftime('%Y%m%d') + '.xlsx'
    try:
        existing_data = pd.read_excel(filename)
    except FileNotFoundError:
        existing_data = pd.DataFrame()

    new_data = pd.DataFrame(data, columns=[f'{session_start.strftime("%H:%M")} Coin', f'{session_start.strftime("%H:%M")} Volume'])
    result = pd.concat([existing_data, new_data], axis=1)
    result.to_excel(filename, index=False)

def main():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)

    while True:
        if should_run():
            session_start = datetime.now()
            data = fetch_data(driver)
            save_to_excel(data, session_start)
            next_run = session_start.replace(minute=15)
            sleep_time = (next_run - datetime.now()).seconds
        else:
            sleep_time = 60  # check every minute if it's time to run

        tm.sleep(sleep_time)

    driver.quit()

if __name__ == '__main__':
    main()
