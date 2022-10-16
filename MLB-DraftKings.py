from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time


wait_time = 10
FirefoxDriver = "./geckodriver"
driver = webdriver.Firefox(executable_path=FirefoxDriver)
print("Getting DraftKings website")
url = "https://sportsbook.draftkings.com/leagues/baseball/88670847"
driver.get(url)
time.sleep(2)
Player_Stats = ['Strikeouts']
master_list = []
ids = []

box = driver.find_element(By.CLASS_NAME, "playspan-footer")

print(box)

print("Fetching Games ...")
games = driver.find_element(By.CLASS_NAME,'sportsbook-offer-category-card')
links = games.find_elements(By.TAG_NAME,'a')
for link in links:
    try:
        url = link.get_attribute('href')
        url = url.strip('?sgpmode=true')
        if 'baseball' not in url and url not in ids:
            ids.append(url)
            print(url)
    except:
        pass

print("Found {} games".format(len(ids)))

for i in range(0,4):
    print("Extracting Game {}".format(i+1))
    driver.get(ids[i])
    time.sleep(2)
    try:
        tab_list = driver.find_element(By.CLASS_NAME,'sportsbook-tabbed-subheader__tabs')
    except:
        continue
    tabs = tab_list.find_elements(By.XPATH,"./child::*")
    for j in range(0, len(tabs)):
        tab = tabs[j]
        if tab.text.lower() == 'pitcher props':
            driver.execute_script("arguments[0].click();", tab)
            time.sleep(2)
            table = driver.find_element(By.CSS_SELECTOR,'.sportsbook-responsive-card-container__card.selected')
            games_table = table.find_element(By.XPATH,'./div[2]')
            games = games_table.find_elements(By.XPATH,"./child::*")
            for game in games:
                info = game.text.split('\n')
                if info[0] in Player_Stats:
                    print("-------------- {} --------------".format(info[0]))
                    cnt = 5
                    while cnt+7 <= len(info):
                        item_dict = {}
                        name = info[cnt]
                        total = info[cnt+5]
                        under_odds = info[cnt+6]
                        item_dict['Player Name'] = name
                        item_dict['Total'] = total[1:]
                        item_dict['Under Odds'] = under_odds
                        item_dict['Over Odds'] = info[cnt+3]
                        item_dict['Stat'] = info[0]
                        master_list.append(item_dict)
                        print(item_dict)
                        cnt += 7
        tab_list = driver.find_element(By.CLASS_NAME,'sportsbook-tabbed-subheader__tabs')
        tabs = tab_list.find_elements(By.XPATH,"./child::*")


path = '/Users/ryanmccarroll/Google Drive/PyCharm Output/mlb - data.xlsx'
writer = pd.ExcelWriter(path, engine = 'openpyxl', mode='a')
book = load_workbook(path)
writer.book = book
try:
    del book['Draftkings Raw']
except:
    pass
result = pd.DataFrame(master_list)
result = result[['Player Name', 'Stat', 'Total', 'Under Odds', 'Over Odds']]
result.to_excel(writer, sheet_name='Draftkings Raw')
writer.save()
writer.close()
print("Draftkings is DONE!")
driver.close()