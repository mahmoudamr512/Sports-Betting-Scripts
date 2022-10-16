from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pandas as pd
import sys
import os
import time

# Global Variables needed for the application
FireFoxDriverPath = "./geckodriver"
driver = webdriver.Firefox(executable_path=FireFoxDriverPath)
wait_time = 10
target_sport = "MLB"
master_list = []

# Disable
def blockPrint():
    sys.stdout = open(os.devnull, 'w')

# Restore
def enablePrint():
    sys.stdout = sys.__stdout__

def openMLB():
    """
    Function to BetOnline website and click on MLB and wait 5 seconds until it loaded. 
    """
    url = "https://bv2.digitalsportstech.com/betbuilder?sb=betonline"
    driver.get(url)

    sports = WebDriverWait(driver, wait_time).until(
        ec.visibility_of_element_located((By.CLASS_NAME, "ligues-slider__item.sportNames")))
    sports = driver.find_elements(By.CLASS_NAME,'ligues-slider__item.sportNames')
    for sport in sports:
        print(sport.text)
        if target_sport in sport.text:
            print(sport.text)
            sport.click()
            time.sleep(5)
            break

def main_item():
    """
    Function to return main item stat
    """
    items = WebDriverWait(driver, wait_time).until(
    ec.visibility_of_all_elements_located((By.CLASS_NAME, "main-stat__header")))

    items = driver.find_elements(By.CLASS_NAME,'main-stats__item.main-stat')
    items[0].find_element(By.CLASS_NAME,'main-stat__header').click()
    items[1].find_element(By.CLASS_NAME,'main-stat__header').click()
    time.sleep(0.5)
    items = driver.find_elements(By.CLASS_NAME,'main-stats__item.main-stat')
    item = items[1]
    time.sleep(0.5)
    return item


openMLB()
item = main_item()

stat = item.text
stat = stat.split(" ")[1]
stat = stat.strip('(').strip(')')
stat = stat.split(')')[0]
time.sleep(1)
games = item.find_elements(By.CLASS_NAME,'main-stat__content')
games_cards = item.find_elements(By.CLASS_NAME,'tiered_block__top__controls')

for i in range(0, len(games)):
    try:
        games_cards[i].click()
        time.sleep(1.5)
        matches = games[i].find_element(By.CLASS_NAME,'over-under-block')
        players = matches.find_elements(By.CLASS_NAME,'over-under-block__item')
        print("length of players: ", len(players))
        
        for player in players:
            item_dict = {}
            name = player.find_element(By.CLASS_NAME,'over-under-block__player-name').text
            over_under = player.find_elements(By.CLASS_NAME,'over-under-block__selector-value')
            totalDiv = player.find_element(By.CLASS_NAME, "over-under-block__selector-text")
            total = totalDiv.find_element(By.CLASS_NAME, "highlight-text-color").text
            over = over_under[0].text
            under = over_under[1].text
            item_dict['Player Name'] = name
            item_dict['Stat'] = stat
            item_dict['Total'] = total
            item_dict['Under Odds'] = under
            item_dict['Over Odds'] = over
            master_list.append(item_dict)
            print(item_dict)
        driver.execute_script("arguments[0].scrollIntoView(true);", games_cards[i])    
        driver.execute_script("arguments[0].click();", games_cards[i])  
        i+=1
    except Exception as ex:
        pass

path = '/Users/ryanmccarroll/Google Drive/PyCharm Output/mlb - data.xlsx'
writer = pd.ExcelWriter(path, engine='openpyxl', mode='a')
book = load_workbook(path)
writer.book = book
try:
    del book['BetOnline Raw']
except Exception as ex:
    pass

result = pd.DataFrame(master_list)
result = result[['Player Name', 'Stat', 'Total', 'Under Odds', 'Over Odds']]
result.to_excel(writer, sheet_name='BetOnline Raw')
writer.save()
writer.close()
print("BetOnline is DONE!")
driver.close()