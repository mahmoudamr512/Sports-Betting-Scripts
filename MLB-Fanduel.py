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
print("Getting Fanduel website")
url = "https://pa.sportsbook.fanduel.com/navigation/mlb"
driver.get(url)
time.sleep(4)
actions = ActionChains(driver)
games_pos = driver.find_elements(By.XPATH,'//*[@style="flex-direction: column; overflow: hidden auto; display: flex; min-width: 0px;"]')

if 'TOTAL' in games_pos[0].text:
    games_pos = games_pos[0]
else:
    games_pos = games_pos[1]

games = games_pos.find_elements("xpath","./child::*")
Stats = ['Player Strikeouts']

Player_props_tabs = ['Popular']

master_list = []
links = []


def scroll(element):
    desired_y = (element.size['height'] / 2) + element.location['y']
    window_h = driver.execute_script('return window.innerHeight')
    window_y = driver.execute_script('return window.pageYOffset')
    current_y = (window_h / 2) + window_y
    return desired_y - current_y


for i in range(0, len(games)):
    try:
        links.append(games[i].find_element(By.TAG_NAME,'a').get_attribute('href'))
    except:
        pass


print("Found {} Games".format(len(links)))
stop = 0
for j in range(0,1):
    if stop:
        break
    print("Extracting Game {}".format(j+1))
    driver.get(links[j])
    time.sleep(3)
    stats = driver.find_elements(By.XPATH,'//*[@style="flex-direction: column; overflow: hidden auto; display: flex; min-width: 0px;"]')
    if 'All Sports' in stats[0].text:
        stats = stats[1]
    else:
        stats = stats[0]
    stats = stats.find_elements(By.XPATH,"./child::*")

    for i in range(0, len(stats)):
        stat = stats[i].text.split('\n')[0]
        if stat in Stats:
            print("-------------- {} --------------".format(stat))
            button = stats[i].find_element(By.XPATH,'.//div[@role="button"]')
            scroll_y_by = scroll(button)
            driver.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by)
            if len(stats[i].text.split('\n')) < 2:
                button.click()
            time.sleep(0.5)
            try:
                show_more = driver.find_element(By.XPATH,'//span[contains(text(), "' + 'Show more' + '")]')
                driver.execute_script("window.scrollBy(0, arguments[0]);", scroll(show_more))
                show_more.click()
                time.sleep(0.5)
            except:
                pass

            stats = driver.find_elements(By.XPATH,
                '//*[@style="flex-direction: column; overflow: hidden auto; display: flex; min-width: 0px;"]')
            if 'All Sports' in stats[0].text:
                stats = stats[1]
            else:
                stats = stats[0]
            stats = stats.find_elements(By.XPATH,"./child::*")

            matches = stats[i].text.split('\n')
            cnt = 3
            while cnt+5 <= len(matches):
                item_dict = {}
                name = matches[cnt]
                total = matches[cnt+1].split(' ')[1]
                under_odds = matches[cnt+4]
                over_odds = matches[cnt+2]
                item_dict['Player Name'] = name
                item_dict['Total'] = total
                item_dict['Stat'] = stat
                item_dict['Under Odds'] = under_odds
                item_dict['Over Odds'] = over_odds
                master_list.append(item_dict)
                cnt+=5
                print(item_dict)
            try:
                button.click()
            except:
                pass

path = '/Users/ryanmccarroll/Google Drive/PyCharm Output/mlb - data.xlsx'
writer = pd.ExcelWriter(path, engine='openpyxl', mode='a')
book = load_workbook(path)
writer.book = book
try:
    del book['Fanduel Raw']
except:
    pass
result = pd.DataFrame(master_list)
result = result[['Player Name', 'Stat', 'Total', 'Under Odds', 'Over Odds']]
result.to_excel(writer, sheet_name='Fanduel Raw')
writer.save()
writer.close()
print("Fanduel is DONE!")
driver.close()
