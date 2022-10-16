from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from openpyxl import load_workbook
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
import pandas as pd
import re
import time

# Global Variable Definitions
wait_time = 10
FirefoxDriver = "./geckodriver"
driver = webdriver.Firefox(executable_path=FirefoxDriver)
stats = ["Pitcher Strikeouts (Over/Under)"]
master_list = []
ids = []
print("Opening MLB Page to scrape...")
url = "https://mtairycasino.foxbet.com/#/baseball/competitions/8661971"
driver.get(url)

# Handle Popup 
def close_popup():
    try:
        pop = driver.find_element(By.CLASS_NAME,'blue-button')
        pop.click()
        driver.get(url)
    except:
        pass

#Wait Page fully loaded
def load_page():
    loaded = WebDriverWait(driver, wait_time).until(
    ec.visibility_of_all_elements_located((By.CLASS_NAME, "event-schedule-additional-markets")))

#Getting All Games IDs
def get_all_games():
    print("Fetching All Games!")
    games = driver.find_elements(By.CLASS_NAME,'event-schedule-additional-markets')

    for game in games:
        link = game.find_element(By.TAG_NAME,'a').get_attribute('href')
        id = re.findall('\d+', link)
        ids.append(id[0])

    print("{} Games to lookup at".format(len(ids)))

#Check if tabs exist
def check_tabs():
        tabs = WebDriverWait(driver, wait_time).until(
        ec.presence_of_element_located((By.CSS_SELECTOR, ".nav.nav-pills.market-groups.horizontalMenu__scroller")))
        driver.execute_script("arguments[0].scrollIntoView(true);", tabs)  
        time.sleep(2)
        tabs =  driver.find_element(By.CSS_SELECTOR, ".nav.nav-pills.market-groups.horizontalMenu__scroller")
        return tabs

#Extract Player Props
def     extract_props(prop):
    # We found Player Props! Now, click on Prop and wait
    driver.execute_script("arguments[0].click();", prop)  
    time.sleep(0.5)
    table = driver.find_element(By.ID,'markets-view')
    table = table.find_element(By.XPATH,"./child::*")

    events = table.find_elements(By.XPATH, "./child::*")
    found = False
    # Now we had all events, we have to check for Strikeouts only!
    for k in range(0, len(events)):
        stat = events[k].text.split('\n')[0]
        if stat in stats:
            found = True
            driver.execute_script("arguments[0].scrollIntoView();", events[k])
            time.sleep(0.5)
            print("-------- {} --------".format(stat))
            match = events[k].find_element(By.CSS_SELECTOR, '.selectionBody.collapseToggle__content')
            names = match.find_elements(By.CSS_SELECTOR, '.price.grid-market--aggregated-title')
            over_under = match.find_elements(By.CLASS_NAME, 'button__bet__title')
            odds = match.find_elements(By.CLASS_NAME, 'button__bet__odds')
            for j in range(0, len(names)):
                name = names[j].text
                under_odds = odds[j*2+1].text 
                over_odds = odds[j*2].text 
                over_under[j*2].text.split(' ')[0]
                total = over_under[j*2].text.split(' ')[1]
                item_dict = {}
                item_dict['Player Name'] = name
                item_dict['Stat'] = stat.strip(' (Over/Under)')
                item_dict['Total'] = total
                item_dict['Under Odds'] = under_odds
                item_dict['Over Odds'] = over_odds
                print(item_dict)
                master_list.append(item_dict)

                match = events[k].find_element(By.CLASS_NAME,
                    'selectionBody.collapseToggle__content')
                names = match.find_elements(By.CLASS_NAME, 'price.grid-market--aggregated-title')
        # table = driver.find_element(By.ID, 'markets-view')
        # table = table.find_element(By.XPATH,"./child::*")
        # events = table.find_elements(By.XPATH, "./child::*")
    if not found:
        print("No Strikeouts found here!")

# Function Calls
close_popup()
load_page()
get_all_games()


# change the start and the end game
for i in range(0, 19):
    #Getting the game URL to open
    print("Extracting Game {} ----------------------------:-".format(i+1))

    try:
        url = "https://mtairycasino.foxbet.com/#/baseball/competitions/event/{}".format(ids[i])
        driver.get(url)
    except:
        #Game doesn't exist
        print("Game doesn't exist anymore!")
        time.sleep(1)
        continue

    #Now Check if tabs exist
    tabs = check_tabs()
    tabs_list = ""
    #NECESSARY! The LI gets stale after a couple of mins!!!!
    try: 
        tabs_list = tabs.find_elements(By.TAG_NAME,'li')
    except:
        tabs_list = tabs.find_elements(By.TAG_NAME,'li')
    # Previous code always checked for index 2! But it is changable! Now we have to iterate!
    propsFlag = False
    try:
        for prop in tabs_list:
            if prop.text == "Player Props":
                propsFlag = True
                extract_props(prop)
    except:
        tabs_list = tabs.find_elements(By.TAG_NAME,'li')
        for prop in tabs_list:
            if prop.text == "Player Props":
                propsFlag = True
                extract_props(prop)
    # No Data Handling
    if not propsFlag:
        print("No Data To Get here")
        time.sleep(1)
        continue

   

result = pd.DataFrame(master_list)
path = '/Users/ryanmccarroll/Google Drive/PyCharm Output/mlb - data.xlsx'
writer = pd.ExcelWriter(path, engine='openpyxl')
book = load_workbook(path)
writer.book = book
try:
    del book['FoxBet Raw']
except:
    pass
result = pd.DataFrame(master_list)
result = result[['Player Name', 'Stat', 'Total', 'Under Odds', 'Over Odds']]
result.to_excel(writer, sheet_name='FoxBet Raw')
writer.save()
writer.close()
print("MLB-FoxBet.xlsx Is Ready!")

driver.close()
