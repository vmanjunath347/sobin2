import time
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
import pandas


def getUsersArray():
    wb = pandas.read_excel(io = 'Planectcode_Registration.xlsx',sheet_name = "File32-Planectcode_Registration",index_col = None,engine = 'openpyxl',)
    wb_list = wb.values.tolist()
    return wb_list

def operations(data):
    capabilities = DesiredCapabilities.CHROME
    capabilities["goog:loggingPrefs"] = {"performance": "ALL"}

    #for row in data:
    for q in range(1251,len(data)):
        row = data[q]
        username = row[1]
        #print(row[0]+" "+username)
        password = row[2]
        driver = webdriver.Chrome(executable_path='./chromedriver',desired_capabilities=capabilities)
        driver.maximize_window()
        driver.get("https://www.planetcode.in/")
        time.sleep(1)
        initialLoginButton = driver.find_element_by_xpath('//*[@id="responsive"]/li[4]/a')
        initialLoginButton.click()
        time.sleep(2)
        username_window = driver.find_element_by_xpath('//*[@id="Loginform"]/div[1]/div/input')
        password_window = driver.find_element_by_xpath('//*[@id="Loginform"]/div[2]/div/input')
        login_button = driver.find_element_by_xpath('//*[@id="Loginform"]/div[3]/div[2]/button')

        username_window.clear()
        username_window.send_keys(username)

        password_window.clear()
        password_window.send_keys(password)

        login_button.click()
        time.sleep(1)
        try:
            driver.find_element_by_xpath('/html/body/div[2]/div/div[3]/div/button').click()
        except:
            pass
        time.sleep(2)
        beginners_course =  driver.find_element_by_xpath('//*[@id="wrapper"]/div[2]/div/div[7]/div[1]/a/div')
        beginners_course.click()
        time.sleep(1)
        about_game_development = driver.find_element_by_xpath('//*[@id="wrapper"]/div[2]/div/div[5]/div/ul/li[1]/div')
        about_game_development.click()
        time.sleep(5)

        try:
            driver.find_element_by_xpath('/html/body/div[2]/div/div[2]/div[1]/button').click()
            time.sleep(0.5)
            driver.find_element_by_xpath('/html/body/div[2]/div/div[4]/div/button').click()
        except:
            pass

        time.sleep(5)
        
        iframe  = driver.find_element_by_xpath('//*[@id="myFrame1"]')
        driver.switch_to.frame(iframe)
        '''
        action = ActionChains(driver)
        video_wndow = driver.find_element_by_xpath('//*[@id="player"]/div[1]')
        action.move_to_element(video_wndow).perform()
        '''

        play_button = driver.find_element_by_xpath('//*[@id="player"]/div[7]/div[3]/button')
        play_button.click()
        time.sleep(110)
        driver.close
        
def main():
    
    data = getUsersArray()
    operations(data)


    #driver = login(driver)

if __name__ == "__main__":

    main()