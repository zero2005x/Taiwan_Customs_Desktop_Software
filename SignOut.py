from selenium import webdriver


from selenium.webdriver.support.ui import WebDriverWait



from selenium.webdriver.support.ui import Select


import time


import datetime


import os


import pyautogui

class SignOut():

	def __init__(self):
		options = webdriver.ChromeOptions()



		prefs = {

    		'profile.default_content_setting_values':

        		{

            		'notifications': 2

        		}

		}



		options.add_experimental_option('prefs', prefs)



		options.add_argument("disable-infobars")  



		

		driver = webdriver.Chrome(options=options)

		self.driver = driver


		self.driver.get("http://aci.customs.gov.tw/portal/Login_main")



		

		ID = input("請輸入帳號: \n")

		
		Password = input("請輸入密碼: \n")


		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[1]/td[2]/font/input[1]').send_keys(ID)




		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[2]/td[2]/input').send_keys(Password)


		

		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[4]/td/input[1]').click()



		time.sleep(1)





if __name__=='__main__':


	First = SignOut()	

	remain_time = input("請輸入倒數時間(秒數): ")

	time.sleep(int(remain_time))

	#點擊下班按鈕
	commit = First.driver.find_element_by_id("SignOut")
	commit.click()

	#doSignInOut(this.id)
  
	#執行下班程式
	First.driver.execute_script("doSignInOut(this.id);")
	time.sleep(5)
  
  #一分鐘後自動關機
	os.system("shutdown -s -t 60")
