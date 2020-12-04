from selenium import webdriver

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support.ui import Select

import time

class IG08():
	def application_number_check(self):
		if len(self.application_number) > 14 or len(self.application_number) < 12:
			return False
		else:
			return True


	def application_number_convert(self):
		if len(self.application_number) == 14:
			return self.application_number
		elif len(self.application_number) == 12:
			return self.application_number[0:2] + "  " + self.application_number[2:14]
		else:
			return ''


	def application_number_transformation(self):
		if self.application_number_check() == True:
			return self.application_number_convert()
		else:
			return ''


	def filling_application_number(self):

		application_number_transformation_finished = self.application_number_transformation() 

		search_input = self.driver.find_element_by_name("declNo1")
		search_input.clear()
		search_input.send_keys(application_number_transformation_finished[0:2])

		search_input = self.driver.find_element_by_name("declNo2")
		search_input.clear()
		search_input.send_keys(application_number_transformation_finished[2:4])

		search_input = self.driver.find_element_by_name("declNo3")
		search_input.clear()
		search_input.send_keys(application_number_transformation_finished[4:6])

		search_input = self.driver.find_element_by_name("declNo4")
		search_input.clear()
		search_input.send_keys(application_number_transformation_finished[6:9])

		search_input = self.driver.find_element_by_name("declNo5")
		search_input.clear()
		search_input.send_keys(application_number_transformation_finished[9:14])


	def pressing_searching_button_F6(self):
		commit = driver.find_element_by_id("master0_searchButton2")

		while commit == '':
			time.sleep(0.5)
			commit = driver.find_element_by_id("master0_searchButton2")

		commit.click()

	def checking_status_Message(self):

		executable_javascript_commend = "return $('#statusMsg').val();"	
		time.sleep(1)
		check_empty = ''
		while check_empty == '':
			check_empty = str(self.driver.execute_script(executable_javascript_commend))
			time.sleep(1)

		if status_message == "[稅單資料不存在！]":
			print("稅單資料不存在")
			return False
		elif status_message =="[押金核准新增錯誤:報單尚未完成分估複核]":
			print("status_message")
			return False
		else:
			return True
			


	def __init__(self):
		ID = input("輸入帳號:\n")
		Password = input("輸入密碼:\n")
		options = webdriver.ChromeOptions()

		prefs = {
    		'profile.default_content_setting_values':
        		{
            		'notifications': 2
        		}
		}

		options.add_experimental_option('prefs', prefs)

		options.add_argument("disable-infobars")  

		# 打啟動selenium，務必確認driver檔案跟python檔案要在同個資料夾中
		driver = webdriver.Chrome(options=options)


		driver.get("http://aci.customs.gov.tw/portal/Login_main")

		time.sleep(0.5)

		context = driver.find_element_by_css_selector('#userId')

		
		context.send_keys(ID)

		#輸入password

		context = driver.find_element_by_css_selector('#userPwd')
		#Password = "xup6xu;4Wu/6"
		context.send_keys(Password)

		commit = driver.find_element_by_css_selector('#loginByUserPwd')

		commit.click()

		driver.get("http://aci.customs.gov.tw/APIM/IE07?opener=true&?cust_Cd=AW")

		time.sleep(1)

		self.duty_Treatment_DEPOSIT_REASON_Talbe = {
  			"38": "07",
  			"61": "11",
  			"62": "10",
  			"63": "27",
  			"64": "99",
 		  	"65": "99",
  			"66": "17",
  			"67": "05",
  			"69": "05",
  			"71": "06",
  			"73": "09",
 			"74": "09",
  			"79": "99",
		}
		self.application_number = ""

if __name__=='__main__':
	First = IG08()
	finish_state = True

	while finish_state:
		First.application_number = input("請輸入報單號碼,輸入z結束\n")
		if First.application_number == "z":
			finish_state = False
			break
		First.driver.execute_script("sf1Clear();")
		First.filling_application_number()
		First.pressing_searching_button_F6()

		#確認報單狀態
		if checking_status_Message() != True:
			continue
#Deposit_Payment_Bill_Generator_IG08.py
