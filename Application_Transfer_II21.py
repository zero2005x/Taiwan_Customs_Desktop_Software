from selenium import webdriver



from selenium.webdriver.support.ui import WebDriverWait



from selenium.webdriver.support.ui import Select




import time


import pyautogui




class II21():


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

		commit = self.driver.find_element_by_id("main_searchButton5")

		while commit == '':

			time.sleep(0.5)

			commit = self.driver.find_element_by_id("main_searchButton5")

		commit.click()



	def pressing_exchanging_button_F3(self):

		commit = self.driver.find_element_by_id("main_transfer")

		while commit == '':

			time.sleep(0.5)

			commit = self.driver.find_element_by_id("main_transfer")

		commit.click()

	def pressing_cleaning_button_SF1(self):

		#main_resetButton5
		commit = self.driver.find_element_by_id("main_resetButton5")

		while commit == '':

			time.sleep(0.5)

			commit = self.driver.find_element_by_id("main_resetButton5")

		commit.click()


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



		# 打啟動selenium，務必確認driver檔案跟python檔案要在同個資料夾中

		driver = webdriver.Chrome(options=options)

		self.driver = driver


		self.driver.get("http://aci.customs.gov.tw/portal/Login_main")


		self.application_number = ""
		

		ID = input('請輸入帳號')

		

		Password = input('請輸入密碼')



		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[1]/td[2]/font/input[1]').send_keys(ID)




		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[2]/td[2]/input').send_keys(Password)

		



		self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td[3]/div/div[1]/table/tbody/tr[4]/td/input[1]').click()




		self.driver.get("http://aci.customs.gov.tw/APIM/II21?opener=true&?cust_Cd=AW")



		time.sleep(1)


if __name__=='__main__':

	

	First = II21()
  
  #all the application number must be place the first column in "Sheet1'.

	wb = load_workbook('applicatoinNumberList.xlsx')

	a_sheet = wb.get_sheet_by_name('Sheet1')

	time.sleep(5)

	#Filling your employer ID right now, or your program will fail.
	for index in range(1, 456):

		time.sleep(1)

		b4_too = a_sheet.cell(row=index, column=1)

		First.application_number = str(str(b4_too.value.replace("/", "").replace(" ", "")))

		print(First.application_number)

		First.filling_application_number()


		First.pressing_searching_button_F6()

		time.sleep(1)
		executable_javascript_commend = "return $('#statusMsg').val();"	
		Msg = str(First.driver.execute_script(executable_javascript_commend))
		if Msg != "[查無資料]":

			ckeckbox = "/html/body/div[2]/table/tbody/tr/td/form[1]/table[2]/tbody/tr/td/div/div[3]/div[3]/div/table/tbody/tr[2]/td[2]/input"

			First.driver.find_elements_by_id("checkBox1")[0].click()


			First.pressing_exchanging_button_F3()

			time.sleep(2)

			pyautogui.press('enter')

			time.sleep(2)
