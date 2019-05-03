#import packages
import getpass
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time

#set variables
warrantyURL = 'https://pcsupport.lenovo.com/us/en/warrantylookup'
assetURL = 'https://fcnhelpdesk/helpdesk/WebObjects/Helpdesk.woa'
officeScan = 'https://fcntrend.fcn.net:4343/officescan/console/html/cgi/cgiChkMasterPwd.exe?id=0011'
usernameXpath = '//*[@id="userName"]'
passwordXpath = '//*[@id="password"]'
loginXpath = '//*[@id="dialog"]/div[1]/a/div/div[2]'
assetXpath = '//*[@id="navigation"]/div[1]/div[4]/a'
newAssetXpath = '//*[@id="ListHeaderDiv"]/table/tbody/tr/td/div[1]/table/tbody/tr/td[1]/a/div/div[2]'
assetDetails = '//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/a/table/tbody/tr/td'
saveXpath = '//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[14]/td/table/tbody/tr/td[2]/div/a[3]/div/div[2]'
assetNoXpath = '//*[@id="assetNumberField"]'
assetTypeXpath = '//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[3]/td[2]/div[1]/select'
locationXpath = '//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[5]/td[2]/select'
departmentXpath = ''
chromedriver = 'C:/Scripts/chromedriver/chromedriver.exe'
setPurchaseDate = '//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[2]/a/div'
assets = '//*[@id="navigation"]/div[1]/div[4]/a/div'
assetName = '//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[3]/td/form/table/tbody/tr/td/table/tbody/tr[1]/td[2]/input'
search = '//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[3]/td/form/table/tbody/tr/td/table/tbody/tr[5]/td/table/tbody/tr/td/div[1]/a[2]/div/div[2]'
asset = '//*[@id="ListUpdateDiv"]/table/tbody/tr[2]/td[6]'
editAsset = '//*[@id="headerItems"]/tbody/tr/td[2]/table/tbody/tr/td/input'
name = '//*[@id="assetNumberField"]'
submit = '//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[14]/td/table/tbody/tr/td[2]/div/a[3]/div/div[2]'

#get warranty info on asset
def get_warranty_date(serialNum):
	driver = webdriver.Chrome(chromedriver)
	driver.get(warrantyURL)
	driver.find_element_by_xpath('//*[@id="input_sn"]').send_keys(serialNum)
	driver.find_element_by_xpath('//*[@id="search-section"]/section/div/div/section/div[1]/button').click()
	time.sleep(5)
	warranty = driver.find_element_by_xpath('//*[@id="W-Warranties"]/section/div[1]/div/div[2]/ul/li[1]/div')
	return warranty.text

#format the date into the correct xx/xx/xxxx format
def date_format_laptop(serialNum):
	formatThis = get_warranty_date(serialNum)
	formatted = formatThis.split('-')
	formatted[0] = int(formatted[0]) - 4
	formatted[0] = str(formatted[0])
	formatted.append(formatted[0])
	formatted.remove(formatted[0])
	date = '/'.join(formatted)
	return date

def date_format_desktop(serialNum):
	formatThis = get_warranty_date(serialNum)
	formatted = formatThis.split('-')
	formatted[0] = int(formatted[0]) - 5
	formatted[0] = str(formatted[0])
	formatted.append(formatted[0])
	formatted.remove(formatted[0])
	date = '/'.join(formatted)
	return date

def get_loaction_and_department():
	location = str(input('Enter the number corresponding with the location of the device. 0 - Admin, 1 - Bellingham Bay Family Medicine, 2 - Birch Bay Family Medicine, 3 - Christian Health Care Center, 4 - Everson Family Medicine, 5 - Family Health Associates, 6 - Ferndale Family Medical Center, 7 - Inpatient Services, 8 - Island Family Physicians, 9 - Lynden Family Medicine, 10 - Medical Testing Center, 11 - North Cascade Family Physicians, 12 - North Sound Family Medicine , 13 - Squalicum Family Medicine, 14 - Urgent Care Center, 15 - Whatcom Family Medicine :    '))
	department = str(input('Enter the department. 0 - AACW, 1 - Admin, 2 - BBFM, 3 - Behavioral Health, 4 - BellB, 5 - Care Support Team, 6 - CHCC, 7 - Clinical, 8 - EFM, 9 - EMR, 10 - FCN-IPS, 11 - FFMC, 12 - FHA, 13 - Finance, 14 - HR, 15 - IFP, 16 - IT, 17 - Lab, 18 - LFM, 19 - MTC-UCC, 20 - NCFP, 21 - NSFM, 22 - PAD, 23 - Servers, 24 - SFM, 25 - SNF, 26 - WFM :    '))
	return department, location

def headless_chrome():
	options = Options()
	options.add_argument('--headless')
	options.add_argument('--disable-gpu')
	driver = webdriver.Chrome(chromedriver, chrome_options=options)
	return driver
	
#create asset
def new_asset(assetNum):
	model = input('Enter the model ')	
	serialNum = input('Enter the serial number ')
	location = input('Enter the full name of the location that this asset is going to: ')
	department = input('Enter the department: ')
	if 'L' in assetNum:
		date = date_format_laptop(serialNum)
	elif 'D' in assetNum:
		date = date_format_desktop(serialNum)
	driver = headless_chrome()
	driver.get(assetURL)
	# login done here
	driver.find_element_by_xpath(usernameXpath).send_keys(username)
	driver.find_element_by_xpath(passwordXpath).send_keys(password)
	driver.find_element_by_xpath(loginXpath).click()
	# navigating to new asset setup
	try:
		driver.find_element_by_xpath(assetXpath).click()
		driver.find_element_by_xpath(newAssetXpath).click()
		#Assigning asset number, type, locatio, department, and status
		driver.find_element_by_xpath(assetNoXpath).send_keys(assetNum)
		if 'L' in assetNum:
			select = Select(driver.find_element_by_xpath(assetTypeXpath))
			select.select_by_value('16')
		elif 'D' in assetNum:
			select = Select(driver.find_element_by_xpath(assetTypeXpath))
			select.select_by_value('7')
		select = Select(driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[4]/td[2]/div[1]/select'))
		select.select_by_visible_text(model)
		select = Select(driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[5]/td[2]/select'))
		select.select_by_visible_text(location)
		time.sleep(1)
		select = Select(driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[6]/td[2]/select'))
		select.select_by_visible_text(department)
		select = Select(driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[8]/td[2]/div[1]/select'))
		select.select_by_value('2')
		#changing to asset details page
		driver.find_element_by_xpath(assetDetails).click()
		#entering serial number of asset, setting purchase date with formatted date, and saving asset
		driver.find_element_by_xpath('//*[@id="TabPanelUpdateContainer"]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input').send_keys(serialNum)
		driver.find_element_by_xpath(setPurchaseDate).click()
		driver.find_element_by_xpath('//*[@id="date_7_31_0_0_0_0_2_1_0_4_3_1_9_3_1_0_1_3_1_3_0_1_1_16_1_3_1_1"]').send_keys(date)	
		driver.find_element_by_xpath('//*[@id="date_7_31_0_0_0_0_2_1_0_4_3_1_9_3_1_0_1_3_1_3_0_1_1_24_3_1"]').click()
		driver.find_element_by_xpath(saveXpath).click()
	except:
		driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[12]/td/table/tbody/tr/td/div/a[1]/div/div[2]')
	driver.close()
	print('asset ' + assetNum + ' has been added to inventory and status set to deployed.')

#retire asset
def retire_asset(assetNum):
	driver = headless_chrome()
	driver.get(assetURL)
	driver.find_element_by_xpath(usernameXpath).send_keys(username)
	driver.find_element_by_xpath(passwordXpath).send_keys(password)
	driver.find_element_by_xpath(loginXpath).click()
	time.sleep(1)
	driver.find_element_by_xpath(assets).click()
	driver.find_element_by_xpath(assetName).send_keys(assetNum)
	driver.find_element_by_xpath(search).click()
	time.sleep(1)
	driver.find_element_by_xpath(asset).click()
	driver.find_element_by_xpath(editAsset).click()
	driver.find_element_by_xpath(name).clear()
	driver.find_element_by_xpath(name).send_keys(assetNum + ' - OLD')
	select = Select(driver.find_element_by_xpath('//*[@id="AssetBasicsPanelDiv"]/table/tbody/tr[8]/td[2]/div[1]/select'))
	select.select_by_value('3')
	time.sleep(2)
	driver.find_element_by_xpath(submit).click()
	driver.close()
	print('asset ' + assetNum + ' has been rename and status set to retired.')

#user input
answer = ''
#Program here
while answer != 'new' or answer != 'retire' or answer != 'exit':
	answer = input('New asset or Retire asset (new/retire) or exit to quite: ')
	if answer == 'exit':
		print('exiting app.')
		break
	username = input('what is your username: ')
	password = getpass.getpass()
	if answer == 'new':
		assetNum = input('Enter the asset number: ')
		new_asset(assetNum)
	elif answer == 'retire':
		assetNum = input('Enter the asset number: ')
		retire_asset(assetNum)
