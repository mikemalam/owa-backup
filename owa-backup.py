#! /usr/bin/python3
# owa2pdf.py - Sauvegarde les emails d'une messagerie Outlook via un site web.

# import libraries.
import sys, time, re, datetime, logging, random, ftplib, zipfile, os, unidecode, pdfkit, pyautogui
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

# Logging options
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

def zipdir(path, ziph):
	# zip ine zipfile handle
	for root, dirs, files in os.walk(path):
		for f in files:
			ziph.write(os.path.join(root, f))

def get_firefox_driver(download_dir):
	mime_types = "application/pdf,application/vnd.adobe.xfdf,application/vnd.fdf,application/vnd.adobe.xdp+xml,text/html,text/plain,application/octet-stream,application/vnd.ms-powerpoint, application/msword, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, image/jpeg, image/png, application/zip, application/x-zip, application/download"
	fp = webdriver.FirefoxProfile()
	fp.set_preference("browser.download.folderList",2)
	fp.set_preference("browser.download.manager.showWhenStarting", False)
	fp.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
	fp.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
	fp.set_preference("pdfjs.disabled", True)
	fp.set_preference("browser.download.dir", download_dir)
	logging.info('The downloading directory is ' + download_dir)
	return webdriver.Firefox(firefox_profile=fp)

def cleanFileNames(SauvegardingName):
	path = os.path.join(os.getcwd(), SauvegardingName)
	for root, dirs, files in os.walk(path):
		for f in files:
			os.rename(os.path.join(root, f), os.path.join(root, unidecode.unidecode(f)))


# Main ? porogram ?
if __name__ == "__main__":
	# Retrieve user informations
	print('###\nYou will be asked some informations to proceed to a backup \
	from your Outlook Web App (owa) messagery.\n###\n')
	emailSaving = input('Your personal email: ')
	siteAdress = input('Site address: ')
	siteUsername = input('Username: ')
	sitePassword = input('Password: ')

	# Create the personalized name for differents files
	today = str(datetime.date.today())
	nameForSaving = 'SAUVEGARDE_' + siteUsername + '_' + today

	# Create directories to save files
	os.makedirs(nameForSaving, exist_ok=True)
	os.makedirs(os.path.join(nameForSaving, 'files'), exist_ok=True)

	# Connecting to the site with identification.
	browser = get_firefox_driver(os.path.join(os.getcwd(), nameForSaving, 'files'))
	logging.info('\nConnecting to ' + siteAdress + '...\n')
	browser.get(siteAdress)
	time.sleep(1) # To wait for the page redirection.
	lightOption = browser.find_element_by_id('chkBsc')
	lightOption.click()
	loginElement = browser.find_element_by_id('username')
	passwordElement = browser.find_element_by_id('password')
	loginElement.send_keys(siteUsername)
	passwordElement.send_keys(sitePassword)
	linkElem = browser.find_element_by_class_name('btn')
	linkElem.click() #Sign in to loginto the site.

	# Connect to last message.
	lastMessageTitle = browser.find_element_by_tag_name('h1')
	lastMessageTitle.click()

	# Connect to this specific message to not start from the beginning if you are searching for a bug. Uncomment th following line with adequat adress and comment the 2 lines before.
	# browser.get('https://TYPE_YOUR_ADRESS_HERE')

	#Navigate threw messages	
	#Began iteration
	string = u'''
<!DOCTYPE html>
<html>
	<head>
		<title>Backup</title>
	</head>
  	<body>
		<table>
	'''
	condition = True
	numberEmail = 0
	numberEmailFile = 0
	# Create a text file to save html code and crush if already exists.
	textFile = open(os.path.join(nameForSaving, 'emails.html'), 'w')
	while condition == True:
		try:
			# Find message
			try:
				titleMessage = browser.find_element_by_class_name('msgHd')
			except Exception as err:
				condition = False # If no message header it probably menans it was last message
				logging.warning('An exception happened in try line %s : %        s' % (sys.exc_info()[-1].tb_lineno, str(err)))
			logging.debug(titleMessage.text)
			messageDisplay = browser.find_elements_by_class_name('w100')
			textToWash = messageDisplay[1].get_attribute('innerHTML')
			# TODO: Download each attached file
			try:
				linkAttachedFiles = browser.find_elements_by_id('lnkAtmt')
				for uniqueFile in linkAttachedFiles:
					try:
						regexTextLink = re.compile(r'\.[a-zA-Z]*')
						regexTextLinkResult = regexTextLink.search(uniqueFile.text)
						logging.debug('Regex search in ' + uniqueFile.text +  ': ' + regexTextLinkResult.group().lower())
						listAvoidedExtensions = ['.html', '.htm', '.vcf']
						if regexTextLinkResult.group().lower() in listAvoidedExtensions:
							pass
						else:
							uniqueFile.click()
					except Exception as err:
						logging.warning('An exception happened in try line %s : %    s' % (sys.exc_info()[-1].tb_lineno, str(err)))
						pass
					numberEmailFile += 1
					rNumber = random.randint(3,4) # waiting to be download and random time to not be spammed
					time.sleep(rNumber)
				
			except Exception as err:
				logging.warning('An exception happened in try line %s : %s' % (sys.exc_info()[-1].tb_lineno, str(err)))
			# Remove unnecessary source Html code
			exRegex = re.compile(r'<table class="tbhd" cellspacing="0".*crvtprt.gif" alt=""></td></tr></tbody></table>', re.DOTALL)
			trashCodeHtml = exRegex.search(textToWash)
			textWashed = textToWash.replace(trashCodeHtml.group(), "")
			string += unidecode.unidecode(textWashed)
			textFile.write(string)
			numberEmail += 1
			if numberEmail % 10 == 0:
				logging.info(str(numberEmail) + ' emails and ' + str(numberEmailFile) + ' files attached have have been downloaded so far...')
				for i in range(2): # Prevent computer to go to sleeping mode
					pyautogui.moveTo(100, 100, duration=1)
					pyautogui.moveTo(200, 100, duration=1)
			string = ''
			# Find link to newt message
			try :
				nextLink = browser.find_element_by_id('lnkHdrnext')
				nextLink.click()
				rNumber = random.randint(3,4)
				time.sleep(rNumber)
			except Exception as err:
				condition = False # If no next link it means it is the last message
				logging.warning('An exception happened in try line %s : %    s' % (sys.exc_info()[-1].tb_lineno, str(err)))
		except Exception as err:
			logging.warning('An exception happened in try line %s : %    s' % (sys.exc_info()[-1].tb_lineno, str(err)))
	#pisa.showLogging()
	#open output file for writing (truncated binary)
	textFile.write(u'''
		</table>
	</body>
</html>
	''')
	textFile.close()
	try:
		pdfkit.from_file(os.path.join(nameForSaving, 'emails.html'),os.path.join(nameForSaving, 'emails.pdf'))

	except Exception as err:
		logging.warning('An exception happened in try line %s : %    s' % (sys.exc_info()[-1].tb_lineno, str(err)))

	# Rename files (downloaded) properly.
	cleanFileNames(nameForSaving)

	# Zip Directory
	zipFile = zipfile.ZipFile(nameForSaving + '.zip', 'w')
	time.sleep(10) # waiting last file to download
	zipdir(nameForSaving, zipFile)
	zipFile.close()

	# Upload file on dl.free.fr
	session = ftplib.FTP('dl.free.fr')
	time.sleep(5)
	session.login(emailSaving, 'password')
	fileToSend = open(nameForSaving + '.zip', 'rb')
	logging.info('Upload begins...')
	session.storbinary('STOR ' + nameForSaving + '.zip', fileToSend)
	fileToSend.close()
	session.quit()
	logging.info('Your backup ' + nameForSaving + '.zip has been uploaded. You can \
	check your messagery ' + emailSaving +'./n And it represents ' + str(numberEmail) + ' emails and ' + str(numberEmailFile) + ' files.')

