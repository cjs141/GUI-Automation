import pyautogui, time
import config
import xlwings as xw

#from poco.drivers.osx.osxui_poco import OSXPoco

pyautogui.FAILSAFE = True

#The funtion of this program will be to engage the network preferences
# navigate to the desired network and login

wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
sheet = wb.sheets['Sheet1']
#print("logging into network...")
#time.sleep(0.5)
# wifilocation = pyautogui.locateOnScreen('wifitool.png')
# wifilocation
# #Box(left=1416, top=562, width=50, height=41)
# wifilocation[0]
# #1416
# wifilocation.left
# #1416
# wifipoint = pyautogui.center(wifilocation)
# wifipoint
# #Point(x=1441, y=582)
# wifipoint[0]
# #1441
# wifipoint.x
# #1441
# wifix, wifiy = wifipoint
print("clicking on the network settings...")
# pyautogui.click(wifix, wifiy)
wifiButtonXCoord = sheet.range('Q3').value
wifiButtonYCoord = sheet.range('R3').value
#pyautogui.click('wifitool.png')
#pyautogui.moveTo(1050, 10, duration=0.5)
#time.sleep(0.6)
pyautogui.click(wifiButtonXCoord, wifiButtonYCoord, 1, 1)
"""
time.sleep(0.3)

print("locating the desired network...")
pyautogui.press('down', presses = 1)
if networklocation == pyautogui.locateOnScreen('2813308004.png'):
	pyautogui.click(pyautogui.locateCenterOnScreen('2813308004.png'))
elif networklocation != pyautogui.locateOnScreen('2813308004.png'):
	while networklocation != pyautogui.locateOnScreen('2813308004.png'):
		pyautogui.press('down', presses = 1)
		if networklocation == pyautogui.locateOnScreen('2813308004.png'):
			pyautogui.click(pyautogui.locateCenterOnScreen('2813308004.png'))
			break

# Getting a Screenshot
    im = pyautogui.screenshot()

    #print(im.getpixel((0, 0)))  # returns RGB tuple of pixel
    #print(im.getpixel((50, 200)))

    # Analyzing the Screenshot
    im.getpixel((50, 200))  # identify RGB to match
    #print(pyautogui.pixelMatchesColor(50, 200, (64, 0, 193)))
    #print(pyautogui.pixelMatchesColor(50, 200, (65, 0, 193)))

    # Image Recognition
    if networklocation = pyautogui.locateOnScreen('2813308004.png'):  # must be pixel-perfect match
    #print(list(pyautogui.locateAllOnScreen('2813308004.png')))
    #print(pyautogui.center((217, 85, 75, 21)))  # xy of center of match area
    	pyautogui.click((networklocation.left + networklocation.width/2 , networklocation.top + networklocation.height/2))

#networklocation
#Box(left=1416, top=562, width=50, height=41)
#networklocation[0]
#1416
#networklocation.left
#1416
#pyautogui.click(pyautogui.center(networklocation))
#networkpoint
#Point(x=1441, y=582)
#networkpoint[0]
#1441
#networkpoint.x
#1441
# networkx, networky = networkpoint
# pyautogui.moveTo(networkx, networky, duration=0.30)
# time.sleep(1)
# pyautogui.click(networkx, networky)
# print("Desired network found")

# pyautogui.press('enter')
# network281location
# #network281location[0]
# network281location.left
# network281point = pyautogui.center(network281location)
# network281point
# #network281point[0]
# network281point.x
# network281x, network281y = network281point
# pyautogui.click('2813308004.png')

#time.sleep(1)
#pyautogui.click(pyautogui.locateOnScreen('2813308004.png'))
#time.sleep(3)
#pyautogui.click(pyautogui.locateCenterOnScreen('txst-bobcats.png'))

time.sleep(2) #allow time for the credential alert box to appear
data = input('Are credentials in the database? (Y/N): ')
if data == 'N':
	howMany = input('is there a username and password needed?(Y/N): ')
	if howMany == 'Y':
		username = print('Username: ')
		password = print('Password: ')
	elif howMany == 'N':
		password = print('Password: ')
elif data == 'Y':
	print("Now entering the in credentials...")
	pyautogui.write(config.HOMEWIFI_PASSWORD, interval=0.25)
pyautogui.press('enter')

#Switch focus from the parent page to the alert pop up

#Using Poco to interact with the windows and apps
# poco = OSXPoco({"appname": "Finder", "windowindex": 0})
# Find the first windows in 'Finder' application
# poco = OSXPoco({"appname_re": "[a][b][c]", "windowtitle": "dirname"}, ("192.168.1.10", 15004))
# Find the window named 'dirname' by regular expression remotely
# poco = OSXPoco({"bundleid": "com.apple.Finder", "windowtitle_re": "*.name"})
"""