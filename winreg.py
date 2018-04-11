import os
import sys
from winreg import *

def Windows_User_Id():
	cut = 0
	User_Num = 0
	try :
		while True:
			EnumKey(HKEY_USERS, cut)
			cut += 1
	except WindowsError:
		User_Num += cut-3
		return User_Num

def Windows_User_Id_Set():
	Set = Windows_User_Id()
	result = EnumKey(HKEY_USERS, Set)
	return result

def Windows_info():
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
    print ('OS_Name : ' + QueryValueEx(Wininfo_Key, 'ProductName')[0])
    print ('OS_Root_Directory : ' + QueryValueEx(Wininfo_Key, 'SystemRoot')[0])
    print ('OS_install_Date : ' + str(QueryValueEx(Wininfo_Key, 'installDate')[0]))
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName")
    print ('ComputerName : ' + QueryValueEx(Wininfo_Key, 'ComputerName')[0])
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\Windows")
    print ('Last_ShutDown_Time : ' + str(QueryValueEx(Wininfo_Key, 'ShutdownTime')[0]))

def Fonts():
	Fonts_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts")
	try :
		cut = 0
		while True:
			name, value, type  = EnumValue(Fonts_Key, cut)
			print (name.replace("(TrueType)", "") + ': ' + value)
			cut +=1
	except WindowsError:
		pass

def Recent_Drawing():
	Drawing_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Paint\\Recent File List")
	try :
		cut = 0
		while True:
			name, value, type = EnumValue(Drawing_Key, cut)
			print (name + ': ' + value)
			cut += 1
	except WindowsError:
		pass

def Recent_Wordpad():
	Wordpad_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\wordpad\\Recent File List")
	try :
		cut = 0
		while True:
			name, value, type = EnumValue(Wordpad_Key, cut)
			print (name + ': ' + value)
			cut += 1
	except WindowsError:
		pass

def Recent_Hwp():
	Hwp_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Hnc\\Hwp\\9.0\\HwpFrame_KOR\\RecentFile")
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Hwp_Key, cut)
			print ('Recent_file_count : '+str(value))
			cut += 1
	except WindowsError:
		pass

def User_Profile_List():
	Pfl_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\S-1-5-21-885400413-3935149914-2550887433-1001")
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Pfl_Key,cut)
			print (name + ": " + str(value))
			cut += 1
	except WindowsError:
		pass

def Recent_Office():
	print("*Recent_Excel_file_Path")
	Office_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Microsoft\\Office\\16.0\\Excel\\Place MRU")
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Office_Key,cut)
			print (name + ": " + value)
			cut += 1
	except WindowsError:
		pass
	
	print("*Recent_PowerPoint_Place_Path")
	Office_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Microsoft\\Office\\16.0\\PowerPoint\\Place MRU")
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Office_Key,cut)
			print (name + ": " + value)
			cut += 1
	except WindowsError:
		pass
			
	print("*Recent_Word_file_Path")
	Office_Key = OpenKey(HKEY_CURRENT_USER,r"Software\\Microsoft\\Office\\16.0\\Word\\Place MRU")
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Office_Key,cut)
			print (name + ": " + value)
			cut += 1
	except WindowsError:
		pass
			
def Recent_Login_User():
	check = 0
	LU_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon")
	try:
			QueryValueEx(LU_Key, 'DefaultUserName')
	except WindowsError:
		check +=1
	if (check):
		print("Last_Login_User :" + QueryValueEx(LU_Key,'DefaultUserName'))
	else:
		print("There_is_no_account_last_logged_in.")

def Public_Directory():
	Pd_Key = OpenKey(HKEY_LOCAL_MACHINE,r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders')
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Pd_Key, cut)
			print(name + ": " + value)
			cut += 1
	except WindowsError:
		pass

def Recent_Run():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU"
	Run_Key = OpenKey(HKEY_USERS, Sub_Key)
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Run_Key, cut)
			if(name == 'MRUList'):
				pass
			else:
				print(name + ": " + value)
			cut += 1
	except WindowsError:
		pass

def Internet_Explorer_Config():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Internet Explorer\\Main"
	IEC_Key = OpenKey(HKEY_USERS, Sub_Key)
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(IEC_Key, cut)
			print(name + ": " + str (value))
			cut += 1
	except WindowsError:
		pass

def Internet_Explorer_Search_Log():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Internet Explorer\\TypedURLs"
	IESL_Key = OpenKey(HKEY_USERS,Sub_Key)
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(IESL_Key, cut)
			print(name + ": " + value)
			cut += 1
	except WindowsError:
		pass

def Recent_Dialog():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Windows\\CurrentVersion\Explorer\\ComDlg32\\LastVisitedPidlMRU"
	Dialog_Key = OpenKey(HKEY_USERS, Sub_Key)
	try:
		cut = 0
		while True:
			name, value, type = EnumValue(Dialog_Key, cut)
			print(name + ": " + str(value))
			cut += 1
	except WindowsError:
		pass

while(1):
	tmp = input("command: ")

	if(tmp == 'Windows_info'):
		Windows_info()
	elif(tmp == 'Fonts'):
		Fonts()
	elif(tmp == 'Recent_Drawing'):
		Recent_Drawing()
	elif(tmp == 'Recent_Wordpad'):
		Recent_Wordpad()
	elif(tmp == 'Recent_Hwp'):
		Recent_Hwp()
	elif(tmp =='User_Profile_List'):
		User_Profile_List()
	elif(tmp == 'Recent_Office'):
		Recent_Office()
	elif(tmp == 'Recent_Login_User'):
		Recent_Login_User()
	elif(tmp == 'Public_Directory'):
		Public_Directory()
	elif(tmp == 'Recnet_Run'):
		Recent_Run()
	elif(tmp == 'Internet_Explorer_Config'):
		Internet_Explorer_Config()
	elif(tmp == 'Internet_Explorer_Search_Log'):
		Internet_Explorer_Search_Log()
	elif(tmp == 'Recent_Dialog'):
		Recent_Dialog()


