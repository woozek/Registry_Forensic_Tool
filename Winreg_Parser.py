import os
import struct
from datetime import datetime, timedelta
from winreg import *

def Win_ts(timestamp):
    WIN32_EPOCH = datetime(1601, 1, 1)
    return WIN32_EPOCH + timedelta(microseconds=timestamp//10, hours=9)

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
    Last_ShutDown_Time = QueryValueEx(Wininfo_Key, 'ShutdownTime')[0]
    Last_ShutDown_Time = Win_ts(struct.unpack_from("<Q", Last_ShutDown_Time)[0]).strftime('%Y:%m:%d-%H:%M:%S.%f')
    print (Last_ShutDown_Time)

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
		cut = 3
		while True:
			name, value, type = EnumValue(Hwp_Key, cut)
			value = value.decode('UTF-16')
			print (name + ': '+ str(value))
			cut += 2
	except WindowsError:
		pass

def User_Profile_List():
	Sub_Key = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList"
	Sub_Key += "\\"+Windows_User_Id_Set()
	Pfl_Key = OpenKey(HKEY_LOCAL_MACHINE,Sub_Key)
	print(Sub_Key)
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
	cut = 0
	LU_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon")
	try:
		QueryValueEx(LU_Key, 'DefaultUserName')
	except WindowsError:
		cut +=1
	if (cut):
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

def Regedit_LastKey():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit"
	Regedit_LastKey_Key = OpenKey(HKEY_USERS, Sub_Key)
	print("Regedit_LastKey : "+QueryValueEx(Regedit_LastKey_Key,"LastKey")[0])

def Recent_Dialog():
	Sub_Key = Windows_User_Id_Set()
	Sub_Key += "\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ComDlg32\\LastVisitedPidlMRU"
	Dialog_Key = OpenKey(HKEY_USERS, Sub_Key)
	try:
		cut = 1
		Data_Name, Data_Value = [], []
		while True:
			name, value, type = EnumValue(Dialog_Key, cut)
			Data_Name.append(int(name))
			Data_Value.append(value.decode('utf-16',errors="ignore").split("\x00")[0])
			cut += 1
	except WindowsError:
		i = 1
		Data_Name.sort()
		while True:
			print(str(Data_Name[i-1])+": "+str(Data_Value[i-1]))
			i += 1
			if i >= cut:
				break
			else:
				pass

def run():
	command_check = ["Windows_info()", "Recent_Drawing()", "Recent_Wordpad()", "Recent_Hwp()", "User_Profile_List()", "Recent_Office()", "Recent_Login_User()", "Public_Directory()", "Recent_Run()", "Internet_Explorer_Config()", "Internet_Explorer_Search_Log()", "Recent_Dialog()", "Regedit_LastKey()" ]
	help_check = ["rg -h()", "rg-h()", "rg -help()", "rg-help()"]
	cut = 0
	while True:
		tmp = input("command: ")
		command = tmp + '()'
		if(command in command_check):
			eval(command)
		elif(command in help_check):
			while True:
				print('- '+command_check[cut])
				cut+=1
				if(cut == 13):
					cut = 0
					break
		else:
			print('error')

run()
