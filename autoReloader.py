# Python program to Automatically refresh web browser when file is modified
# Date :- 7-jan-2017
# Author :- Kartik

import os
import time
import webbrowser
import win32com.client


print("PYTHON PROGRAM TO AUTOMATICALLY REFRESH WEBBROWSER WHEN FILE IS MODIFIED"
      "\nFILES SUPPORTED ARE :- \n\t\t\t1) HTML"
      "\n\t\t\t2) HTM"
      "\n\t\t\t3) JS"
      "\n\t\t\t4) PHP\n"
      "\n-------------------------CREATED BY KARTIK K-----------------------\n")

print("Enter the File name to add to watchlist    Ex :-  D:\html\index.html")
print("\n\t\t 'OR'")
print("\nDrag and Drop the FILE to add to watchlist")
print("\n\t\t 'OR'")
print("\n'CTRL + C' to Terminate")

# Getting the file Path
path = input()


# Remove '"' double quotes
if path[0] == '"':
    path = path.strip('"')
    print('  " -- (Double quotes)'" Successfully stripped of")

# Checking extensions
if not path.endswith('.html') | path.endswith('.htm') | path.endswith('.js') | path.endswith('.php') :
    print("Only Files with extensions '.html' '.htm' '.js' '.php' are Supported")
    input("Press Any key to Exit")
    exit()

checkPath = path[0]
if not checkPath.isupper():
    path = path[1:]


# Getting the last modified time
try:
    lastTime = os.path.getmtime(path)
except:
    print("Enter the proper path... , Process terminated.")
    input("Press Any key to Exit")
    exit()


# Dispatch Wscript.shell
shell = win32com.client.Dispatch("Wscript.shell")


webbrowser.open(path,new=0)

print("WatchDog started on file", path)
while True:
    time.sleep(1)
    newTime = os.path.getmtime(path)
    if lastTime != newTime:
        print("File Modified  --->  Update browser ", newTime)
        lastTime = os.path.getmtime(path)
        # bring chrome browser to focus and stroke 'F5' key
        shell.AppActivate("chrome")
        # print("Focus on Google chrome")
        shell.SendKeys("{F5}",0)
        # print("Successfull stroke of 'F5' key")
        

input("Press Any key to Exit")
