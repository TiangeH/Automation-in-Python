# -*- coding: utf-8 -*-
"""
Created on Thu Oct 28 15:51:10 2021

@author: tiange.hou
"""
import pyautogui
import time
import pandas as pd


print(pyautogui.position())
#use below function to get cell location (x,y)
print(pyautogui.position())
test_for_enter=pd.read_excel(r"C:\Users\Tiange\Desktop\test_auto_file.xlsx","tab_name",dtype=str, keep_default_na=False)
num_row = len(test_for_enter['NAME'])
#print(num_row)
group_infor=[]
new_row=0
pyautogui.click(x=1172, y=274) #this x, y is the start place you want to add your first NAME
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_enter['NAME'][ind-1]!=test_for_enter['NAME'][ind]:
        use_append=test_for_enter['INFORMATION'][new_row:ind]
        pyautogui.typewrite(test_for_enter['NAME'][ind-1])
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.8)
        for INFORMATION in use_append:           
            pyautogui.typewrite(INFORMATION)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.press('left')
        time.sleep(0.8)
        new_row=ind
