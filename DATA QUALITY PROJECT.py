# -*- coding: utf-8 -*-
"""
Created on Thu Oct 28 15:51:10 2021

@author: tiange.hou
"""
#autogui is used to peform automation, time is used to slow down the entry action
import pyautogui
import time
import pandas as pd

#use below function to get cell location (x,y)
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\test_auto_file.xlsx","tab_name",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['NAME'])
#print(num_row)
group_infor=[]
new_row=0
pyautogui.click(x, y) #this x, y is the start place you want to add your first NAME
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_reorder['NAME'][ind-1]!=test_file['NAME'][ind]:
        use_append=test_file['INFORMATION'][new_row:ind]
        pyautogui.typewrite(test_file['NAME'][ind-1])
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.8)
        for vn_nbr in use_append:           
            pyautogui.typewrite(INFORMATION)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.press('left')
        time.sleep(0.8)
