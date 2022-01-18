# -*- coding: utf-8 -*-
"""
Created on Thu Oct 28 15:51:10 2021

@author: tiange.hou
"""
#autogui is used to peform automation, time is used to slow down the entry action
import pyautogui
import time
import pandas as pd

#use below function to get cell location
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\test_auto_file.xlsx","tab_name",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['ITEM_NBR'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
pyautogui.click(x=1398, y=-93)
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        #pyautogui.click(x=1516, y=-245)
        #time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.8)
        for vn_nbr in use_append:           
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        #pyautogui.press('enter')
        #time.sleep(0.8)
        pyautogui.press('left')
        time.sleep(0.8)
  
