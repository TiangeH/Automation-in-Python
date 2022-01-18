# -*- coding: utf-8 -*-
"""
Created on Thu Oct 28 15:51:10 2021

@author: tiange.hou
"""
import pyautogui
import time
import pandas as pd


print(pyautogui.position())
#INACTIVE ITEMS
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","Inactive_item",dtype=str, keep_default_na=False)
pyautogui.click(1844,-155)
time.sleep(0.8)
for ind in VENDOR_BUPDATE.index:
    pyautogui.typewrite('14')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)##
    pyautogui.typewrite(VENDOR_BUPDATE['itm_nbr'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)##
    pyautogui.press('enter')
    time.sleep(0.8)##
    
    
#inactive items HHS
print(pyautogui.position())
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","Inactive_item",dtype=str, keep_default_na=False)
pyautogui.click(x=1646, y=-94)
time.sleep(0.8)
for ind in VENDOR_BUPDATE.index:
    pyautogui.typewrite(VENDOR_BUPDATE['itm_nbr'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)##
    pyautogui.click(x=2500, y=659)
    time.sleep(6)#

#BECAUSE HHS HAS NON-STOCK REQ AND REQ ISSUE,USE THIS TEMPLATE TO AVOID PRINT PREVIEW ISSUE
print(pyautogui.position())
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","Inactive_item",dtype=str, keep_default_na=False)
pyautogui.click(x=1646, y=-94)
time.sleep(0.8)
for ind in VENDOR_BUPDATE.index:
    pyautogui.typewrite(VENDOR_BUPDATE['itm_nbr'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)##
    pyautogui.click(x=2500, y=659)
    time.sleep(4)#
    pyautogui.click(x=2382, y=617)
    time.sleep(1.2)#
    pyautogui.click(x=3214, y=-230)
    time.sleep(1.2)#
    pyautogui.click(x=3063, y=-136)
    time.sleep(1.2)#
    pyautogui.click(x=1646, y=-94)
    
    

    



#ENTER DEPTS IN REQ
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","Sheet1",dtype=str, keep_default_na=False)
pyautogui.click(1381,-184)
time.sleep(0.8)

for ind in VENDOR_BUPDATE.index:
    pyautogui.typewrite(VENDOR_BUPDATE['REQUISITION'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)##
    
#ADD VENDOR
print(pyautogui.position())
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","ADD_VENDOR",dtype=str, keep_default_na=False)
pyautogui.click(1523,-246)
for ind in VENDOR_BUPDATE.index:   
    pyautogui.typewrite(VENDOR_BUPDATE['itm_nbr'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    #pyautogui.click(1441,-44)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(3273,-279)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)
    #pyautogui.click(1391,4)
    #time.sleep(0.8)
    pyautogui.typewrite("P0000000582")
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(1756,122)
    time.sleep(0.5)
    pyautogui.press('tab')#FROM THIS PART IS FOR THP
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)#TILL THIS PART IS FOR THP
    pyautogui.click(1828,118)
    time.sleep(0.5)
    pyautogui.click(1391,73)
    time.sleep(0.5)
    pyautogui.typewrite(VENDOR_BUPDATE['vn_uop'][ind])
    time.sleep(0.5)
    pyautogui.click(1600,72)
    time.sleep(0.5)
    pyautogui.typewrite(VENDOR_BUPDATE['vu_cost_uop'][ind])
    time.sleep(0.5)
    pyautogui.click(1520,136)
    time.sleep(0.5)
    pyautogui.typewrite(VENDOR_BUPDATE['vCatNo'][ind])
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(1821,262)
    time.sleep(0.5)
    pyautogui.click(3272,-281)
    time.sleep(0.5)
    pyautogui.click(1765,264)
    time.sleep(0.5)
    pyautogui.click(2197,-274)
    time.sleep(0.5)
    pyautogui.click(2197,-274)
    time.sleep(0.5)
    pyautogui.click(1762,421)
    time.sleep(0.5)
    
    
#REACTIVE ITEM
print(pyautogui.position())
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REACTIVE",dtype=str, keep_default_na=False)
pyautogui.click(1523,-246)
for ind in VENDOR_BUPDATE.index:   
    pyautogui.typewrite(VENDOR_BUPDATE['itm_nbr'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)    
    pyautogui.typewrite("1")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(1958,-110)
    time.sleep(0.8)
    pyautogui.press('backspace')
    time.sleep(0.8)
    pyautogui.typewrite("Y")
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)
    
#reorder vn

print(pyautogui.position())

test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
#print(test_for_reorder)
test_for_reorder['VN_AMOUNT']=test_for_reorder['VN_AMOUNT'].astype(str)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("6")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(1498,-20)
    time.sleep(0.8)   
    if test_for_reorder['VN_AMOUNT'][ind] =='2':
        pyautogui.typewrite(test_for_reorder['VN_1'][ind])
        time.sleep(0.8)
        pyautogui.press('enter')
        pyautogui.typewrite(test_for_reorder['VN_2'][ind])
        time.sleep(0.8)
        pyautogui.press('enter')
        pyautogui.click(3272,-277)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click(2195,-272)
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.5)
    elif test_for_reorder['VN_AMOUNT'][ind] =='3':
        pyautogui.typewrite(test_for_reorder['VN_1'][ind])
        time.sleep(0.8)
        pyautogui.press('enter')
        pyautogui.typewrite(test_for_reorder['VN_2'][ind])
        time.sleep(0.8)
        pyautogui.press('enter')
        pyautogui.typewrite(test_for_reorder['VN_3'][ind])
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.click(3272,-277)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click(2195,-272)
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.5)
        
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
#print(test_for_reorder)
test_for_reorder['VN_AMOUNT']=test_for_reorder['VN_AMOUNT'].astype(str)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("6")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(1498,-20)
    time.sleep(0.8)   
    pyautogui.typewrite('P0000000582')
    time.sleep(0.8)
    pyautogui.press('enter')
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)

#REORDER OSL VENDOR FOR VN_AMOUNT WITHOUT REFORMAT THE FILE
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['ITEM_NBR'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
pyautogui.click(x=1516, y=-245)
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        #pyautogui.click(x=1516, y=-245)
        #time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.typewrite("6")
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.click(x=1490, y=-19)
        for vn_nbr in use_append:           
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.click(x=3273, y=-279)
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.click(x=2200, y=-275)
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        new_row=ind 

#REORDER HHS VENDOR FOR VN_AMOUNT WITHOUT REFORMAT THE FILE
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['ITEM_NBR'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
pyautogui.click(x=1588, y=-77)
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        #pyautogui.click(x=1516, y=-245)
        #time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.click(x=1811, y=678)
        time.sleep(0.8)
        pyautogui.click(x=1699, y=288)
        time.sleep(0.8)       
        for vn_nbr in use_append:           
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.click(x=2535, y=535)
        time.sleep(4)
        pyautogui.click(x=2501, y=679)
        time.sleep(0.8)
        new_row=ind 


#for excel on stakcflow
#REORDER OSL VENDOR FOR VN_AMOUNT WITHOUT REFORMAT THE FILE
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
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
##############################################        

#reorder vendor order in HHS Expanse(move HHS to the left corner under the alt bar)
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
#test_for_reorder['VN_AMOUNT']=test_for_reorder['VN_AMOUNT'].astype(str)
pyautogui.click(x=1585, y=-101)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1800, y=670)
    time.sleep(0.8)
    pyautogui.click(x=1697, y=273)
    time.sleep(0.8)
    pyautogui.typewrite('H01465')
    time.sleep(0.5)
    pyautogui.click(x=2536, y=494)
    time.sleep(8)
    pyautogui.click(x=2483, y=664)
    time.sleep(0.8)


#CHANGE VCAT IN HHS
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","change_vcat",dtype=str, keep_default_na=False)
#test_for_reorder['VN_AMOUNT']=test_for_reorder['VN_AMOUNT'].astype(str)
pyautogui.click(x=1585, y=-101)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1511, y=678)
    time.sleep(0.8)
    pyautogui.press('backspace')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['vcat'][ind])
    pyautogui.press('enter')  
    time.sleep(0.8)
    pyautogui.click(x=2070, y=281)
    time.sleep(0.8) 
    pyautogui.click(x=2438, y=630)
    time.sleep(4)
    pyautogui.click(x=2365, y=636)
    time.sleep(2.5)
    pyautogui.click(x=2490, y=655)
    time.sleep(2.5)
  
#desktop version
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","change_vcat",dtype=str, keep_default_na=False)
#test_for_reorder['VN_AMOUNT']=test_for_reorder['VN_AMOUNT'].astype(str)
pyautogui.click(x=1949, y=12)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1887, y=689)
    time.sleep(0.8)
    pyautogui.press('backspace')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['vcat'][ind])
    pyautogui.press('enter')  
    time.sleep(1.5)
    pyautogui.click(x=2217, y=373)
    time.sleep(1.5)
    pyautogui.click(x=2623, y=682)
    time.sleep(3.5)
    pyautogui.click(x=2533, y=687)
    time.sleep(1.5)
    pyautogui.click(x=2704, y=671)
    time.sleep(1.5)








print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REORDER_VN",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['ITEM_NBR'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
pyautogui.click(x=1585, y=-101)
time.sleep(0.8)
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        #pyautogui.click(x=1516, y=-245)
        #time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        pyautogui.click(x=1800, y=670)
        time.sleep(0.8)
        pyautogui.click(x=1697, y=273)
        time.sleep(0.8)
        for vn_nbr in use_append:           
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.click(x=2536, y=494)#done
        time.sleep(20)
        pyautogui.click(x=2483, y=664)
        time.sleep(0.8)
        new_row=ind     
  
---------------------------------------------------------------------------------------------------------------------------------
#REORDER VENDOR ORDER TEST IN EXELEXCEL

print(pyautogui.position())    
        
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","Sheet1",dtype=str, keep_default_na=False)
num_row = len(test_for_reorder['ITEM_NBR'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        print("use thi sappend:",use_append)
        print("old_row=",test_for_reorder['ITEM_NBR'][ind-1])
        pyautogui.click(445,258)
        time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.5)
        # pyautogui.click(549,256)
        # time.sleep(0.8)
        print("appen lenth:",len(use_append))
        pyautogui.click(549,256)
        for vn_nbr in use_append:           
            # pyautogui.click(549,256)
            # time.sleep(0.8)
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        time.sleep(0.5)
        new_row=ind 

#to use this one, must make sure the starting position is clicked
group_infor=[]
new_row=0
pyautogui.click(445,258)
for ind in range(1, num_row):
    if test_for_reorder['ITEM_NBR'][ind-1]!=test_for_reorder['ITEM_NBR'][ind]:
        use_append=test_for_reorder['VN_1'][new_row:ind]
        print("use thi sappend:",use_append)
        print("old_row=",test_for_reorder['ITEM_NBR'][ind-1])        
        time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind-1])
        time.sleep(0.8)
        pyautogui.press('tab')
        time.sleep(0.5)
        print("appen lenth:",len(use_append))
        #pyautogui.click(549,256)
        for vn_nbr in use_append:           
            pyautogui.typewrite(vn_nbr)
            time.sleep(0.8)
            pyautogui.press('down')
            time.sleep(0.5)
        time.sleep(0.5)
        new_row=ind 
        pyautogui.press('left')
-----------------------------------------------------------------------------------------------------------------------------------
#REMOVE INVENTORY (THP with multiple inventories)
print(pyautogui.position())
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REMOVE_INVENTORY",dtype=str, keep_default_na=False)
test_for_reorder['stk_inventory']=test_for_reorder['stk_inventory'].astype(str)
num_row = len(test_for_reorder['stk_inventory'])
#print(num_row)
new_row = 0

group_infor=[]
new_row=0
pyautogui.click(x=1490, y=-247)#inventory field
for ind in range(1, num_row):
    if test_for_reorder['stk_inventory'][ind-1]!=test_for_reorder['stk_inventory'][ind]:
        use_append=test_for_reorder['stk_nbr'][new_row:ind]
        #pyautogui.click(x=1516, y=-245)
        #time.sleep(0.8)
        pyautogui.typewrite(test_for_reorder['stk_inventory'][ind-1])
        time.sleep(0.8)
        pyautogui.press('enter')
        time.sleep(0.8)
        for stk_nbr in use_append:           
            pyautogui.typewrite(stk_nbr)
            time.sleep(0.8)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.press('backspace')
            time.sleep(0.5)
            pyautogui.typewrite('N')
            time.sleep(0.8)
            pyautogui.click(x=3272, y=-276)
            time.sleep(0.8)
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
        pyautogui.click(x=3270, y=-250)#done
        time.sleep(0.8)
        pyautogui.click(x=1444, y=-220)
        time.sleep(0.8)
        pyautogui.click(x=1654, y=262)
        time.sleep(0.8)
        pyautogui.click(x=1831, y=262)
        time.sleep(0.8)
        new_row=ind     

#thp with single inventories

test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","REMOVE_INVENTORY",dtype=str, keep_default_na=False)
pyautogui.click(x=1487, y=-204)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['stk_nbr'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('backspace')
    time.sleep(0.5)
    pyautogui.typewrite('N')
    time.sleep(0.8)
    pyautogui.click(x=3272, y=-276)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

--------------------------------------------------------------------------
#CANCEL REQ
print(pyautogui.position())

test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","CANCEL_REQ",dtype=str, keep_default_na=False)
#print(test_for_reorder)
test_for_reorder['REQ']=test_for_reorder['REQ'].astype(str)
pyautogui.click(1839,-245)
for ind in test_for_reorder.index:
    pyautogui.typewrite('36')
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['SITE'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['REQ'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite('5')   
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(2085,-277)
    time.sleep(0.8)
    
print(pyautogui.position())   
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3.xlsx","Sheet2",dtype=str, keep_default_na=False)

pyautogui.click(1678,-150)
for ind in test_for_reorder.index:
    pyautogui.click(1678,-150)
    pyautogui.press('enter')
    time.sleep(0.8)     
    
#################################
#UPDATE CMN_NAME,vcat,mfrcat FOR OSL
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_cmn_name",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("1")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1767, y=-24)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.click(x=3268, y=-279)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1647, y=146)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.5)   
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(x=1823, y=260)#this line is for vcat is dup
    time.sleep(0.8) 
    pyautogui.click(x=2144, y=146)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.5)
    pyautogui.click(x=3268, y=-279)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)  
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    

#update price for osl
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_price",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1663, y=76)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.typewrite(test_for_reorder['price'][ind])
    time.sleep(0.8)
    pyautogui.click(x=1445, y=216)
    time.sleep(0.8)
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)




#update price for osl
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_price",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1445, y=216)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.click(x=1642, y=75)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.typewrite(test_for_reorder['price'][ind])
    time.sleep(0.8)
    pyautogui.click(x=1445, y=216)
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract'][ind])
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract_expiry'][ind])
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_price",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1642, y=75)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.typewrite(test_for_reorder['price'][ind])
    time.sleep(0.8)
    pyautogui.click(x=1445, y=216)
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract'][ind])
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract_expiry'][ind])
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    
#update price for THP

print(pyautogui.position()) 
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_price",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1384, y=210)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.click(x=1599, y=79)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.typewrite(test_for_reorder['price'][ind])
    time.sleep(0.8)
    pyautogui.click(x=1384, y=215)
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract'][ind])
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract_expiry'][ind])
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","update_price",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("2")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=1642, y=75)
    time.sleep(0.8) 
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    time.sleep(0.8) 
    pyautogui.typewrite(test_for_reorder['price'][ind])
    time.sleep(0.8)
    pyautogui.click(x=1470, y=214)
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract'][ind])
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['contract_expiry'][ind])
    time.sleep(0.8)
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    
    
    
########################SANDY FILE
#ADDING INVENTORY
print(pyautogui.position()) 
test_for_reorder=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","add_inventory",dtype=str, keep_default_na=False)
#print(test_for_reorder)
pyautogui.click(1523,-246)
for ind in test_for_reorder.index:
    pyautogui.typewrite(test_for_reorder['ITEM_NBR'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite("10")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.typewrite(test_for_reorder['stk_inventory'][ind])
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(x=1478, y=252)
    time.sleep(0.5)
    pyautogui.typewrite('EA')
    time.sleep(0.5)   
    pyautogui.click(3272,-277)
    time.sleep(0.8)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.click(2195,-272)
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.5)
    
    
#HHS Reactive non-stock req
print(pyautogui.position())
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","hhs_non_stock",dtype=str, keep_default_na=False)
pyautogui.click(x=1601, y=-61)
for ind in VENDOR_BUPDATE.index:   
    pyautogui.typewrite(VENDOR_BUPDATE['non_stock'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)   
    pyautogui.typewrite("Y")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=2496, y=661)
    time.sleep(0.8)

#INACTIVE
VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","hhs_non_stock",dtype=str, keep_default_na=False)
pyautogui.click(x=1601, y=-61)
for ind in VENDOR_BUPDATE.index:   
    pyautogui.typewrite(VENDOR_BUPDATE['non_stock'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)   
    pyautogui.typewrite("N")
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.click(x=2496, y=661)
    time.sleep(0.8)

VENDOR_BUPDATE=pd.read_excel(r"C:\Users\Tiange.Hou\Desktop\Book3_3_3.xlsx","hhs_non_stock",dtype=str, keep_default_na=False)
pyautogui.click(x=1601, y=-61)
for ind in VENDOR_BUPDATE.index:   
    pyautogui.typewrite(VENDOR_BUPDATE['non_stock'][ind])
    pyautogui.press('enter')
    time.sleep(0.8)
    pyautogui.press('enter')
    time.sleep(0.8)   
    pyautogui.click(x=2418, y=670)
    time.sleep(0.8)