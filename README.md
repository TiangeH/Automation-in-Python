# GUI-and-Data-Entry
use PyAutoGUI and for loop to remove manual data entry
A PyAutoGUI‑powered tool for automating repetitive data entry in legacy healthcare EMR systems.

Overview
Many older healthcare Electronic Medical Record (EMR) systems have a major limitation:
you cannot update a single field without manually navigating through every field in the record.

For example, if you only want to update the price of a needle, the EMR forces you to click through:

Item name
Item size
Vendor
Address
Contact info
Category
Price
And every other field
This makes even a simple update slow, repetitive, and error‑prone.

To eliminate this manual work, I built an automation tool using PyAutoGUI that reads data from Excel and performs the exact mouse clicks and keystrokes needed to update only the fields I care about — without touching the rest.

This tool automates all of that using PyAutoGUI, turning hours of clicking into minutes of automated work.

How It Works
1. Reads Excel Input
python
pd.read_excel("test_auto_file.xlsx", dtype=str, keep_default_na=False)
2. Uses PyAutoGUI to control the EMR
pyautogui.click(x, y) → click specific fields
pyautogui.typewrite(text) → type values
pyautogui.press('tab') → move to next field
pyautogui.press('down') → move down a list
time.sleep() → wait for EMR to load


Technologies Used
Python
PyAutoGUI
Pandas
Excel input files
Legacy EMR UI automation

Notes & Limitations
This tool interacts with the EMR UI exactly like a human user
It does not modify the EMR database directly
It requires stable screen layout and consistent coordinates
Time delays may need adjustment depending on EMR speed
Should be used responsibly in compliance with healthcare policies

Purpose of This Project
Real‑world automation of a legacy healthcare system
Practical use of PyAutoGUI for workflow optimization
Reducing repetitive manual labor
Improving accuracy and consistency
Building tools that save time in clinical operations


