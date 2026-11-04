import pyautogui
import time
import pyperclip
import re
import os
import openpyxl
import pyscreeze    
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1
from datetime import datetime

# Get current date and time
file_date = datetime.now().strftime("%Y-%m-%d")
file_time = datetime.now().strftime("%H:%M:%S")
# Get current working folder
current_folder = os.getcwd()

# Excel file path
file_path = os.path.join(current_folder, 'daily_report_'+file_date+'.xlsx')

# open Microsoft Edge Browser
print("Opening Microsoft Edge Browser...")
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')
time.sleep(1)
#enter website URL
pyautogui.typewrite('https://www.accuweather.com/en/in/coimbatore/206673/hourly-weather-forecast/206673', interval=0.1)
pyautogui.press('enter')
# wait for the page to load
time.sleep(5)
# Select all content and copy
pyautogui.hotkey('ctrl','a')
pyautogui.hotkey('ctrl','c')
time.sleep(2)

# Store copied content in variable
website_data = pyperclip.paste()

time.sleep(2)
# Close browser
pyautogui.hotkey("alt", "f4")
time.sleep(2)
print("Website content stored successfully!")
print("Browser closed!")
temperature = re.search(
    r"\d{1,2}\s*°\s*C",
    website_data
)

if temperature:
    temp_value = temperature.group()
    print("Temperature:", temp_value)
else:
    print("Temperature not found")

# Create a new workbook (overwrite existing file)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Weather Data"

# Add headers
ws.append(["DATE", "TIME", "TEMPERATURE", "Comment"])

# Add weather data
ws.append([
    datetime.now().strftime("%d-%m-%Y"),
    datetime.now().strftime("%H:%M:%S"),
    temp_value,
    "Weather temperature for Coimbatore"
])

# Save Excel
wb.save(file_path)
print("Excel saved successfully!")
print("File location:", file_path)
# Check whether Excel file exists
if os.path.exists(file_path):

    # Open Excel workbook
    os.startfile(file_path)

    print("Excel workbook opened successfully!")

    # Wait for Excel to load
    time.sleep(5)

    # Take screenshot
    screenshot = pyautogui.screenshot()

    # Save screenshot in current working folder
    screenshot_path = os.path.join(
        current_folder, "weather_screenshot.png"
    )

    screenshot.save(screenshot_path)

    print("Screenshot captured successfully!")
    print("Screenshot location:", screenshot_path)
    # Close Excel
    pyautogui.hotkey("alt", "f4")

    time.sleep(2)

else:
    print("Excel workbook not found!")


