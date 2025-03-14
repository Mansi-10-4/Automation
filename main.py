import time
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Setup Chrome driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Open the website
driver.get("https://books.toscrape.com/")
time.sleep(3)

print("Title: ", driver.title)
print("""fetching data from website using selenium and 
                storing it in csv file and then opening it in excel 
                using pyautogui and entering data in it using pyautogui""")

# Scrape product names and prices
products = driver.find_elements(By.CSS_SELECTOR, ".product_pod h3 a")
prices = driver.find_elements(By.CSS_SELECTOR, ".product_price .price_color")

# Store data in a list
data = []
for product, price in zip(products, prices):
    data.append([product.text, price.text])

# Convert to DataFrame
df = pd.DataFrame(data, columns=["Product Name", "Price"])

# Save data to a CSV file
df.to_csv("product_data.csv", index=False)

# Close the browser
driver.quit()

# Read data from the saved CSV file
df = pd.read_csv("product_data.csv")

# Open Excel (Manually open Excel or use PyAutoGUI to do it)
pyautogui.hotkey("win", "r")  # Open Run command
time.sleep(1)
pyautogui.write("excel")  # Type 'excel' to open it
pyautogui.press("enter")
time.sleep(5)  # Wait for Excel to open

# Click on the first cell in Excel
pyautogui.click(x=200, y=200)  # Adjust based on your screen resolution

# Loop through DataFrame and enter data
for index, row in df.iterrows():
    pyautogui.write(row["Product Name"])  # Type Product Name
    pyautogui.press("tab")  # Move to the next cell
    pyautogui.write(str(row["Price"]))  # Type Price
    pyautogui.press("enter")  # Move to the next row
    time.sleep(2)

# Close Excel
pyautogui.hotkey("alt", "f4")  # Close Excel
time.sleep(1)
pyautogui.press("right")  # Navigate to "Don't Save" option
time.sleep(1)
pyautogui.press("enter")  # Confirm "Don't Save"
