Web Scraping and Excel Automation

Overview

This project demonstrates how to scrape data from a website using Selenium, save the extracted data to a CSV file using Pandas, and then automate data entry into Microsoft Excel using PyAutoGUI.

Prerequisites

Ensure you have the following installed:

Python 3.x

Google Chrome Browser

Required Python Libraries:

selenium

pandas

pyautogui

webdriver-manager

Installation

Run the following command to install the necessary dependencies:

pip install selenium pandas pyautogui webdriver-manager

How It Works

Web Scraping

The script uses Selenium to open the https://books.toscrape.com/ website.

It extracts book titles and prices.

The scraped data is saved into a CSV file named product_data.csv.

Excel Automation

The script opens Microsoft Excel using PyAutoGUI.

It enters the extracted book names and prices into an Excel sheet.

The script then closes Excel without saving.

Usage

Run the script using:

python script.py

Ensure that Microsoft Excel is installed on your system for the automation part to work.

Notes

The script assumes that Excel is opened via the Run command (win + r).

Adjust pyautogui.click(x=200, y=200) coordinates as per your screen resolution to target the correct cell.

PyAutoGUI controls the mouse and keyboard, so do not use your computer while the script is running.

Disclaimer

This script automates browser and system interactions. Ensure you understand how it works before running it on your machine. Modify it as needed for safety and efficiency.
