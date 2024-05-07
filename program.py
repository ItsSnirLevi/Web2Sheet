import requests
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

file_path = ""

def add_to_excel():
    status_label.config(text="Adding data to Excel file...", fg="blue")
    global file_path
    url = url_entry.get()

    if not url.startswith('https://carwiz.co.il/used-cars/') or len(url.split('/')[-1]) == 0:
        status_label.config(text="Error: Invalid URL.", fg="red")
        return

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raises an exception for 4xx and 5xx status codes
    except requests.RequestException:
        status_label.config(text="Error: Invalid URL or unable to connect.", fg="red")
        return

    driver.get(url)

    try:
        # Wait for the details button to be clickable with a longer timeout
        btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "MuiButtonBase-root")]')))
        ActionChains(driver).move_to_element(btn).perform()
        driver.execute_script("arguments[0].click();", btn)

        # Find the price
        price = driver.find_element(by='xpath', value='//div[@class="css-lc7fha"]').text

        # Find the details
        details = driver.find_elements(by='xpath', value='//tr[contains(@class, "css-fnxg5y")]')
        detail_dict = {}
        for detail in details:
            key = detail.find_element(by='xpath', value='.//h3[contains(@class, "css-1ikgxyv")]').text
            value = detail.find_element(by='xpath', value='.//td[contains(@class, "css-w2uogn")]').text
            detail_dict[key] = value

        # if a path is not provided than create a new file
        if file_path == "":
            create_new_file()

        existing_data = pd.DataFrame()
        if file_path != "":
            try:
                existing_data = pd.read_excel(file_path)
            except FileNotFoundError:
                pass

        # Add new data to existing data
        new_data = pd.DataFrame(detail_dict, index=[0])
        new_data['Price'] = price
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)

        # Save updated data to Excel file
        updated_data.to_excel(file_path, index=False)

        status_label.config(text="Data added to Excel file.", fg="green")

    except Exception as e:
        print("Error:", e)
        status_label.config(text="Error: Failed to add data to Excel file.", fg="red")

def choose_file():
    global file_path
    new_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if new_file_path:
        file_path = new_file_path
        add_button.config(text="Add Data to file")
    file_label.config(text="File: " + file_path)

def create_new_file():
    global file_path
    new_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if new_file_path:
        file_path = new_file_path
        file_label.config(text="File: " + file_path)
        add_button.config(text="Add Data to file")


def clear_status(event):
    status_label.config(text="")

# Initialize WebDriver
options = Options()
options.add_argument('--headless')
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Create GUI
root = tk.Tk()
root.title("Web2Sheet")
root.iconbitmap('icon.ico')

# Choose file button
file_label = tk.Label(root, text="Choose Excel file:", padx=10, pady=5)
file_label.grid(row=0, column=0, sticky="w")
choose_file_button = tk.Button(root, text="Choose File", command=choose_file)
choose_file_button.grid(row=0, column=1, padx=10, pady=5)

# Create new file button
create_new_button = tk.Button(root, text="Create New File", command=create_new_file)
create_new_button.grid(row=0, column=2, padx=10, pady=5)

# URL entry
url_label = tk.Label(root, text="Enter URL:", padx=10, pady=5)
url_label.grid(row=1, column=0, sticky="w")
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=1, column=1, padx=10, pady=5)
url_entry.bind("<Button-1>", clear_status)

# Add button
add_button = tk.Button(root, text="Create file and Add Data", command=add_to_excel)
add_button.grid(row=2, column=0, columnspan=2, pady=10)

# Status label
status_label = tk.Label(root, text="", padx=10, pady=5)
status_label.grid(row=3, column=0, columnspan=2)

for i in range(3):
    root.columnconfigure(i, weight=1)
root.rowconfigure(0, weight=1)

root.mainloop()
