import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

def extract_year_month(title_text):
    """
    Extracts year and month from the title text, assuming it's in the format:
    "YYYY оны эхний M сарын ..."
    """
    import re
    year_match = re.search(r"(\d{4}) оны", title_text)
    month_match = re.search(r"(\d+) сарын", title_text)
    year = year_match.group(1) if year_match else "unknown_year"
    month = month_match.group(1).zfill(2) if month_match else "unknown_month"
    return year, month


def rename_downloaded_file(folder, year, month):
    """
    Renames the most recent file in the download folder to 'yy-mm.xlsx'.
    """
    latest_file = max([f for f in os.listdir(folder) if f.endswith('.xlsx')], key=lambda f: os.path.getctime(os.path.join(folder, f)))
    new_name = f"{year[-2:]}-{month}.xlsx"
    os.rename(os.path.join(folder, latest_file), os.path.join(folder, new_name))

# Set up download options for Selenium
download_folder = "C:/Users/sugarkhuu/Downloads"
chrome_options = Options()

user_data_dir = r"C:\Users\sugarkhuu\AppData\Local\Google\Chrome\User Data"
profile_name = "Default"  # Change this if you want to use a different profile

# Set up Chrome options to use your existing profile
options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={user_data_dir}")
options.add_argument(f"profile-directory={profile_name}")

chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_folder,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.5938.92 Safari/537.36")


# Set up the WebDriver
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

# Open the URL
driver.get("https://gaali.mn/statistic")

# Iterate through each year (2014 to 2024)
for year_index in range(1, 12):  # Adjust for actual number of years
    # Iterate through each file in the selected year
    for file_index in range(1, 13):  # Adjust for actual number of files
        driver.get("https://gaali.mn/statistic")
        # Open the year dropdown
        year_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/section/section/main/div/div[3]/div[3]/div[2]/div[1]/div/div[2]/div/form/div/div[2]/div/span/div/div/div")))
        year_dropdown.click()
        
        # Select the year
        year_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/div/div/div/ul/li[{year_index}]")))
        year_option.click()
        time.sleep(1)  # Allow time for the files to load
        

        try:
            # Click on the file link to open its page
            file_link = wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[1]/div/div/section/section/main/div/div[3]/div[3]/div[2]/div[3]/div/div/div/div/div[{file_index}]/div/a")))
            file_link.click()
            
            # Extract the year and month information from the page
            file_title = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/section/section/main/div[2]/div/div[1]/div/div/div[1]/div/div[1]/h3"))).text
            year, month = extract_year_month(file_title)
            
            # Click the download button
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div/section/section/main/div[2]/div/div[1]/div/div/div[4]/a")))
            download_button.click()
            
            # Wait for download to complete
            time.sleep(5)  # Adjust if needed for download time
            
            # Rename the downloaded file
            rename_downloaded_file(download_folder, year, month)
            
            # Go back to the main page for the next file
            driver.back()
            time.sleep(1)
            
        except Exception as e:
            print(f"Failed to process file {file_index} for year {year_index}: {e}")

        
        # back_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/section/section/main/div[2]/div/div[1]/div/div/div[1]/div/div[2]/a    ")))
        # back_button.click()

# Close the driver
driver.quit()



