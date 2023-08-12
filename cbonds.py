import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

# Set up Selenium driver
url = 'https://cbonds.hnx.vn/to-chuc-phat-hanh/thong-tin-phat-hanh'
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)

# Access the below page to download `chromedriver.exe`.
# https://chromedriver.chromium.org/downloads
# Please remember to select the right version based on the version of Chrome browser.

# Define the path to the folder that contains the file chromedriver.exe
chrome_path = r'D:\...\chromedriver.exe'
service = Service(chrome_path)
driver = webdriver.Chrome(options=chrome_options, service=service)

# Load page
driver.get(url)

# Wait for page to load
time.sleep(2)

# Change dropdown to show 100 records per page
select_element = driver.find_element(By.ID, 'slChangeNumberRecord_1')
select_element.send_keys('100')

# Wait for table to load with new page size
time.sleep(2)

# Initialize variables
data = []
total_pages = 18 # 1790 / 100 = 17.9

# Loop through each page
for page_number in range(1, total_pages + 1):
    # Click the page link
    link = driver.find_element(By.LINK_TEXT, str(page_number))
    link.click()
    print(f'Page {page_number} is processing...')

    # Wait for page to load
    time.sleep(5)

    # Parse page with BeautifulSoup
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')

    # Extract data rows
    table = soup.find('table', id='tbReleaseResult')
    rows = table.find_all('tr')[1:]

    for row in rows:
        row_data = [td.text.strip() for td in row.find_all('td')]
        data.append(row_data)

print('Scraping finished!')


# Extract column headers
headers = [th.text.strip() for th in table.find('tr').find_all('th')]

# Create dataframe
df = pd.DataFrame(data, columns=headers)

# Export to Excel
df.to_excel('cbonds.xlsx', index=False)

print('Export to Excel successfully!')

# Close browser
driver.quit()
