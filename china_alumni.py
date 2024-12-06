import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import urllib.parse
import openpyxl
from openpyxl.styles import Font

# URL of the page containing rankings
url = 'https://www.4icu.org/cn/'

# Send a GET request to fetch the HTML content
response = requests.get(url)
html_content = response.text

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Find the table containing the university rankings
table = soup.find('table')

# Initialize a list to hold the names of the top 100 universities
top_universities = []

# Iterate over each row in the table (skipping the header)
for row in table.find_all('tr')[1:]:
    # Get all columns in the row
    cols = row.find_all('td')
    if cols:
        # Extract and append the university name from the second column
        university_name = cols[1].get_text(separator='|').split('|')[0].strip()
        try:
            if float(university_name):
                continue
        except ValueError:
            top_universities.append(university_name)

# Limit to top 100 universities
top_universities = top_universities[:60]

# Initialize Selenium WebDriver (make sure ChromeDriver is in your PATH)
driver = webdriver.Chrome()

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "University URLs"
# Write header row
sheet.append(["Name", "Link"])

try:
    # Dictionary to hold university names and their corresponding URLs
    university_urls = {}

    # Loop through each university to fetch URLs using search
    for university in top_universities:
        print(f"Searching for {university}...")

        # Prepare the search query
        query = urllib.parse.quote(university + " alumni official site")
        driver.get(f"https://www.bing.com/search?q={query}")

        # Wait for results to load
        time.sleep(2)

        # Find all result links on the search results page
        results = driver.find_elements(By.CSS_SELECTOR, "li.b_algo h2 a")

        # Extract the first valid URL from search results
        for result in results:
            url = result.get_attribute('href')
            if url:
                university_urls[university] = url
                break  # Stop after finding the first URL

    # Write results to the Excel sheet
    for i, (name, url) in enumerate(university_urls.items(), start=1):
        print(f"{i}. {name}: {url}")
        sheet.append([name, url])

        # Add hyperlink formatting to the URL
        cell = sheet.cell(row=i+1, column=2)
        cell.hyperlink = url
        cell.value = url
        # Make it look like a hyperlink
        cell.font = Font(color="0000FF", underline="single")

    # Save the Excel workbook
    workbook.save('university_urls_cn60.xlsx')

    print("URL extraction complete. Check 'university_urls_cn60.xlsx' for results.")

finally:
    # Close the WebDriver
    driver.quit()
