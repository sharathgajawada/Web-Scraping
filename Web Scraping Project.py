from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup as bs
import pandas as pd
import itertools
import numpy
import openpyxl
import requests
import time
from selenium.common.exceptions import TimeoutException


driver: WebDriver = webdriver.Chrome(executable_path="C:/Users/shara/Downloads/chromedriver_win32/chromedriver.exe")
driver.get('https://nhts.telangana.gov.in/Supervisors/CGMStatusReport.aspx')
page = requests.get('https://nhts.telangana.gov.in/Supervisors/CGMStatusReport.aspx')
soup1 = bs(page.content,"html.parser")
soup2 = bs(soup1.prettify(),"html.parser")
#'JAN',
months = [ 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
years = ['2022', '2023']
districts = ['ADILABAD', 'KUMARAMBHEEM-ASIFABAD', 'MANCHERIAL', 'NIRMAL', 'NIZAMABAD','JAGITYAL','PEDDAPALLI','JAYASHANKAR-BHUPALAPALLI','BHADRADRI-KOTHAGUDEM','MAHABUBABAD','WARANGAL','KARIMNAGAR','RAJANNA-SIRICILLA','KAMAREDDY','SANGAREDDY','MEDAK','SIDDIPET','JANGAON','YADADRI-BHUVANGIRI','MEDCHAL - MALKAJIGIRI','HYDERABAD','RANGAREDDY','VIKARABAD','MAHABUBNAGAR','JOGULAMBA-GADWAL','WANAPARTHY','NAGARKURNOOL','SURYAPET','MULUGU','NARAYANPET']
##JAN defective--'HANMAKONDA','NALGONDA','KHAMMAM'
wait = WebDriverWait(driver, 10)
for month in months:
    elementM = driver.find_element(By.ID, "ddlMonth")
    month_select = Select(elementM)
    month_select.select_by_visible_text(month)

    for year in years:
        elementY = driver.find_element(By.ID, "ddlYear")
        year_select = Select(elementY)
        year_select.select_by_visible_text(year)

        for district in districts:
            elementD = driver.find_element(By.ID, "ddlDist")
            district_select = Select(elementD)
            district_select.select_by_visible_text(district)

            # Select the projects
            elementP = driver.find_element(By.ID, "ddlProject")
            project_select = Select(elementP)
            projects = [option.get_attribute('value') for option in project_select.options][
                       1:]  # exclude the first option ('--Select--')

            projectst = [option.text for option in project_select.options][1:]

            for project in projects:
                project_select.select_by_value(project)
                
                # Select the sectors for each project
                wait = WebDriverWait(driver, 10)
                elementS = driver.find_element(By.ID, "ddlSector")
                sector_select = Select(elementS)
                sectorst = [option.text for option in sector_select.options][1:]
                sectors = [option.get_attribute('value') for option in sector_select.options][
                          1:]  # exclude the first option ('--Select--')

            
                for sector in sectors:
                    sector_select.select_by_value(sector)

                    # Submit the form
                    elementSub = driver.find_element(By.ID, "btnSubmit")
                    elementSub.click()

                    # Wait for the table to load
                    try:
                        table = wait.until(EC.presence_of_element_located((By.ID, "gvList")))
                    # If the table was found, continue with the rest of the code
                    except TimeoutException:
                        print("Table with ID 'gvList' not found. Skipping to next iteration...")
                        continue

                    # Get the table HTML and parse it with BeautifulSoup
                    html = table.get_attribute("innerHTML")
                    soup = bs(html, "html.parser")

                    # Find all table rows and extract the cells as a numpy array
                    rows = soup.find_all("tr")
                    arr = []
                    for row in rows:
                        cells = row.find_all("td")
                        arr.append(numpy.array(cells))
                    df = pd.DataFrame(arr, columns = [ 'AWC_ID', 'AWC_Name', 'Children(0to6M)', 'Measured (0to6M)', 'Children(7Mto3Y)', 'Measured (7Mto3Y)', 'Children(3Yto5Y)', 'Measured (3Yto5Y)'])
                    df[['AWC_ID', 'Children(0to6M)', 'Measured (0to6M)', 'Children(7Mto3Y)', 'Measured (7Mto3Y)', 'Children(3Yto5Y)', 'Measured (3Yto5Y)']] = df[['AWC_ID', 'Children(0to6M)', 'Measured (0to6M)', 'Children(7Mto3Y)', 'Measured (7Mto3Y)', 'Children(3Yto5Y)', 'Measured (3Yto5Y)']].fillna(0).astype(int)
                    df[['AWC_Name']] = df[['AWC_Name']].applymap(str)
                    filename = f"{month}_{year}_{district}_{project}_{sector}.xlsx"
                    df.to_excel(filename, index=False)
                    
                    # Go back to the main page
                    driver.back()
                    project_select = Select(driver.find_element(By.ID, "ddlProject"))
                    project_select.select_by_value(project)
                    sector_select = Select(driver.find_element(By.ID, "ddlSector"))
                    sector_select.select_by_value(sector)


