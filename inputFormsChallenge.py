# Import dependancies
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
# Import classes
from classes import Excelfile, Information, Form

# Chrome parameters definition
# projectDownloadFolder definition and default download directory configuration
projectFolder = os.path.dirname(os.path.abspath(__file__))
projectDownloadFolder = projectFolder
chrome_options = Options()
prefs = {"download.default_directory": projectDownloadFolder}
chrome_options.add_experimental_option("prefs", prefs)
# Suppress unnecessary browser logs
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])


# Browser initialization
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://rpachallenge.com/")

# Maximize the window
driver.maximize_window()

try:
    # Reject cookies at the opening of Chrome
    reject_button = WebDriverWait(driver, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//button[.//div[contains(text(), 'efuser')]]")) # only the end of 'Tout refuser' (reject all in french) to capture different possible cases
    )
    reject_button.click()
    print("Cookies rejected.")
except:
    print("No consent window detected.")

excelFile = Excelfile("challenge.xlsx", projectDownloadFolder)
excelFile.download()

# Extract information from excel file
df = excelFile.readFile()

# Assign each line of the df in list_Information_objects
# each item of the list is an instance of the Information class
list_Information_objects = []

for i in range(len(df)):
    # Create an instance of the Information class
    instanceInformation = Information(
        df.iloc[i,0],
        df.iloc[i,1],
        df.iloc[i,2],
        df.iloc[i,3],
        df.iloc[i,4],
        df.iloc[i,5],
        df.iloc[i,6],
    )
    list_Information_objects.append(instanceInformation)

# Click on Start button
element = WebDriverWait(driver, 3).until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Start')]")) 
)
element.click()

# Create an instance of the Form class
instanceForm = Form()

# Main Loop
for i in range(len(df)):
    instanceInformation = list_Information_objects[i]
    for field in ['firstName', 'lastName', 'companyName', 'roleInCompany', 'address', 'email', 'phone']:
        element = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.XPATH, instanceForm.selectors[field]))
        )
        element.send_keys(getattr(instanceInformation, field)) # use of getattr to dynamically retrieve attribut of the object instanceInformation

    # Click on Submit button after the end of the inner loop
    element = WebDriverWait(driver, 1).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@value='Submit']"))
    )
    element.click()

time.sleep(2)

# Close browser
driver.quit()

# Delete excel file
excelFile.delete()