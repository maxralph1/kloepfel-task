import os
import time
import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#
import xml.etree.ElementTree as ET
from lxml import etree

#
from openpyxl import Workbook


# Get the current working directory
current_directory = os.getcwd() 
# current_directory = Path.cwd()
print("Current Directory:", current_directory)

options = webdriver.ChromeOptions()
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": current_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    },
)

# Initialize Web Driver
driver = webdriver.Chrome(options=options)


# Visit the first page
url = "https://www.handelsregister.de/rp_web/welcome.xhtml"
# url = "https://www.handelsregister.de/rp_web/erweitertesuche.xhtml"
driver.get(url)

# Click on the 'Advanced Search' button
advanced_search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "naviForm:erweiterteSucheLink"))
)

advanced_search_button.click()

# Wait for the next page to load
time.sleep(10)

# Fill out the form
schlagwoerter = driver.find_element(By.NAME, "form:schlagwoerter")
schlagwoerter.send_keys("Kloepfel Consulting GmbH")

#
# schlagwortOptionen Section

try:
    # # Click on the "schlagwortOptionen_exact" button
    # schlagwortOptionen_exact = driver.find_element(By.ID, "form:schlagwortOptionen:2")

    # Wait until the element is visible and clickable
    schlagwortOptionen_exact = WebDriverWait(driver, 10).until(
        # EC.element_to_be_clickable(
        #     (
        #         By.ID("form:schlagwortOptionen:2"),
        #         By.NAME("form:schlagwortOptionen"),
        #     )
        # )
        EC.element_to_be_clickable(
            (By.XPATH, '//label[@for="form:schlagwortOptionen:2"]')
        )
        # EC.element_to_be_clickable((By.ID, "form:schlagwortOptionen:2"))
        # EC.element_to_be_clickable(
        #     (By.LINK_TEXT, "contain the exact name of the company.")
        # )
        # EC.element_to_be_clickable((By.LINK_TEXT, "den genauen Firmennamen enthalten."))
    )

    # Scroll the "schlagwortOptionen_exact" button into view
    driver.execute_script("arguments[0].scrollIntoView();", schlagwortOptionen_exact)

    # Move the mouse over the element
    actions = ActionChains(driver)
    actions.move_to_element(schlagwortOptionen_exact).perform()

    # Click on it
    schlagwortOptionen_exact.click()

except selenium.common.exceptions.ElementNotInteractableException:
    print("Element not interactable, trying JavaScript click.")
    driver.execute_script("arguments[0].click();", schlagwortOptionen_exact)

except selenium.common.exceptions.TimeoutException:
    print("Element not found or not clickable within the wait time.")


#
# Suchen Button Section

try:

    # Locate the "Suchen" button and submit the form
    find_button = driver.find_element(By.NAME, "form:btnSuche")

    # Scroll the "Suchen" button into view
    driver.execute_script("arguments[0].scrollIntoView();", find_button)

    # Move the mouse over the element
    actions = ActionChains(driver)
    actions.move_to_element(find_button).perform()

    find_button.click()

except selenium.common.exceptions.ElementClickInterceptedException:
    print("Click intercepted, trying JavaScript click.")
    driver.execute_script("arguments[0].click();", find_button)

except selenium.common.exceptions.TimeoutException:
    print("Element not found or not clickable within the wait time.")

finally:
    pass


# Wait again for the next page to load
time.sleep(10)

# Download the "SI" xhtml file
download_button = driver.find_element(By.LINK_TEXT, "SI")
download_button.click()

# Wait for the download to complete
time.sleep(10)


# Verify download
# Get the latest downloaded file name
files = os.listdir(current_directory)
files = [f for f in files if not f.endswith(".crdownload")]  # Exclude incomplete files
files.sort(key=lambda x: os.path.getctime(os.path.join(current_directory, x)))
latest_file = files[-1]

print(f"Downloaded file: {latest_file}")

if os.name == "nt":
    print(
        f"Downloaded file directory (Windows notation): {current_directory + '\\' + latest_file}"
    )
elif os.name == "posix":
    print(
        f"Downloaded file directory (Linux-like notation): {current_directory + '/' + latest_file}"
    )
else:
    print(f"Downloaded file directory: {current_directory + '/' + latest_file}")


# Close the browser
driver.quit()


#
#
# Traversing the latest downloaded XML file and extracting data
#
#

tree = ET.parse(current_directory + "\\" + latest_file)
root = tree.getroot()

namespaces = {"tns": "http://www.xjustiz.de"}

# erstellungszeitpunkt = root.find(
#     "tns:nachrichtenkopf/tns:erstellungszeitpunkt", namespaces
# ).text
# print(f"Erstellungszeitpunkt: {erstellungszeitpunkt}")

# # Access the 'code' inside 'absender.gericht'
# code = root.find(
#     "tns:nachrichtenkopf/tns:auswahl_absender/tns:absender.gericht/code", namespaces
# ).text
# print(f"Code: {code}")


company_name = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Company Name: {company_name}")

company_info = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Company Info: {company_info}")

city = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:sitz/tns:ort",
    namespaces,
).text
print(f"City: {city}")

status = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Status: {status}")

bezeichnung = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Bezeichnung: {bezeichnung}")

rechtsform = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Rechtsform: {rechtsform}")

straße = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:anschrift/tns:strasse",
    namespaces,
).text
print(f"Straße: {straße}")

hausnummer = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:anschrift/tns:hausnummer",
    namespaces,
).text
print(f"Hausnummer: {hausnummer}")

postleitzahl = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:anschrift/tns:postleitzahl",
    namespaces,
).text
print(f"Postleitzahl: {postleitzahl}")

ort = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:anschrift/tns:ort",
    namespaces,
).text
print(f"Ort: {ort}")

rechtsform = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Company Name: {rechtsform}")

rechtsform = root.find(
    "tns:grunddaten/tns:verfahrensdaten/tns:beteiligung/tns:beteiligter/tns:auswahl_beteiligter/tns:organisation/tns:bezeichnung/tns:bezeichnung.aktuell",
    namespaces,
).text
print(f"Company Name: {rechtsform}")


#
#
# Place the extracted info in an Excel Sheet
#
#
# Create a new Excel workbook
workbook = Workbook()
sheet = workbook.active

# Write headers
# sheet["A1"] = "Erstellungszeitpunkt"
# sheet["B1"] = "Code"
sheet["A1"] = "Company Name"
sheet["B1"] = "Company Info"
sheet["C1"] = "City"
sheet["D1"] = "Status"
sheet["E1"] = "Bezeichnung"
sheet["F1"] = "Rechtsform"
sheet["G1"] = "Straße"
sheet["H1"] = "Hausnummer"
sheet["I1"] = "Postleitzahl"
sheet["J1"] = "Ort"
sheet["K1"] = "Geschäftsführer(in) Vorname"
sheet["L1"] = "Geschäftsführer(in) Nachname"
sheet["M1"] = "Geschäftsführer(in) Geschlecht"
sheet["N1"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O1"] = "Gegenstand"
sheet["P1"] = "Vertretungsbefugnis"

# Write the extracted data
# sheet["A2"] = erstellungszeitpunkt
# sheet["B2"] = code
sheet["A2"] = company_name
sheet["B2"] = "Company Info"
sheet["C2"] = "City"
sheet["D2"] = "Status"
sheet["E2"] = "Bezeichnung"
sheet["F2"] = "Rechtsform"
sheet["G2"] = "Straße"
sheet["H2"] = "Hausnummer"
sheet["I2"] = "Postleitzahl"
sheet["J2"] = "Ort"
sheet["K2"] = "Geschäftsführer(in) Vorname"
sheet["L2"] = "Geschäftsführer(in) Nachname"
sheet["M2"] = "Geschäftsführer(in) Geschlecht"
sheet["N2"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O2"] = "Gegenstand"
sheet["P2"] = "Vertretungsbefugnis"
#
sheet["A3"] = "Company Name"
sheet["B3"] = "Company Info"
sheet["C3"] = "City"
sheet["D3"] = "Status"
sheet["E3"] = "Bezeichnung"
sheet["F3"] = "Rechtsform"
sheet["G3"] = "Straße"
sheet["H3"] = "Hausnummer"
sheet["I3"] = "Postleitzahl"
sheet["J3"] = "Ort"
sheet["K3"] = "Geschäftsführer(in) Vorname"
sheet["L3"] = "Geschäftsführer(in) Nachname"
sheet["M3"] = "Geschäftsführer(in) Geschlecht"
sheet["N3"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O3"] = "Gegenstand"
sheet["P3"] = "Vertretungsbefugnis"
#
sheet["A4"] = "Company Name"
sheet["B4"] = "Company Info"
sheet["C4"] = "City"
sheet["D4"] = "Status"
sheet["E4"] = "Bezeichnung"
sheet["F4"] = "Rechtsform"
sheet["G4"] = "Straße"
sheet["H4"] = "Hausnummer"
sheet["I4"] = "Postleitzahl"
sheet["J4"] = "Ort"
sheet["K4"] = "Geschäftsführer(in) Vorname"
sheet["L4"] = "Geschäftsführer(in) Nachname"
sheet["M4"] = "Geschäftsführer(in) Geschlecht"
sheet["N4"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O4"] = "Gegenstand"
sheet["P4"] = "Vertretungsbefugnis"
#
sheet["A5"] = "Company Name"
sheet["B5"] = "Company Info"
sheet["C5"] = "City"
sheet["D5"] = "Status"
sheet["E5"] = "Bezeichnung"
sheet["F5"] = "Rechtsform"
sheet["G5"] = "Straße"
sheet["H5"] = "Hausnummer"
sheet["I5"] = "Postleitzahl"
sheet["J5"] = "Ort"
sheet["K5"] = "Geschäftsführer(in) Vorname"
sheet["L5"] = "Geschäftsführer(in) Nachname"
sheet["M5"] = "Geschäftsführer(in) Geschlecht"
sheet["N5"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O5"] = "Gegenstand"
sheet["P5"] = "Vertretungsbefugnis"
#
sheet["A6"] = "Company Name"
sheet["B6"] = "Company Info"
sheet["C6"] = "City"
sheet["D6"] = "Status"
sheet["E6"] = "Bezeichnung"
sheet["F6"] = "Rechtsform"
sheet["G6"] = "Straße"
sheet["H6"] = "Hausnummer"
sheet["I6"] = "Postleitzahl"
sheet["J6"] = "Ort"
sheet["K6"] = "Geschäftsführer(in) Vorname"
sheet["L6"] = "Geschäftsführer(in) Nachname"
sheet["M6"] = "Geschäftsführer(in) Geschlecht"
sheet["N6"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O6"] = "Gegenstand"
sheet["P6"] = "Vertretungsbefugnis"
#
sheet["A7"] = "Company Name"
sheet["B7"] = "Company Info"
sheet["C7"] = "City"
sheet["D7"] = "Status"
sheet["E7"] = "Bezeichnung"
sheet["F7"] = "Rechtsform"
sheet["G7"] = "Straße"
sheet["H7"] = "Hausnummer"
sheet["I7"] = "Postleitzahl"
sheet["J7"] = "Ort"
sheet["K7"] = "Geschäftsführer(in) Vorname"
sheet["L7"] = "Geschäftsführer(in) Nachname"
sheet["M7"] = "Geschäftsführer(in) Geschlecht"
sheet["N7"] = "Geschäftsführer(in) Geburtsdatum"
sheet["O7"] = "Gegenstand"
sheet["P7"] = "Vertretungsbefugnis"

# Save the workbook
workbook.save("kloepfel_task_goal_output.xlsx")
