from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook

# Get Workbook and Sheetname
print("PhoneNumberGetter: Use this to get phone numbers and populate an excel sheet.")
workbookName = input("Workbook Name: ")
sheetName = input("Sheet Name: ")

# Open Workbook
workbook = load_workbook(workbookName)
sheet = workbook[sheetName]

# Set first Cell
i = 2
counter = 0
cell = sheet['A' + str(i)].value

# Construct URL for Search
def getURL(i):
    try: 
        query = (sheet['A' + str(i)].value + " " + sheet['B' + str(i)].value).split()
        url = "http://www.google.com/search?hl=en&q="
        for term in query:
            url += term + "%20"
        url += "phone%20number"
        return url
    except:
        raise TypeError("Check characters, error obtaining URL for")

# Execute Search with Chromium Agent
def runSearch(url):
    try: 
        hdr = {'User-Agent': 'Chrome/70.0.3538.77'}
        req = Request(url, headers=hdr)
        page = urlopen(req)
        return BeautifulSoup(page, "html.parser")
    except:
        raise LookupError("Check characters, unable to complete search for")

# Find phone number and save in doc
def returnNumber(i, searchSoup):
    try:
        number = searchSoup.find_all("div", attrs={"class": "BNeawe iBp4i AP7Wnd"})[1].getText()
        return number
    except:
        raise LookupError("Check terms, unable to find phone number from search for")

# Loop through column
while cell != None: 
    try:
        url = getURL(i)
        searchSoup = runSearch(url)
        number = returnNumber(i, searchSoup)
        sheet['E' + str(i)] = number
        print("Found for: " + cell)
        counter += 1
    except Exception as e:
        print(str(e) + ": " + cell)
    finally:
        # Iterate
        i += 1
        cell = sheet['A' + str(i)].value

# Save XLSX
workbook.save(workbookName)
print("Done! " + str(counter) + " phone numbers added...")