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
cell = sheet['A' + str(i)].value

# Loop through column
while cell != None:
    
    # Construct URL for Search
    query = (sheet['A' + str(i)].value + " " + sheet['B' + str(i)].value).split()
    url = "http://www.google.com/search?hl=en&q="
    for term in query:
        url += term + "%20"
    url += "phone%20number"

    # Execute Search with Mozilla Agent
    hdr = {'User-Agent': 'Chrome/70.0.3538.77'}
    req = Request(url, headers=hdr)
    page = urlopen(req)
    soup = BeautifulSoup(page, "html.parser")

    # FOR DEBUG
    #print(soup.prettify())

    # Find phone number and save in doc
    number = soup.find("div", attrs={"class": "BNeawe iBp4i AP7Wnd"}).getText()
    sheet['E' + str(i)] = number

    # Iterate
    i += 1
    cell = sheet['A' + str(i)].value

# Save XLSX
workbook.save('test.xlsx')
print("Done! " + str(i-2) + " phone numbers added...")