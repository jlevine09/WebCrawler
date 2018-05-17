from bs4 import BeautifulSoup
import mechanize
import requests
import webbrowser
import urllib
import string
import os
import xlwt

#Excel
wb = xlwt.Workbook()
url = formhandler(datacode)
r = requests.get(url, auth=('user', 'pass'))
data = r.text
soup = BeautifulSoup(data)
i = 0
x = 0
y = 0

#FILE HANDLING
fp = open('test.xls', 'w')


def excelWorkbook(sheetnum, code, x, datacode):
    print(r.status_code)
    i = 0
    word = ""
    y = 0
    table = soup.find('table', {'class': "ciwqsReportDataTable"})
    tdList0 = table.find_all('tr')[0].text
    tdList = table.find_all('tr')[code].text
    
    sheetnum = wb.add_sheet(sheetnum, cell_overwrite_ok=True)
    
    tdList0 = tdList0[1: ]
    tdList = tdList[1: ]

    while (i in range(0, 25)):

        for td0 in tdList0:
            if (td0 != '\n'):
                word += td0
                
            else:
                texte_bu = td0.encode('utf-8')
                texte_bu = texte_bu.strip()
                sheetnum.write(0, i, word)
                print(0, i, word)
                word = ""
                i = i + 1
    
        i = 0
        for td in tdList:
            if (td != '\n'):
                word += td
                
            else:
                texte_bu = td.encode('utf-8')
                texte_bu = texte_bu.strip()
                sheetnum.write(x, i, word)
                print(0, i, word)
                word = ""
                i = i + 1
    wb.save("Test.xls")    
    
def main():
    table = soup.find('table', {'class': "ciwqsReportDataTable"})
    tdList = table.find_all('tr')
    i = 1
    
    print("Welcome to the WTFS. Please type in a region code.")
    print("1: North Coast")
    print("2: San Francisco Bay")
    print("3: Central Coast")
    print("4: Los Angeles")
    print("5F: Central Valley, Fresno Office")
    print("5R: Central Valley, Redding Office")
    print("5S: Central Valley, Sacramento Office")
    print("6T: Lahontan, Tahoe Office")
    print("6V: Lahontan, Victorville Office")
    print("7: Colorado River")
    print("8: Santa Ana")
    print("9: San Diego\n")
    
    datacode = raw_input("Enter code: ")

    for td in range(0, len(tdList) - 1):
        sheetnum = 'ws' + str(i)
        excelWorkbook(sheetnum, i, 1, datacode)
        i = i + 1
    print("The excel document was saved as Test.xls.")
    print("Thank you and have a nice day.\n")

    return 0

main()
