""" vtu web scraper"""

import bs4    
import requests
from lxml import html
from tkinter import *
import xlsxwriter
import re

resArray = {}

def outputToExcel(fname):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(fname)
    worksheet = workbook.add_worksheet()
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    # Iterate over the data write it out row by row.
    for myRow in (resArray.items()):
        col=0
        for key in (myRow[1].keys()):
            worksheet.write(row, col, key)
            col +=1
        break;

    row=row+1

    for myRow in (resArray.items()):
        col=0
        for key in (myRow[1].keys()):
            worksheet.write(row, col, myRow[1][key])
            col=col+1
        row += 1

    workbook.close()
    print(row )
    print(" rows successfully written to file")



def fetch_reg_results(soup, sem):
        table =  soup.find_all('table')
        row = table[0].find_all('tr')
        columns = row[0].find_all('td')
        sUsn =  columns[1].get_text()
        columns = row[1].find_all('td')
        sName = columns[1].get_text()
        mydict={}
        resArray[sUsn] = {}

        resRows = soup.find_all("div",{"class":"row" })

        if("Semester : "+sem in resRows[6].get_text()):
            divTableRows = resRows[6].find_all("div", {"class": "divTableRow"})
            i=0
            for row in divTableRows:
                i=i+1
                if(i==1):
                    continue
                if(i>10):
                    break
                tdTags  = row.find_all("div", {"class": "divTableCell"})

                if(len(tdTags)<6):
                    break

                #mydict[tdTags[0].get_text()] = tdTags[3].get_text()
                mydict[tdTags[0].get_text()+"_IA"] = tdTags[2].get_text()
                mydict[tdTags[0].get_text()+"_EXT"] = tdTags[3].get_text()
                mydict[tdTags[0].get_text()+"_Total"] = tdTags[4].get_text()
                mydict[tdTags[0].get_text()+"_XRes"] = tdTags[5].get_text()

                resArray[sUsn]['USN'] = sUsn
                resArray[sUsn]['NAME'] = sName

            for key in sorted(mydict.keys()):
                resArray[sUsn][key] = mydict[key]



def fetch_reval_results(soup):
    table =  soup.find_all('table')
    row = table[0].find_all('tr')
    columns = row[0].find_all('td')
    sUsn =  columns[1].get_text()
    columns = row[1].find_all('td')
    sName = columns[1].get_text()
    divTag = soup.find_all("div", {"class": "divTableRow"})
    mydict = {}
    resArray[sUsn] = {}
    resArray[sUsn]['USN'] = sUsn
    resArray[sUsn]['NAME'] = sName
    i=0
    for tag in divTag:
        i=i+1
        if(i==1):
            continue
        if(i>9):
            break
        tdTags  = tag.find_all("div", {"class": "divTableCell"})

        if(re.match("^\d\d\w*",tdTags[0].get_text())): #(tdTags[0].get_text() != "P -> PASS"):
            mydict[tdTags[0].get_text()] = tdTags[0].get_text()+":"+ tdTags[3].get_text()+"=>"+tdTags[2].get_text()+"("+tdTags[5].get_text()+")"

        for key in sorted(mydict.keys()):
            resArray[sUsn][key] = mydict[key]



def fetch_results():
    resType = val1.get()
    scheme = val2.get()
    pat = e3.get()
    usnStart = e4.get()
    usnEnd =e5.get()
    sem = e6.get()
    if(resType=='regular'):
        if(scheme == 'Non-CBCS'):
            cur_url='http://results.vtu.ac.in/vitaviresultnoncbcs18/index.php'
            URL = 'http://results.vtu.ac.in/vitaviresultnoncbcs18/resultpage.php'
            token = 'bjV3V2pBV25HaVdnaXhqeVBzMytQaGVIalV6ZWhwNVYwajUwT1A4bXpieTA1QWl3RG9LTG53bFNvN2NjVTJmUTlMaWNnS0c5R2FQRmQ5RUFDcTl0SlE9PTo6BNtJzyc/sD2gU6paXNPJbg=='
        else:
            cur_url = 'http://results.vtu.ac.in/vitaviresultcbcs2018/index.php'
            URL ='http://results.vtu.ac.in/vitaviresultcbcs2018/resultpage.php'
            token = 'a2xGNVFxcFJQTzFXcTdFdVRxV1hmaXEyb2FINFpHd0FjV1ZwR1JMSTJqZDFxM3BOVUduK0h0bXFmeFJEYWhmRnFlMFNYdUs4QmxFWjM3VThiUTdoZnc9PTo6Y3hidx07r2YwRo63SN37kw=='
    else:
        if(scheme == 'Non-CBCS'):
            URL = 'http://results.vtu.ac.in/vitavirevalresultnoncbcs2018/resultpage.php'
        else:
            URL ='http://results.vtu.ac.in/vitavirevalresultcbcs2018/resultpage.php'




    for i in range(int(usnStart),int(usnEnd)):
        usn = pat+'00'+ str(i)
        if(len(usn)>10):
            usn =pat+'0'+str(i)
        if(len(usn)>10):
            usn= pat+ str(i)
        PARAMS = {'lns': usn, 'token':token, 'current_url':cur_url}
        r = requests.post(url = URL, data = PARAMS) #data --> posted data
        soup = bs4.BeautifulSoup(r.text,"html.parser")
        if(resType=='revals'):
            fetch_reval_results(soup)
        else:
            fetch_reg_results(soup, sem)


    outputToExcel('vtu_res.xlsx');






master = Tk()
master.title("VTU results")

# Create a Tkinter variable
val1 = StringVar(master)
val2 = StringVar(master)

# Dictionary with options
choice1 = { 'regular','revals'}
val1.set('regular') # set the default option

e1 = OptionMenu(master, val1, *choice1)
Label(master, text="Result Type:").grid(row = 0, column = 0)
e1.grid(row = 0, column =1)

# Dictionary with options
choice2 = { 'CBCS','Non-CBCS'}
val2.set('CBCS') # set the default option

e2 = OptionMenu(master, val2, *choice2)
Label(master, text="Scheme").grid(row = 1, column = 0)
e2.grid(row = 1, column =1)

Label(master, text="Enter USN format [4so14cs]:").grid(row=2)
Label(master, text="Enter usn start val:").grid(row=3)
Label(master, text="Enter usn end val:").grid(row=4)
Label(master, text="Enter Semester:").grid(row=5)

e3 = Entry(master)
e4 = Entry(master)
e5 = Entry(master)
e6 = Entry(master)

e3.grid(row=2, column=1)
e4.grid(row=3, column=1)
e5.grid(row=4, column=1)
e6.grid(row=5, column=1)


Button(master, text='Quit', command=master.quit).grid(row=6, column=0, sticky=W, pady=4)
Button(master, text='Generate Results', command=fetch_results).grid(row=6, column=1, sticky=W, pady=18)

mainloop( )

