import urllib
####-Import URL library. This code pulls data from a website, NSE Stock information and creates an excel of top percentages
####-of the stock values in that day.
urlOfFileName = "https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR080519.zip"

urlOfFileName
####-Local Path; 'r' allows for readable directory with using '//' Just Copy & Paste
localZipFilePath = r"C:\Users\Mike\Desktop\Python\Data_NSE/PR080519.zip"

localZipFilePath
#### mask your robots/internet-using-script. hdr makes a User Agent and allows for data extrap.
hdr = {'User Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36', 
       'Accept': 'text/html, application/xhtml+xml,application/xml;q=0.9,*/*q=0.8',
       'Accept-Charset':'ISO-8859-1;utf-8,q=0.7,*;q=0.3',
       'Accept-Encoding':'none',
       'Accept-Language':'en-US,en;q=0.8',
       'Connection':'keep-alive'
      }

hdr

webRequest = urllib.request.Request(urlOfFileName)
####-Try being used for safety. Exception pulled if any issues arise.
try:
    page = urllib.request.urlopen(webRequest)
    content = page.read()
    output = open(localZipFilePath,"wb")
    output.write(bytearray(content))
    output.close()
except(urllib.request.HTTPError, e):
    print(e.fp.read())
    print("File did not pass for url = ",urlOfFileName)

import zipfile, os
####-Zipfile Library allows us to extract the zip file and save it to a new folder with extracted data.
localExtractFilePath = r"C:\Users\Mike\Desktop\Python\Data_NSE\\"
####- If first checks whether zip is local then comes the extraction.
if os.path.exists(localZipFilePath):
    print("Cool! the file exists!")
    listOfFiles = []
    fh = open(localZipFilePath, "rb")
    zipFileHandler = zipfile.ZipFile(fh)
    for filename in zipFileHandler.namelist():
        zipFileHandler.extract(filename, localExtractFilePath)
        listOfFiles.append(localExtractFilePath + filename)
        print("Extracted " + filename + "from the zip file.")
    print("In total, we extracted ", str(len(listOfFiles)), " files")
    fh.close()
####- Datasets come in CSV format, import CSV format to manipulate. This code is pulling the 7th file from the folder.
import csv

oneFileName = listOfFiles[6]
lineNum = 0
listOfLists = []
with open(oneFileName, "r") as csvfile:
    lineReader = csv.reader(csvfile,delimiter=",", quotechar="\"")
    for row in lineReader:
        lineNum = lineNum + 1
        if lineNum == 1:
            print("Skipping the header row")
            continue
        symbol = row[1]
        close = row[2]
        prevClose = row[3]
        pctChange = row[4]
        oneResultRow = [symbol, pctChange]
        listOfLists.append(oneResultRow)
    print("Finished iterating over the file contents - the file is closed now!")
    print("We have stock info for " + str(len(listOfLists)) + " stocks")
####-Lambda function to sort the list of lists by column 2 and putting it in decending order.
listOfListsSortedByQty = sorted(listOfLists, key=lambda x:x[1], reverse=True)

listOfListsSortedByQty
####- Changing our CVS file to excel. Then only looking at the highest percentage change. 
import xlsxwriter

excelFileName = r"C:\Users\Mike\Desktop\Python\Data_NSE\GI080519.xlsx"

workbook = xlsxwriter.Workbook(excelFileName)
worksheet = workbook.add_worksheet("Summary")

worksheet.write_row("A1", ["Top Traded Stocks"])
worksheet.write_row("A2", ["Stock", "% Change", "Value Traded (INR)"])

for rowNum in range(5):
    oneRowToWrite = listOfListsSortedByQty[rowNum]
    worksheet.write_row("A" + str(rowNum + 3), oneRowToWrite)
workbook.close()
