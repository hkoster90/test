import os
import csv
import pandas as pd
from datetime import datetime, timedelta
import codecs
import openpyxl

def transformOT(inputFile, outputfile):
    #init and open inputFile
    ot_file = open(inputFile,'r')

    #init lists for further usage
    campaigns = []
    shortlogins = []
    dates = []
    workhours = []

    #looping through the file and store each column per row of the file into the lists above
    with ot_file as csv_file:
        csvData = csv.reader(csv_file)
        #skip header row
        next(csvData)
        for line in csvData:
            campaign = line[0]
            shortlogin = line[1]
            date = line[2]

            starttimeString = f"{date} {line[3]}"
            endtimeString = f"{date} {line[4]}"

            d1 = datetime.strptime(starttimeString, "%d/%m/%Y %H:%M:%S")
            d2 = datetime.strptime(endtimeString, "%d/%m/%Y %H:%M:%S")

            #checking if endtime is smaller than starttime, if yes add 1 day to correct
            if d2 < d1:
                d2 = d2 + timedelta(days=1)

            #get the hours timerange result in hours to know how many rows have to be added
            delta = (d2 - d1).total_seconds()/3600
            #start the for loop to write each hour in a line until endtime is reached
            for i in range(int(delta)):
                    #add 1 hour to the starttime
                    workhourPrepare = d1 + timedelta(hours=i)
                    #converte datetime object into string to be able to write into excel
                    exportDate = workhourPrepare.strftime("%d/%m/%Y")
                    #extract the hour as integer to write into workhour column
                    workhour = int(workhourPrepare.strftime("%H"))
                    
                    #append the results of above into the lists of beginn of the function
                    campaigns.append(campaign)
                    shortlogins.append(shortlogin)
                    dates.append(exportDate)
                    workhours.append(workhour)

    #close the input file to free up memory
    ot_file.close()

    #write a dict to use in pandas
    data = {'campaign': campaigns, 'shortlogin': shortlogins, 'date': dates, 'workhour': workhours}

    #create pandas dataframe
    df = pd.DataFrame(data)
    #show dataframe in console
    print(df)
    #export new CSV for backup purpose in correct format
    df.to_csv('ot_test_drilled_down.csv', index=False, header=True, encoding='utf-8')

    sheetname = 'OT TEST'
    with pd.ExcelWriter(outputfile) as writer:
        #remove blanks from shortlogin: shortlogins should have only 4 chars
        df['shortlogin'] = df['shortlogin'].str.replace(" ", "")
        #export dataframe to excel
        df.to_excel(writer, sheet_name=sheetname, index=False)

    #additional (not mandatory but nice to see)
    #transform output excelsheet into excel table
    wb = openpyxl.load_workbook(filename = outputfile)
    tab = openpyxl.worksheet.table.Table(displayName="ot-test", ref=f'A1:{chr(len(df.columns)+64)}{len(df)+1}')
    wb[sheetname].add_table(tab)
    wb.save(outputfile)

if __name__ == '__main__':
    #define variables for the function calling. place the inputfile into same folder as your script
    inputFile = "ot_test.csv"
    outputfile = "ot_test.xlsx"

    transformOT(inputFile, outputfile)
