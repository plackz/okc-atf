### Script to measure the ATFs file every 15 minutes

import os
import datetime
import time
import csv
import re
from pprint import pprint
from operator import itemgetter
import win32com.client
import pandas as pd

while True:
    # create list of files in ATF folder to count
    path = 'I:\\Quality Control/After the Fact Documentation/ATF Sort'
    mn = 0
    custnameList=[]
    custnumList=[]
    custList=[]

    folders = ([name for name in os.listdir(path)
                 if os.path.isdir(os.path.join(path, name)) and name[0].isalpha()]) # get all directories 
    for folder in folders:
        contents = os.listdir(os.path.join(path,folder)) # get list of contents
        if len(contents) > mn: # if greater than the limit, print folder and number of contents
            custList.append(folder)
            custnumList.append(sum([len(files) for r, d, files in os.walk(path+'/'+folder)]))
    total=sum(custnumList)

    # get time stamp
    timeNow = datetime.datetime.now()

    # print message to screen
    print("As of: " + str(timeNow) + " there were " + str(total) + " ATFs in queue.")
    print("Saved to file.")
    print("")
    
    # write to csv file
    try:
        outputFile = open(r"I:\\Quality Control\\After the Fact Documentation\\ATF_time_study.csv","a", newline="")
        outputWriter = csv.writer(outputFile)
        outputWriter.writerow([timeNow, total])
        outputFile.close()
    except:
        outputFile = open(r"I:\\Quality Control\\After the Fact Documentation\\ATF_time_study_backup.csv","a", newline="")
        outputWriter = csv.writer(outputFile)
        outputWriter.writerow([timeNow, total])
        outputFile.close()
    
    time.sleep(1800) # sleep 30 minutes
