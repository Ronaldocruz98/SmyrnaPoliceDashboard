#Importing Python Libraries that will be of use for my scrape
import pyodbc
import sqlalchemy
from sqlalchemy import create_engine
import csv
import time
import pandas as pd
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.keys import Keys

# This bit of code here is used to launch our instance of Chrome that is compatible with Selenium and going straight to the site I want to scrape
PATH = r"C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)
driver.get("https://p2c.townofsmyrna.org/cad/callsnapshot.aspx")

#testing my code to see if it prints out the title of the page
print(driver.title)

# right here I am telling the program to select the 200 records on the page.
driver.find_element_by_class_name('ui-pg-selbox').click()
# noinspection PyDeprecation
driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/div[2]/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[5]/select/option[5]').click()
time.sleep(1)


## I then have the program open a new CSV (excel) file so it can start write data that it scrapes
## I am declaring "rows" and writing a for loop so that it may start to loop through all of the rows in the table.
## It then is taking each element of the whole row is putting them in their own cells in the CSV file until the loop ends.

f = csv.writer(open('SPDJuypter12.csv', 'w', newline=''))
f.writerow(['Agency','Service','Case_Num','Start_Time', 'End_Time','Nature','Address'])
rows = driver.find_elements_by_xpath("/html/body/form/table/tbody/tr[2]/td/div[2]/div/div[3]/div[3]/div/table/tbody/tr")
Agency =[]
for row in rows:
    colagency = row.find_elements_by_tag_name("td")[0]
    colservice= row.find_elements_by_tag_name("td")[1]
    colcase = row.find_elements_by_tag_name("td")[2]
    colstartime = row.find_elements_by_tag_name("td")[3]
    colendtime = row.find_elements_by_tag_name("td")[4]
    colnature = row.find_elements_by_tag_name("td")[5]
    coladdress = row.find_elements_by_tag_name("td")[6]
    Agency = colagency.text
    Service = colservice.text
    CaseNum = colcase.text
    Startime = colstartime.text
    Endtime = colendtime.text
    Nature = colnature.text
    Address = coladdress.text

    f.writerow([Agency, Service, CaseNum, Startime, Endtime, Nature, Address])

time.sleep(2)
driver.quit()

## I am incorporating Pandas and asking it to read my CSV file
df = pd.read_csv('SPDJuypter12.csv', encoding='cp1252')

#I am splitting the  Start/End Date/Time and creating a seperate DATE and TIME column
df[['Start_Date', 'Start_Time']] = df["Start_Time"].str.split(" ", 1, expand=True)
df[['End_Date', 'End_Time']] = df["End_Time"].str.split(" ", 1, expand=True)
print(df.columns)
# Here I re arranged the Data Frame columns to include the newly separated Start/End Dates/Times thus creating a new Dataframe.
new_df = df[['Agency','Service','Case_Num', 'Start_Date','Start_Time','End_Date','End_Time', 'Nature', 'Address',]]
new_df
## Creating a new CSV file with current date and with the newly arrange Dataframe.
date=datetime.now().strftime('%Y%m%d%H')
new_df.to_csv(f'SPD_CallRecords{date}.csv')

# Here I will connect to my Database and upload the Dataframe to my Database
engine = sqlalchemy.create_engine('mssql://DESKTOP-MRQUA3V/Smyrna Police Calls''?driver=ODBC+Driver+17+for+SQL+Server')
data = new_df
data.to_sql('PoliceCalls',engine, if_exists='append')
