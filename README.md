
**How to Execute the code**

STEP 1  Download the folder

STEP 2  Open project in Pycharm and make sure it contains the following files to run the project successfully:

1. VISUAL_SINGLE.py
2. VISUAL_AUTO.py
3. VISUAL_COMPARE.py
4. Exchange rates f11hist-1969-2009.xls
5. Exchange rates f11hist.xls
6. GDP - 5204001_key_national_aggregates.xls
7. Tourism GDP 52490do001_201617.xls
8. 340101.xls
9. 340101.csv
10. 340102.xls
11. 340102.csv
12. 6302003.xls
13. 6302003.csv
14. f11-data.csv
15. mon_pax_web.csv 

STEP 3  Once downloaded - run the following files in pycharm, 

Run **VISUAL_SINGLE.py** - for single view of data (one page)

Run **VISUAL_AUTO.py** - 
this code will download the required data, converts and cleanses it before converting visualisations in .png files


**What each file does:**

**VISUAL_SINGLE.py** 

is a single python file to extract the data from manually downloaded files, cleanse and visualise the graphs.
It shows the data for the year 1997-98 to 2016-17 and was used for initial basic analysis

To run this file, make sure that the following data files are present in the folder

a. Exchange rates f11hist-1969-2009.xls - downloaded from https://www.rba.gov.au
b. Exchange rates f11hist.xls - downloaded from https://www.rba.gov.au
c. GDP - 5204001_key_national_aggregates.xls - downloaded from http://search.abs.gov.au
d. Tourism GDP 52490do001_201617.xls - downloaded from http://www.abs.gov.au

It produces the following individual graphs
1. Direct Tourism GDP
2. International visitors tourism GDP
3. Total GDP, current prices
4. Tourism share of total GDP
5. Average annual AUD/US exchange rate



**VISUAL_AUTO.py** 

is a single python file to extract the data automatically downloaded from the web sites,
cleanse and visualise the five graphs and save them as pdf files.
It shows and compares the data from Jan 2010 to Dec 2017 in five different graphical visualisations.

A. International passengers - Arrivals and Departures compared with the fluctuation in the value of AUD in USD

B. International passengers arrivals and departures in two separate graphs and shown by city to illustrate the popular destinations over time

C. Total international passenger arrivals compared with the number of residents and tourists

D. Total international passenger departures compared with the number of residents and tourists 

E. International passengers - Arrivals and Departures compared with average weekly earnings



**VISUAL_COMPARE.py** is a single python file which performs essentially the same function as VISUAL_AUTO.py but without automation.
It extracts the data from the manually downloaded files, cleanses and visualises the graphs
It shows and compares the data from Jan 2010 to Dec 2017

To run this file, need to make sure that the following data files are present in the folder
a. 340101.xls - downloaded from http://abs.gov.au
b. 340101.csv - converted and stored from xlsx
c. 340102.xls - downloaded from http://abs.gov.au
d. 340102.csv - converted and stored from xlsx
e. 6302003.xls - downloaded from http://www.abs.gov.au
f. 6302003.csv - converted and stored from xlsx
g. f11-data.csv - downloaded from https://www.rba.gov.au
h. mon_pax_web.csv - downloaded from https://data.gov.au


It visualises the following graphs in loop:

A. International passengers - Arrivals and Departures compared with the fluctuation in the value of AUD in USD

B. International passengers arrivals and departures in two separate graphs and shown by city to illustrate the popular destinations over time

C. Total international passenger arrivals compared with the number of residents and tourists

D. Total international passenger departures compared with the number of residents and tourists 

E. International passengers - Arrivals and Departures compared with average weekly earnings
