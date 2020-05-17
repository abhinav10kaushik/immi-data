# This program produces a visualisation of economic conditions in Australia, including
# in the tourism industry, for the period 1997-98 to 2016-17.

import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import xlrd

############################################
#  Step 1 - Extract and transform data     #
############################################

############# Tourism GDP data #############

# Assign file handle and open the workbook

Tourism_GDP_Excel = "Tourism GDP 52490do001_201617.xls"
xl_workbook = xlrd.open_workbook(Tourism_GDP_Excel)

# List sheet names, and pull a sheet by name

sheet_names = xl_workbook.sheet_names()

Table1_sheet = xl_workbook.sheet_by_name(sheet_names[1])  # Table 1 includes Direct tourism GDP and share of total GDP
Table2_sheet = xl_workbook.sheet_by_name(sheet_names[2])  # Table 2 includes Direct tourism GDP for international
                                                          # visitors
# Pull the required rows by index

row_year = Table1_sheet.row(4)              # 5th row, Table 1 - references year
row_TGDP = Table1_sheet.row(8)              # 9th row, Table 1 - references Direct tourism GDP in $m
row_share_GDP = Table1_sheet.row(13)        # 14th row, Table 1 - references Tourism share of total GDP in $m
row_Int_TGDP = Table2_sheet.row(11)         # 12th row, Table 2 - references international visitors tourism GD in $m

# Create lists of values from each row

year = []
Tourism_GDP = []
Tourism_share_GDP = []
Int_TGDP = []

for idx, cell_obj in enumerate(row_year):
    year.append(cell_obj.value)

del(year[0]) # we don't want the first cell which is blank

for idx, cell_obj in enumerate(row_TGDP):
    Tourism_GDP.append(cell_obj.value)

del(Tourism_GDP[0]) # we don't want the first cell which is the label

for idx, cell_obj in enumerate(row_share_GDP):
    Tourism_share_GDP.append(cell_obj.value)
del(Tourism_share_GDP[0]) # we don't want the first cell which is the label

for idx, cell_obj in enumerate(row_Int_TGDP):
    Int_TGDP.append(cell_obj.value)
del(Int_TGDP[0]) # we don't want the first cell which is the label

############# GDP current prices data #############

# Assign file handle for GDP current prices data and open workbook

GDP_excel = "GDP - 5204001_key_national_aggregates.xlsx"
GDP_workbook = xlrd.open_workbook(GDP_excel)

# List sheet names, and pull a sheet by name

sheet_names = GDP_workbook.sheet_names()

GDP_data_sheet = GDP_workbook.sheet_by_name(sheet_names[1])  # Sheet with GDP data

# Pull the required column by index (GDP current prices in $m)

column_GDP_cp = GDP_data_sheet.col(25)  # 'Z' col - GDP current prices

# Add values from relevant columns and subset to required range (year ending June 1998-  year endingJune 2017)

GDP_cp = []

for idx, cell_obj in enumerate(column_GDP_cp):
    if idx >= 48:
        GDP_cp.append(cell_obj.value)

############# Exchange rate data #############

# Assign file handle for early exchange rate data and open workbook

exchange_rate_excel = "Exchange rates f11hist-1969-2009.xls"
er_workbook = xlrd.open_workbook(exchange_rate_excel)

# List sheet names, and pull a sheet by name

sheet_names = er_workbook.sheet_names()

exchange_rate_sheet = er_workbook.sheet_by_name(sheet_names[0])  # Sheet with exchange rate data

# Pull the required column by index

column_ER_aud_us = exchange_rate_sheet.col(2)  # 'C' col - AUD/US exchange rate data

# Add values from July 1997 onwards

ER_aud_us = []

for idx, cell_obj in enumerate(column_ER_aud_us):
    if idx >= 347:  # from 19978-98 to Dec 09
        ER_aud_us.append(cell_obj.value)

# Assign file handle for later exchange rate data and open workbook

exchange_rate_excel2 = "Exchange rates f11hist.xls"
er_workbook2 = xlrd.open_workbook(exchange_rate_excel2)

# List sheet names, and pull a sheet by name

sheet_names = er_workbook2.sheet_names()

exchange_rate_sheet2 = er_workbook2.sheet_by_name(sheet_names[0])  # Sheet with exchange rate data

# Pull the required column by index

column_ER_aud_us2 = exchange_rate_sheet2.col(1)  # 'B' col - AUD/US exchange rate data

# Add values from Jan 2010 to June 2017s

for idx, cell_obj in enumerate(column_ER_aud_us2):
    if 11<= idx < 101:  # from Jan 2010 to June 2017
        ER_aud_us.append(cell_obj.value)

# Aggregate monthly data to annual

lower = 0

annual_ER_aud_us = []  # list of annual average exchange rage  97-98 to 16-17

for i in range(20):
    upper = lower + 12
    annual_sum = sum(ER_aud_us[lower:upper])
    mean_annual_ER = annual_sum/12
    lower = upper
    annual_ER_aud_us.append(mean_annual_ER)

############################
#  Step 2 - Visualise data #
############################

# Set up plots for economic data visualisation

fig, (ax1, ax3, ax4, ax2, ax5) = plt.subplots(5, 1, figsize=(17, 10),sharex=True)

fig.suptitle('Economic Analysis 1997-98 to 2016-17', fontsize=14, color='b')

ax1.set_title('Direct tourism GDP',fontsize=12)
ax2.set_title('Tourism share of total GDP', fontsize=12)
ax3.set_title('International visitors tourism GDP', fontsize=12)
ax4.set_title('Total GDP, current prices', fontsize=12)
ax5.set_title('Average annual AUD/US exchange rate', fontsize=12)

ax1.set_ylabel('$m', fontsize=10)
ax2.set_ylabel('%', fontsize=10)
ax3.set_ylabel('$m', fontsize=10)
ax4.set_ylabel('$m', fontsize=10)
ax5.set_ylabel('AUD/USD exchange rate', fontsize=10)

# Set up function for updating data each year

def update(num, x, y, line, ylim):
    line.set_data(x[:num], y[:num])
    line.axes.axis([0, 19, 0, ylim])
    return line,

# Set up data and run animation for Tourism GDP

line, = ax1.plot(year, Tourism_GDP, color='k')

ani = FuncAnimation(fig, update, len(year)+1, fargs=[year, Tourism_GDP, line, 60000],
                               interval=200, blit=True, repeat=False)  # Interval gives speed

# Set up data and run animation for Tourism share of GDP

line, = ax2.plot(year, Tourism_share_GDP, color='b')

ani2 = FuncAnimation(fig, update, len(year)+1, fargs=[year, Tourism_share_GDP, line, 5],
                               interval=200, blit=True, repeat=False)  # Interval gives speed

# Set up data and run animation for International visitors contribution to GDP

line, = ax3.plot(year, Int_TGDP, color='g')

ani3 = FuncAnimation(fig, update, len(year)+1, fargs=[year, Int_TGDP, line, 20000],
                               interval=200, blit=True, repeat=False)  # Interval gives speed

# Set up data and run animation for Total GDP

line, = ax4.plot(year, GDP_cp, color='r')

ani4 = FuncAnimation(fig, update, len(year)+1, fargs=[year, GDP_cp, line, 2000000],
                               interval=200, blit=True, repeat=False)  # Interval gives speed

# Set up data and run animation for exchange rate

line, = ax5.plot(year, annual_ER_aud_us, color='y')

ani5 = FuncAnimation(fig, update, len(year)+1, fargs=[year, annual_ER_aud_us, line, 1.2],
                               interval=200, blit=True, repeat=False)  # Interval gives speed

plt.xlabel('year')
plt.savefig("Economic Analysis.pdf")
plt.show()