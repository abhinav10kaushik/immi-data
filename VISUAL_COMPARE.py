import pandas as pd
import xlrd
import datetime
import csv

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.font_manager import FontProperties
from matplotlib.ticker import MaxNLocator

'''
This function is used to extract Data1 sheet from 340101.xls and 340102.xls files and convert it to csv for easier 
processing. It clears the rows and keep only the rows where the useful information is
File 340101 contains all Arrival movements related to migration and downloaded from 
http://abs.gov.au/AUSSTATS/ABS@Archive.nsf/log?openagent&340101.xls&3401.0&Time%20Series%20Spreadsheet&23F247356B1A250CCA25830800119CF6&0&Jul%202018&17.09.2018&Latest
File 340102 contains all Outward movements related to migration and downloaded from 
http://abs.gov.au/AUSSTATS/ABS@Archive.nsf/log?openagent&340102.xls&3401.0&Time%20Series%20Spreadsheet&0A2A23A6256D8AB0CA25830800119F0C&0&Jul%202018&17.09.2018&Latest
'''


def convert_to_csv(name):
    with xlrd.open_workbook('{}.xls'.format(name)) as wb:
        # Sheet with index 1 is Data1 sheet
        sh = wb.sheet_by_index(1)
        num_rows = sh.nrows
        num_cols = sh.ncols
        skip = True

        with open('{}.csv'.format(name), 'w', newline="") as f:
            c = csv.writer(f)
            for r in range(num_rows):
                # skip until find 'Series ID'
                if sh.cell(r, 0).value == "Series ID":
                    skip = False
                    continue
                if skip:
                    continue
                # date is in float format so it must be converted to string date format
                date = datetime.datetime(*xlrd.xldate_as_tuple(sh.cell(r, 0).value, wb.datemode))
                date = str(date)
                date = date.split(" ")[0]
                # write the new row to the file
                c.writerow([date] + [sh.cell(r, j).value for j in range(1, num_cols)])


'''
This function is used to clean the .xls file that contains the national average weekly earnings and converts it 
into .csv for easier processing
File 6302003 is downloaded from 
http://www.abs.gov.au/ausstats/meisubs.NSF/log?openagent&6302003.xls&6302.0&Time%20Series%20Spreadsheet&F35E044B5A542FDBCA2582EA00193CA8&0&May%202018&16.08.2018&Latest
'''


def convert_earnings_to_csv():
    with xlrd.open_workbook('6302003.xls') as wb:
        sh = wb.sheet_by_index(1)
        num_rows = sh.nrows
        num_cols = sh.ncols
        skip = True

        with open('6302003.csv', 'w', newline="") as f:
            c = csv.writer(f)
            for r in range(num_rows):
                # skip until find 'Series ID'
                if sh.cell(r, 0).value == "Series ID":
                    skip = False
                    continue
                if skip:
                    continue

                # date is in float format so it must be converted to string date format
                date = datetime.datetime(*xlrd.xldate_as_tuple(sh.cell(r, 0).value, wb.datemode))
                date = str(date)
                date = date.split(" ")[0]
                # write the new row to the file
                c.writerow([date, sh.cell(r, 9).value])


'''
This function is used to download the file mon_pax_web from 
https://data.gov.au/dataset/cc5d888f-5850-47f3-815d-08289b22f5a8/resource/38bdc971-cb22-4894-b19a-814afc4e8164/download/mon_pax_web.csv
'''


def load_pax_in_out():
    
    # load the csv file
    df = pd.read_csv("mon_pax_web.csv")
    # columns we want to keep
    columns = ["AIRPORT", "Year", "Month", "Int_Pax_In", "Int_Pax_Out"]
    df = df[columns]
    # keep only the entries for the year Jan 2010 - Dec 2017 and remove total numbers of passengers from data
    df = df.loc[(df["Year"] >= 2010) & (df["Year"] <= 2017)]
    df = df.loc[(df["AIRPORT"] != "All Australian Airports")]
    # return the data frame
    return df


def load_filtered_pax_in_out():
    # load only the desired columns of .csv (using task1)
    df = load_pax_in_out()
    # get all airport names
    all_airports = set(df["AIRPORT"].unique())

    # the airport names we want to keep as they are
    listed_airports = set(["SYDNEY", "ADELAIDE", "PERTH", "BRISBANE", "HOBART", "DARWIN", "GOLD COAST", "MELBOURNE"])
    # create a mapping for airport names
    mapping = {}
    for airport in all_airports:
        # if the aiport name exists in the list then keep the same name
        if airport in listed_airports:
            mapping[airport] = airport
        # otherwise change the airport name to 'OTHERS'
        else:
            mapping[airport] = "OTHERS"
    # apply the mapping rule into the airport names
    df.AIRPORT = [mapping[item] for item in df.AIRPORT]
    return df

'''
This function is used to download the file f11_data.csv from 
https://www.rba.gov.au/statistics/tables/csv/f11-data.csv
'''

def load_us_exchange_rate():
# clear the first rows of the file
    f = open("f11-data.csv", "r")
    skip = True
    month_dict = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10,
                  "Nov": 11, "Dec": 12}
    columns = ["Month", "Year", "USXR"]

    # extract the month and year
    def extract_date(date_string):
        splitted = date_string.split("-")
        return month_dict[splitted[1]], int(splitted[2]),

    data = []
    for line in f:
        line = line.replace("\n", "")
        # skip until 'Series ID' is found
        if "Series ID" in line:
            skip = False
            continue
        if skip:
            continue
        splitted = line.split(",")
        # extract month and year from the string
        month, year = extract_date(splitted[0])
        data.append([month, year, eval(splitted[1])])
    f.close()
    # create a dataframe from the lists
    df = pd.DataFrame.from_records(data, columns=columns)
    df = df.loc[(df["Year"] >= 2010) & (df["Year"] <= 2017)]
    return df


def load_movements_csv(name):
    # open file
    f = open("{}.csv".format(name), "r")
    # define the columns of the dataframe
    columns = ["Month", "Year", "Citizens", "Permanent Residents", "LT Visitors(Long Term)", "ST Resident(Short Term)",
               "ST Visitors(Tourist)"]
    data = []

    def extract_date(data_string):
        splitted = data_string.split("-")
        return int(splitted[1]), int(splitted[0])

    for line in f:
        line = line.replace("\n", "")
        splitted = line.split(",")
        splitted = splitted[:len(columns) - 1]
        month, year = extract_date(splitted[0])
        data.append([month, year] + [int(eval(splitted[i])) for i in range(1, len(splitted))])
    f.close()

    # create a dataframe from the lists
    df = pd.DataFrame.from_records(data, columns=columns)
    # keep only the entries between year 2010 and 2017
    df = df.loc[(df["Year"] >= 2010) & (df["Year"] <= 2017)]
    # create two new columns 'Residents' and 'Visitors' which
    # are the sum of the corresponding columns
    df["Residents"] = df["Citizens"] + df["Permanent Residents"] + df["ST Resident(Short Term)"]
    df["Visitors"] = df["LT Visitors(Long Term)"] + df["ST Visitors(Tourist)"]
    return df


def load_arrival_movements():
    return load_movements_csv(340101)


def load_departure_movements():
    return load_movements_csv(340102)


def load_average_weekly_earnings():
    # open file
    f = open("6302003.csv", "r")
    # define the columns of the dataframe
    columns = ["Month", "Year", "Total Earnings"]
    data = []

    # this function is used to exctract month and year from the data string
    def extract_date(data_string):
        splitted = data_string.split("-")
        return int(splitted[1]), int(splitted[0])

    # parse all the lines of the file
    for line in f:
        line = line.replace("\n", "")
        splitted = line.split(",")
        splitted = splitted[:len(columns) - 1]
        # extract month and year from the data string
        month, year = extract_date(splitted[0])
        # add new row to the data
        data.append([month, year] + [int(eval(splitted[i])) for i in range(1, len(splitted))])
    f.close()

    # create dataframe from the lists
    df = pd.DataFrame.from_records(data, columns=columns)
    df = df.loc[(df["Year"] >= 2010) & (df["Year"] <= 2017)]
    return df


# VISUALISATION STARTS HERE

# will be used to set properties of legends
fontP = FontProperties()
fontP.set_size('small')
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# generates a random color that will be used for plots
def random_color():
    r = np.random.uniform(0, 0.7)
    g = np.random.uniform(0, 0.7)
    b = np.random.uniform(0, 0.7)
    return (r, g, b)


# mode: monthly, quarterly, half yearly, yearly
def get_data(df, column, mode):
    if mode == "yearly":
        n = 2
    elif mode == "half yearly":
        n = 3
    elif mode == "quarterly":
        n = 4
    elif mode == "monthly":
        n = 12
    else:
        n = 2

    years = df["Year"].unique()
    var = []
    labels = []

    for year in years:
        # create months intervals
        # for example for half yearly the intervals are [[1,6], [7,12]]
        intervals = np.linspace(1, 12, n).astype("int")
        for i in range(n - 1):
            # get the proper data
            value = df[(df["Year"] == year) & ((df["Month"] >= intervals[i]) & (df["Month"] <= intervals[i + 1]))][
                column].mean()
            var.append(value)
            # add (month + year) label
            labels.append(months[intervals[i + 1] - 1] + " " + str(year))
            intervals[i + 1] += 1
    return labels, var


def pax_in_out_vs_exchange_rate_plot(mode="yearly"):
    # get proper data using the mode
    years = df1["Year"].unique()
    xaxis, int_pax_in = get_data(df1, "Int_Pax_In", mode)
    xaxis, int_pax_out = get_data(df1, "Int_Pax_Out", mode)
    xaxis, exchange_rate = get_data(df3, "USXR", mode)

    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()
    # this line prevents printing numbers with exponential representation
    ax1.ticklabel_format(useOffset=False, style='plain')
    ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
    plt.gcf().subplots_adjust(bottom=0.15)
    plt.title("International Passengers (Pax) vs Exchange Rate (US$) ({})".format(mode))
    # plot the lines
    line1, = ax1.plot(int_pax_in, 'b-', label="Arrivals")
    line2, = ax1.plot(int_pax_out, 'r-', label="Departures")
    line3, = ax2.plot(exchange_rate, 'g-', label="Exchange Rate (US$)")
    ax1.set_xlabel("Year")
    ax1.set_xticklabels(years)

    ax2.set_ylabel("1.0 AUD in USD", color="g")
    ax1.set_ylabel("International Passengers (x1000)", color="b")
    box = ax1.get_position()
    # plot the labels of each line
    ax1.legend(handles=[line1, line2, line3], loc='upper center', bbox_to_anchor=(0.5, -0.11), fancybox=True,
               shadow=True, ncol=3)
    plt.savefig("visualizationA.pdf")
    print('saving graph as "visualizationA.pdf"')
    


def city_wise_pax_in_out_plot():
    airports = df2["AIRPORT"].unique()
    years = df2["Year"].unique()
    pax = {}
    for airport in airports:
        pax[airport] = {}
        pax[airport]["Int Pax In"] = []
        pax[airport]["Int Pax Out"] = []
        for year in years:
            int_pax_in = df2[(df2["AIRPORT"] == airport) & (df2["Year"] == year)]["Int_Pax_In"].sum()
            int_pax_out = df2[(df2["AIRPORT"] == airport) & (df2["Year"] == year)]["Int_Pax_Out"].sum()
            pax[airport]["Int Pax In"].append(int_pax_in)
            pax[airport]["Int Pax Out"].append(int_pax_out)

    ind = np.arange(len(years))
    width = 0.15
    rects = {}

    random_colors = [random_color() for i in range(len(pax))]
    for p in ["Int Pax In", "Int Pax Out"]:
        rects[p] = []
        offset = 0
        # create 2 subplots. On the top there will be the Int_Pax_In plot
        # On the bottom there will be the Int_Pax_Out plot
        if p == "Int Pax In":
            plt.subplot(211)
        else:
            plt.subplot(212)

        ax = plt.gca()

        # plot the bar for each airport
        for i in range(len(airports)):
            airport = airports[i]
            rect = ax.bar(ind + offset, pax[airport][p], width, color=random_colors[i])
            rects[p].append(rect)
            offset += width
        ax.ticklabel_format(useOffset=False, style='plain')
        plt.title("{} (per city)".format(p))
        if p != "Int Pax In":
            plt.xlabel("Years")
        plt.ylabel("International Passengers", color="b")
        ax.set_xticklabels(years)
        ax.set_xlim(0, len(years) + 1)
        ax.legend(rects[p], airports, loc="upper right", prop=fontP)
    plt.savefig("visualizationB.pdf")
    print('saving graph as "visualizationB.pdf"')




def pax_in_vs_residents_and_tourists_plot(mode="yearly"):
    years = df1["Year"].unique()
# get proper data using the mode
    xaxis, int_pax_in = get_data(df1, "Int_Pax_In", mode)
    xaxis, residents = get_data(df4, "Residents", mode)
    xaxis, visitors = get_data(df4, "Visitors", mode)

    ind = np.arange(len(xaxis))
    width = 0.15
    rects = {}

    fig, ax1 = plt.subplots()
    plt.title("Arrivals - All Passengers vs Residents and Tourists ({})".format(mode))

    ax2 = ax1.twinx()
    ax1.ticklabel_format(useOffset=False, style='plain')
    ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
    ax1.set_xticklabels(xaxis)
    ax1.set_ylabel("Passengers", color="g")
    ax2.set_ylabel("Residents and Tourists", color="b")

    rect0 = ax1.bar(ind, int_pax_in, width, color="g", label="Arrivals")
    line1, = ax2.plot(residents, color="b", label="Residents")
    line2, = ax2.plot(visitors, color="r", label="Visitors")

    box = ax1.get_position()
    ax1.legend(handles=[rect0, line1, line2], loc='upper center', bbox_to_anchor=(0.5, -0.07), fancybox=True,
               shadow=True, ncol=3)
    plt.savefig("visualizationC.pdf")
    print('saving graph as "visualizationC.pdf"')




def pax_out_vs_residents_and_tourists_plot(mode="yearly"):
    # get proper data using the mode
    xaxis, int_pax_out = get_data(df1, "Int_Pax_Out", mode)
    xaxis, residents = get_data(df4, "Residents", mode)
    xaxis, visitors = get_data(df4, "Visitors", mode)

    ind = np.arange(len(xaxis))
    width = 0.15
    rects = {}

    fig, ax1 = plt.subplots()
    plt.title("Departures - All Passengers vs Residents and Tourists ({})".format(mode))

    ax2 = ax1.twinx()
    ax1.ticklabel_format(useOffset=False, style='plain')
    ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
    ax1.set_xticklabels(xaxis)
    ax1.set_ylabel("Passengers", color="r")
    ax2.set_ylabel("Residents and Tourists", color="b")

    rect0 = ax1.bar(ind, int_pax_out, width, color="r", label="Departures")
    line1, = ax2.plot(residents, color="g", label="Residents")
    line2, = ax2.plot(visitors, color="b", label="Visitors")

    box = ax1.get_position()
    ax1.legend(handles=[rect0, line1, line2], loc='upper center', bbox_to_anchor=(0.5, -0.07), fancybox=True,
               shadow=True, ncol=3)
    plt.savefig("visualizationD.pdf")
    print('saving graph as "visualizationD.pdf"')

def pax_in_out_vs_weekly_earnings_plot():
    int_pax_in = []
    int_pax_out = []
    exchange_rate = []
    years = df1["Year"].unique()

    xaxis = []
    for year in years:
        xaxis.append(months[5] + " " + str(year))
        xaxis.append(months[11] + " " + str(year))

    for year in years:
        pax_in_first_half_year = df1[(df1["Year"] == year) & ((df1["Month"] >= 1) & (df1["Month"] <= 6))][
            "Int_Pax_In"].mean()
        int_pax_in.append(pax_in_first_half_year)
        pax_in_second_half_year = df1[(df1["Year"] == year) & ((df1["Month"] >= 7) & (df1["Month"] <= 12))][
            "Int_Pax_In"].mean()
        int_pax_in.append(pax_in_second_half_year)
        pax_out_first_half_year = df1[(df1["Year"] == year) & ((df1["Month"] >= 1) & (df1["Month"] <= 6))][
            "Int_Pax_Out"].mean()
        int_pax_out.append(pax_out_first_half_year)
        pax_out_second_half_year = df1[(df1["Year"] == year) & ((df1["Month"] >= 7) & (df1["Month"] <= 12))][
            "Int_Pax_Out"].mean()
        int_pax_out.append(pax_out_second_half_year)

    total_earnings = list(df6["Total Earnings"])

    ind = np.arange(len(int_pax_in))
    width = 0.15

    fig, ax1 = plt.subplots()
    plt.title("International Passengers vs Average Weekly Earnings")

    ax2 = ax1.twinx()
    ax1.ticklabel_format(useOffset=False, style='plain')
    ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
    ax1.set_xticklabels(xaxis)
    ax1.set_ylabel("Passengers", color="g")
    ax2.set_ylabel("Avg Weekly Earnings", color="b")

    rect0 = ax1.bar(ind, int_pax_in, width, color="g", label="Arrivals")
    rect1 = ax1.bar(ind + width, int_pax_out, width, color="r", label="Departures")
    line, = ax2.plot(total_earnings, color="b", label="Avg Weekly Earnings")

    box = ax1.get_position()
    ax1.legend(handles=[rect0, rect1, line], loc='upper center', bbox_to_anchor=(0.5, -0.07), fancybox=True,
               shadow=True, ncol=3)
    plt.savefig("visualizationE.pdf")
    print('saving graph as "visualizationE.pdf"')
    plt.show()


df1 = load_pax_in_out()
df2 = load_filtered_pax_in_out()
df3 = load_us_exchange_rate()
df4 = load_arrival_movements()
df5 = load_departure_movements()
df6 = load_average_weekly_earnings()

# mode: monthly, quarterly, half yearly, yearly
mode = "yearly"
pax_in_out_vs_exchange_rate_plot(mode=mode)
city_wise_pax_in_out_plot()
pax_in_vs_residents_and_tourists_plot(mode=mode)
pax_out_vs_residents_and_tourists_plot(mode=mode)
pax_in_out_vs_weekly_earnings_plot()
