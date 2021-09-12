import pandas as pd
from pandas.tseries.offsets import CustomBusinessDay
from datetime import datetime

startTime = datetime.now()

# I defined an Excel writer object and the target file
excel_writer = pd.ExcelWriter("nse_multiple2021aug.xlsx", engine="xlsxwriter")

# these are editable dates to start and end web scrap (yyyymmdd)
first_date = "2021-06-25"  # 2006-09-11 this is the first date in history
last_date = "2021-09-06"

# empty list
pages = []  # to store url
df_list = []  # to store dataframes
days = []  # to store days

# LOOP THROUGH ALL DAYS TO SCRAP and exclude holidays
my_custom_calendar = CustomBusinessDay(
    holidays=[
        # '2011-01-01', '2011-04-22', '2011-04-25', '2011-05-01', '2011-06-01', '2011-10-20',
        # '2011-12-12', '2011-12-25', '2011-12-26',

        # '2012-01-01', '2012-01-02','2012-04-06', '2012-04-09', '2012-05-01', '2012-06-01',"2012-08-20",
        # '2012-10-20','2012-12-12', '2012-12-25', '2012-12-26',
        #
        # '2013-01-01', '2013-03-29', '2013-04-01', '2013-05-01', '2013-06-01','2013-08-09','2013-10-20',
        # '2013-10-21','2013-12-12', '2013-12-25', '2013-12-26',
        #
        # '2014-01-01', '2014-04-18', '2014-04-21', '2014-05-01', '2014-06-01', '2014-07-28','2014-10-05',
        # '2014-10-20','2014-12-12', '2014-12-25', '2014-12-26',
        #
        # '2015-01-01', '2015-04-03', '2015-04-06', '2015-05-01', '2015-06-01','2015-06-18', '2015-10-20',
        # '2015-11-26','2015-12-12', '2015-12-25', '2015-12-26',
        #
        # '2016-01-01', '2016-03-25', '2016-03-28', '2016-05-01', '2016-06-01', "2016-07-07", "2016-09-11",
        # '2016-10-20', '2016-12-12', '2016-12-25', '2016-12-26', '2016-12-27',
        #
        # '2017-01-01', '2017-01-02','2017-04-14', '2017-04-17', '2017-05-01', '2017-06-01', '2017-06-26',
        # '2017-08-08',"2017-09-01",'2017-10-20','2017-10-26','2017-12-12', '2017-12-25', '2017-12-26',
        #
        # '2018-01-01', '2018-03-30', '2018-04-02', '2018-05-01', '2018-06-01', '2018-06-15', '2018-08-21',
        # '2018-10-20', '2018-10-20', '2018-12-12', '2018-12-25', '2018-12-26',
        #
        # '2019-01-01', '2019-04-19', '2019-04-22', '2019-05-01', '2019-06-01', '2019-06-05', '2019-08-12',
        # '2019-10-20', '2019-10-21', '2019-10-27', '2019-12-12', '2019-12-25', '2019-12-26',
        #
        # '2020-01-01', '2020-04-22', '2020-04-25', '2020-05-01', '2020-06-01', '2020-10-20',
        # '2020-12-12', '2020-12-25', '2020-12-26'

        '2021-01-01'
    ])
bizday = pd.date_range(start=first_date, end=last_date, freq=my_custom_calendar)

for day in bizday:
    day = day.strftime("%Y%m%d")
    days.append(day)
    url = 'https://live.mystocks.co.ke/price_list/{}'.format(day)
    pages.append(url)

# USE URL, SCRAP DATA AND APPEND DATAFRAMES TO A LIST
for day, url in zip(days, pages):
    dfs = pd.read_html(url, attrs={'class': 'tblHoverHi'}, header=1)
    df = dfs[0]
    df = pd.DataFrame(df,
                      columns=['Date', 'CODE', 'Previous', 'High.1', 'Low.1', 'Price', 'Volume'])  # create date column
    datetimeobject = datetime.strptime(day, '%Y%m%d')
    day = datetimeobject.strftime('%m-%d-%Y')
    df.Date = day  # assign day to date column
    df_list.append(df)  # append df to df_list

# concat all data frames
df = pd.concat(df_list)
# drop all empty rows
df = df.dropna()
# drop all rows with length >6 in the 'Previous' column
df = df.drop(df[df['Previous'].map(len) > 7].index)
# df['Previous'][:10]
# drop rows with indices and banking
df = df[~df['Previous'].isin(['Indices', 'Banking'])]
# replace ^ with nothing
df['CODE'] = df['CODE'].str.replace(u"^", "")
# loop through columns from previous and replace '-' with 0
for column in df.columns[2:]:
    df[f"{column}"] = df[f"{column}"].replace(['-'], '0')

# loop through columns from previous and convert data type to float
for column in df.columns[2:]:
    df[f"{column}"] = pd.to_numeric(df[f"{column}"])

# Rename all columns
df.columns = ['Date', 'CODE', 'Previous', 'High', 'Low', 'Close', 'Volume']

# concatenate df_list to form 1 DataFrame
df.to_excel(excel_writer, index=False)

# And finally save the file
excel_writer.save()
print(datetime.now() - startTime)
