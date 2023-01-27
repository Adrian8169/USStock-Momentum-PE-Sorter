import os
import json
import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from scipy import stats

#This is a momentum based pe sorting bot which focuses on companies with a high implied momentum (an average over a number of time periods)
#and a low p/e Ratio to display a list of 15 companys. This was made as a 'fun' project and should not be used in a professional setting without modification.

#Issues for consideration - this does not remove outliers from the pool, new/hype stocks with large growth rates over 1 year may be unfairly weighted.
# Pe ratios are useful only if companies satisfy certain criterion, such a filter is not used in this code, and should be added for higher accuracy.

stocks = pd.read_csv("C:/Users/YOURNAMEHERE/PycharmProjects/GitHub - Portfolio allocator by momentum and PE/sp_500_stocks.csv")
from secrets import IEX_CLOUD_API_TOKEN

symbol = 'AAPL'
api_url = f'https://cloud.iexapis.com/stable/stock/{symbol}/stats?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()

def chunks(lst, n):
    "Yield successive n-sized chunks from lst."
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

hqm_columns = [
                'Ticker',
                'Price',
                'Number of Shares to Buy',
                'One-Year Price Return',
                'One-Year Return Percentile',
                'Six-Month Price Return',
                'Six-Month Return Percentile',
                'Three-Month Price Return',
                'Three-Month Return Percentile',
                'One-Month Price Return',
                'One-Month Return Percentile',
                'HQM Score',
                'PE Ratio',
                'HQ PE'
                ]

print('Thank you for using this portfolio bot,'
      , os.linesep,
      ' this specific bot gives a list of stocks based on'
      , os.linesep,
      ' their long run momentum as well as their p/e ratio.'
      , os.linesep, os.linesep,
      '(FOR USE AS AN EXAMPLE, NOT REAL INVESTMENT ADVICE)'
      , os.linesep,
      )

def portfolio_input():
    global portfolio_size
    portfolio_size = input(" Enter the value of your portfolio:")

    try:
        val = float(portfolio_size)
    except ValueError:
        print(" That's not a number! \n Try again:")
        portfolio_size = input(" Enter the value of your portfolio:")

portfolio_input()
print('Calculating... Please wait up to 45 seconds.')
hqm_dataframe = pd.DataFrame(columns=hqm_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://cloud.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = pd.concat([hqm_dataframe, pd.DataFrame({
            hqm_dataframe.columns[0]: [symbol],
            hqm_dataframe.columns[1]: [data[symbol]['quote']['latestPrice']],
            hqm_dataframe.columns[2]: ['N/A'],
            hqm_dataframe.columns[3]: [data[symbol]['stats']['year1ChangePercent']],
            hqm_dataframe.columns[4]: ['N/A'],
            hqm_dataframe.columns[5]: [data[symbol]['stats']['month6ChangePercent']],
            hqm_dataframe.columns[6]: ['N/A'],
            hqm_dataframe.columns[7]: [data[symbol]['stats']['month3ChangePercent']],
            hqm_dataframe.columns[8]: ['N/A'],
            hqm_dataframe.columns[9]: [data[symbol]['stats']['month1ChangePercent']],
            hqm_dataframe.columns[10]: ['N/A'],
            hqm_dataframe.columns[11]: ['N/A'],
            hqm_dataframe.columns[12]: [data[symbol]['quote']['peRatio']],
            hqm_dataframe.columns[13]: ['N/A']
        }
        )
                                     ],
                                    ignore_index=True)

hqm_dataframe.columns
time_periods = [
                'One-Year',
                'Six-Month',
                'Three-Month',
                'One-Month'
                ]
for row in hqm_dataframe.index:
    for time_period in time_periods:

        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        if hqm_dataframe.loc[row, change_col] == None:
            hqm_dataframe.loc[row, change_col] = 0.0
for row in hqm_dataframe.index:
    for time_period in time_periods:
        hqm_dataframe.loc[row, f'{time_period} Return Percentile'] = stats.percentileofscore(hqm_dataframe[f'{time_period} Price Return'], hqm_dataframe.loc[row, f'{time_period} Price Return'])/100

from statistics import mean

for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'HQM Score'] = (mean(momentum_percentiles*100))

hqm_dataframe['HQ PE'] = hqm_dataframe['HQM Score']/hqm_dataframe['PE Ratio']

hqm_dataframe = hqm_dataframe.sort_values(by = 'HQM Score', ascending = False, ignore_index=True)
hqm_dataframe = hqm_dataframe.sort_values(by = 'HQ PE', ascending = False, ignore_index=True)

searchset = hqm_dataframe
print(searchset)

hqm_dataframe = hqm_dataframe[:16]

position_size = float(portfolio_size) / len(hqm_dataframe.index)
for i in range(0, len(hqm_dataframe['Ticker'])-1):
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / hqm_dataframe['Price'][i])
print(hqm_dataframe)
SSearch = input('To search for a specific company please enter the ticker here (Uppercase only please!)')
print(searchset.loc[searchset['Ticker']==SSearch])
SSearch = input('To search for a specific company please enter the ticker here (Uppercase only please!)')
print(searchset.loc[searchset['Ticker']==SSearch])

writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, sheet_name='Momentum Strategy', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_template = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_template = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_template = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

percent_template = writer.book.add_format(
        {
            'num_format':'0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

column_formats = {
                    'A': ['Ticker', string_template],
                    'B': ['Price', dollar_template],
                    'C': ['Number of Shares to Buy', integer_template],
                    'D': ['One-Year Price Return', percent_template],
                    'E': ['One-Year Return Percentile', percent_template],
                    'F': ['Six-Month Price Return', percent_template],
                    'G': ['Six-Month Return Percentile', percent_template],
                    'H': ['Three-Month Price Return', percent_template],
                    'I': ['Three-Month Return Percentile', percent_template],
                    'J': ['One-Month Price Return', percent_template],
                    'K': ['One-Month Return Percentile', percent_template],
                    'L': ['HQM Score', integer_template]
                    }

for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_template)

writer.save()