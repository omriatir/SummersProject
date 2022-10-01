import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
# Importing the list of stocks
stocks = pd.read_csv('sp_500_stocks.csv')
stocks = stocks[~stocks['Ticker'].isin(['DISCA', 'HFC','VIAC','WLTW'])]

# Acquiring and using an API Token
from secrets import IEX_CLOUD_API_TOKEN

symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
price = data['latestPrice']
market_cap = data['marketCap']
my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of shares to buy']
final_dataframe = pd.DataFrame(columns=my_columns)
final_dataframe = final_dataframe.append(pd.Series([symbol, price, market_cap, 'N/A'], index=my_columns), ignore_index=True)
# Looping trough the list of stocks (inefficient)
# final_dataframe = pd.DataFrame(columns=my_columns)
# for symbol in stocks['Ticker'][:1]:
#    api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
#    data = requests.get(api_url).json()
#    final_dataframe = final_dataframe.append(pd.Series([symbol, data['latestPrice'], data['marketCap'], 'N/A'], index=my_columns), ignore_index=True)


# Improving performance
def chunks(list, n):
    for i in range(0, len(list), n):
        yield list[i:i + n]


symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))
final_dataframe = pd.DataFrame(columns=my_columns)
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series([symbol, data[symbol]['quote']['latestPrice'], data[symbol]['quote']['marketCap'], 'N/A'], index=my_columns), ignore_index=True)
# Calculating the number of shares to buy
portfolio_size = input('Enter the value of your portfolio: ')
try:
    val = float(portfolio_size)
except ValueError:
    print('That is not an integer!! \n Please try again: ')
    portfolio_size = input('Enter the value of your portfolio: ')
    val = float(portfolio_size)
position_size = val/len(final_dataframe.index)
for i in range(0, len(final_dataframe['Ticker'])):
    final_dataframe.loc[i, 'Number of shares to buy'] = math.floor(position_size/final_dataframe['Stock Price'][i])
print(final_dataframe)

# Formatting Excel Output
writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index=False)
background_color = '#0a0a23'
font_color = '#ffffff'
string_format = writer.book.add_format({'font_color': font_color, 'bg_color': background_color, 'border': 1})
dollar_format = writer.book.add_format({'num_format': '$0.00', 'font_color': font_color, 'bg_color': background_color, 'border': 1})
integer_format = writer.book.add_format({'num_format': '0', 'font_color': font_color, 'bg_color': background_color, 'border': 1})
column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}
for column in column_formats:
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()
