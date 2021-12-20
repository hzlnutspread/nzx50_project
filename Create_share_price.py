import time
import datetime
import pandas as pd

tickers = ['fph', 'aia', 'spk', 'mft', 'cen', 'ift', 'fbu', 'mel', 'rym', 'ebo', 'atm', 'mcy', 'cnu', 'sum', 'gmt',
           'skc', 'pot', 'fre', 'pct', 'kpg', 'zel', 'gne', 'pfi', 'arg', 'pph', 'hgh', 'vhp', 'skl', 'arv', 'vct',
           'kmd', 'peb', 'spg', 'oca', 'air', 'scl', 'sko', 'ipl', 'vgl', 'nzx', 'tpw', 'rbd', 'skt', 'fsf', 'san',
           'thl', 'sml', 'nph']
period1 = int(time.mktime(datetime.datetime(2016, 9, 30, 11, 59).timetuple()))
period2 = int(time.mktime(datetime.datetime.today().timetuple()))
interval = '1d'

for ticker in tickers:
    print(f'retrieving {ticker}')
    query_string = f'https://query1.finance.yahoo.com/v7/finance/download/{ticker}.NZ?period1={period1}&period2={period2}' \
                   f'&interval={interval}&events=history&includeAdjustedClose=true'

    df = pd.read_csv(query_string)

    arr = []
    for index, row in df.iterrows():
        date = pd.to_datetime(row['Date'])
        if date.weekday() == 4:
            subarr = [date.strftime('%d-%b-%Y'), row['Close']]
            arr.append(subarr)

    df2 = pd.DataFrame(arr, columns=['Date', 'Close'])

    df2.to_csv(f'./csvs/{ticker}-share-price.csv', index=False, header=False)
