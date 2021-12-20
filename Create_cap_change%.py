import pandas as pd
import numpy as np

tickers = ['fph', 'aia', 'spk', 'mft', 'cen', 'ift', 'fbu', 'mel', 'rym', 'ebo', 'atm', 'mcy', 'cnu', 'sum', 'gmt',
           'skc', 'pot', 'fre', 'pct', 'kpg', 'zel', 'gne', 'pfi', 'arg', 'pph', 'hgh', 'vhp', 'skl', 'arv', 'vct',
           'kmd', 'peb', 'spg', 'oca', 'air', 'scl', 'sko', 'ipl', 'vgl', 'nzx', 'tpw', 'rbd', 'skt', 'fsf', 'san',
           'thl', 'sml', 'nph']

for ticker in tickers:
    print(f'opening {ticker}-cap')

    cap_csv = f'./csvs/{ticker}-cap.csv'
    df = pd.read_csv(cap_csv)

    df.columns = ['Date', 'Cap']

    cap_change = pd.Series(df['Cap'].pct_change() * 100).round(decimals=2).array #do this with a loop you pussy
    date = pd.Series(df['Date']).array

    df2 = pd.DataFrame(date, columns=['Date'])
    df3 = pd.DataFrame(cap_change, columns=['Cap'])
    df4 = pd.concat([df2, df3], axis=1)
    df5 = df4.drop(index=0)

    df5.to_csv(f'./csvs/{ticker}-cap-change%.csv', index=False, header=False)


