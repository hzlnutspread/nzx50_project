import openpyxl
import pandas as pd

tickers = ['fph', 'aia', 'spk', 'mft', 'cen', 'ift', 'fbu', 'mel', 'rym', 'ebo', 'atm', 'mcy', 'cnu', 'sum', 'gmt',
           'skc', 'pot', 'fre', 'pct', 'kpg', 'zel', 'gne', 'pfi', 'arg', 'pph', 'hgh', 'vhp', 'skl', 'arv', 'vct',
           'kmd', 'peb', 'spg', 'oca', 'air', 'scl', 'sko', 'ipl', 'vgl', 'nzx', 'tpw', 'rbd', 'skt', 'fsf', 'san',
           'thl', 'sml', 'nph']

wb = openpyxl.load_workbook('C:/Users/user/NZX 50.xlsx', read_only=True)

for ticker in tickers:

    arr = []
    for sheet in wb.sheetnames:
        sh = wb[f'{sheet}']

        checker = False
        for row in sh.iter_rows():
            if checker is True:
                break
            for cell in row:

                if cell.value == f'{ticker}'.upper():
                    checker = True
                    date_val = sh.cell(row=cell.row, column=12).value
                    if date_val is None:
                        raise Exception('You fucked up ' + sheet)
                    cell_val = sh.cell(row=cell.row, column=4).value
                    subarr = [date_val.strftime('%d-%b-%Y'), cell_val]
                    arr.append(subarr)
                    break

    df = pd.DataFrame(arr, columns=['Date', 'Rank'])
    df.to_csv(f'./csvs/{ticker}-rank.csv', index=False, header=False)
