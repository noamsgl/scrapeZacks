import configparser
import os
import sys
import tkinter as tk
from datetime import datetime
from io import StringIO
from tkinter import messagebox, Tk

import pandas as pd
import pyfiglet
import requests
from styleframe import StyleFrame, Styler

# Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
root = Tk()
canvas1 = tk.Canvas(root, width=300, height=300)

config = configparser.ConfigParser()
config.read('config.ini')
timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
processed_fnames = ['updated_dataset.xlsx', os.path.join('history', 'updated_dataset_{}.xlsx'.format(timestamp))]
scraped_fname = 'rank_1.xls'
urls = {"LOGIN": "https://www.zacks.com/logout.php",
        "BUY-LIST": "https://www.zacks.com/stocks/buy-list"}
xpaths = {"USER": r'//*[@id="login"]/form/table/tbody/tr[1]/td[2]/input',
          "PASSWORD": r'//*[@id="login"]/form/table/tbody/tr[2]/td[2]/input',
          "SUBMIT": r'//*[@id="login"]/form/table/tbody/tr[4]/td[2]/input',
          'RANK': r'//*[@id="stocks-menu"]/li[2]/a',
          'EXPORT': r'#export_1_rank_excel_data'}
header_cols = ['Symbol', 'Company Name', 'Industry', 'Price', 'Date Added']
numerical_cols = ['Dividend Yield(%)', 'Price Movers: 1 Day(%)', 'Price Movers: 1 Week(%)',
                  'Price Movers: 4 Week(%)', 'Biggest Est. Chg. Current Year(%)',
                  'Biggest Est. Chg. Next Year(%)', 'Biggest Surprise Last Qtr(%)', 'Market Cap (mil)',
                  'Projected Earnings Growth (1 Yr)(%)',
                  'Projected Earnings Growth (3-5 Yrs)(%)',
                  'Price / Sales']
date_cols = ['Date Added']
rank_ascend = ['Value Score', 'Growth Score', 'Momentum Score',
               'P/E (F1)', 'PEG', 'Price / Sales']
rank_descend = ['Dividend Yield(%)', 'Price Movers: 1 Day(%)', 'Price Movers: 1 Week(%)',
                'Price Movers: 4 Week(%)', 'Biggest Est. Chg. Current Year(%)',
                'Biggest Est. Chg. Next Year(%)', 'Biggest Surprise Last Qtr(%)',
                'Market Cap (mil)', 'Projected Earnings Growth (1 Yr)(%)',
                'Projected Earnings Growth (3-5 Yrs)(%)',
                'VGM Score'
                ]
drop_cols = ['Unnamed']
data_columns = ['Dividend Yield(%)', 'Price Movers: 1 Day(%)', 'Price Movers: 1 Week(%)',
                'Price Movers: 4 Week(%)', 'Biggest Est. Chg. Current Year(%)',
                'Biggest Est. Chg. Next Year(%)', 'Biggest Surprise Last Qtr(%)',
                'Market Cap (mil)', 'Projected Earnings Growth (1 Yr)(%)',
                'Projected Earnings Growth (3-5 Yrs)(%)',
                'Value Score', 'Growth Score', 'Momentum Score',
                'P/E (F1)', 'PEG', 'Price / Sales']
score_columns = ['score to {}'.format(col) for col in data_columns]
pair_columns = list(zip(iter(data_columns), iter(score_columns)))

def flatten(lst):
    return [item for sublist in lst for item in sublist]

odd_pair_columns = flatten(pair_columns[1::2])
even_pair_columns = flatten(pair_columns[0::2])
calculated_columns = ['Total score']



def execute(download=False):
    if download:
        try:
            data = scrape()
        except:
            messagebox.showinfo(title="Scraper and Scorer",
                                message='An error has occurred while trying to download data.\nExiting.')
            sys.exit(-1)
    else:
        data = scraped_fname

    try:
        print("Loading Dataframe")
        df = pd.read_csv(data, delimiter='\t')
        print("Processing Dataframe")
        # process dataframe
        df = process_dataframe(df)
    except:
        messagebox.showinfo(title="Scraper and Scorer",
                            message='An error has occurred while processing the data. Exiting.')
        sys.exit(-1)

    try:
        # save dataframe to styled excel workbook
        save(df)
    except:
        messagebox.showinfo(title="Scraper and Scorer",
                            message='An error has occurred while saving the Excel. Exiting.')
        sys.exit(-1)

    messagebox.showinfo(title="Scraper and Scorer",
                        message='Scoring Completed.')
    sys.exit(0)


def scrape():
    header = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:32.0) Gecko/20100101 Firefox/32.0', }
    urls = {"HOME": "https://www.zacks.com",
            "RANK1": "https://www.zacks.com/portfolios/rank/z2_1rank_tab.php?reference_id=all&sort_menu=show_date_added&tab_name=Full%20%231%20List&_=1610808269164",
            "RANK1EXCEL": "https://www.zacks.com/portfolios/rank/rank_excel.php?rank=1&reference_id=all",
            "STOCKSCREENER": "https://www.zacks.com/screening/stock-screener",
            "MYSCREEN": "https://screener-api.zacks.com/myscreen.php?screen_type=1&_=1610807751953",
            "ETFSCREENER": "https://www.zacks.com/screening/mutual-fund-screener"}

    print(20 * "*" + " Begin Scraping " + 20 * "*")
    with requests.Session() as s:
        s.headers.update(header)
        print("Logging in to {}".format(urls["HOME"]))
        payload = {'username': config["APP"]["USER"],
                   'password': config["APP"]["PASSWORD"]}
        p = s.post(urls["HOME"], data=payload)
        print("Downloading Rank1 Excel")
        r = s.get(urls["RANK1EXCEL"])
        bytes_data = r.content
        string_data = str(r.content, 'utf-8')
        data = StringIO(string_data)
    return data

def process_dataframe(df):
    # get rid of unnamed index
    df.drop(df.filter(regex="Unname"), axis=1, inplace=True)

    # parse dates
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], format='%b %d,%Y')

    # parse numbers
    for col in numerical_cols:
        df[col] = df[col].str.rstrip('%')
        df[col] = df[col].str.replace("[^0-9.\-]", "").astype(float)

    # 'Biggest Surprise Last Qtr(%)'
    # rank ascending
    for col in rank_ascend:
        df['score to {}'.format(col)] = df[col].rank(ascending=True, method='min') - 1

    # rank descending
    for col in rank_descend:
        df['score to {}'.format(col)] = df[col].rank(ascending=False, method='min') - 1

    # add total column
    df['Total score'] = df[score_columns].sum(axis=1)

    # reorder columns
    df = df[header_cols + list(sum(list(zip(data_columns, score_columns)), ())) + calculated_columns]

    return df

def save(df):
    print("Saving Excel")
    df.index = df.index + 1
    for processed_fname in processed_fnames:
        excel_writer = StyleFrame.ExcelWriter(processed_fname)
        sf = StyleFrame(df)
        # sf.apply_column_style(cols_to_style=header_cols, styler_obj=Styler(bg_color=utils.colors.white),
        #                       style_header=True)

        sf.apply_column_style(cols_to_style=odd_pair_columns, styler_obj=Styler(bg_color='#ffe28a', wrap_text=True),
                              style_header=True)
        sf.apply_column_style(cols_to_style=even_pair_columns, styler_obj=Styler(bg_color='#d6ffba', wrap_text=True),
                              style_header=True)

        sf.to_excel(
            excel_writer=excel_writer,
            best_fit=list(df.columns),
            columns_and_rows_to_freeze='C2',
            row_to_add_filters=0,
            index=True
        )
        try:
            excel_writer.save()
        except PermissionError as PE:
            print("Could not save file. Please make sure the file: \"{}\" is closed.".format(PE.filename))
            quit(-1)
    print("Excel Saved")


if __name__ == '__main__':
    print(pyfiglet.figlet_format("Hello!!!"))
    button1 = tk.Button(root, text='Download and Score', command=lambda : execute(download=True), bg='#42b6f5', fg='white')
    button1.pack(side=tk.TOP)
    button2 = tk.Button(root, text='Score Only', command=lambda: execute(download=False), bg='#42b6f5', fg='white')
    button2.pack(side=tk.BOTTOM)
    root.mainloop()
