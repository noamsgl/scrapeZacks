#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Project: Zack's #1 Rank
Software Developer: Noam Siegel <noamsi@post.bgu.ac.il>
"""

import configparser
import os
import sys
import tkinter as tk
from datetime import datetime
from io import StringIO
from tkinter import messagebox, Tk, filedialog

import pandas as pd
import pyfiglet
import requests
from styleframe import StyleFrame, Styler, utils
from win32com.client import Dispatch

# Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

root = Tk()
canvas1 = tk.Canvas(root, width=300, height=300)

config = configparser.ConfigParser()
config.read('config.ini')
timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
processed_fnames = ['rank1.xlsx',
                    os.path.join('history', 'rank1', 'rank1_{}.xlsx'.format(timestamp))]
scraped_fname = 'rank_1.xls'
header_cols = ['Index', 'Symbol', 'Company Name', 'Industry', 'Price', 'Date Added']
numerical_cols = ['P/E (F1)', 'Dividend Yield(%)', 'Price Movers: 1 Day(%)', 'Price Movers: 1 Week(%)',
                  'Price Movers: 4 Week(%)', 'Biggest Est. Chg. Current Year(%)',
                  'Biggest Est. Chg. Next Year(%)', 'Biggest Surprise Last Qtr(%)', 'Market Cap (mil)',
                  'Projected Earnings Growth (1 Yr)(%)',
                  'Projected Earnings Growth (3-5 Yrs)(%)',
                  'Price / Sales']
rank_ascend = ['Value Score', 'Growth Score', 'Momentum Score',
               'P/E (F1)', 'PEG', 'Price / Sales']
rank_descend = ['Dividend Yield(%)', 'Price Movers: 1 Day(%)', 'Price Movers: 1 Week(%)',
                'Price Movers: 4 Week(%)', 'Biggest Est. Chg. Current Year(%)',
                'Biggest Est. Chg. Next Year(%)', 'Biggest Surprise Last Qtr(%)',
                'Market Cap (mil)', 'Projected Earnings Growth (1 Yr)(%)',
                'Projected Earnings Growth (3-5 Yrs)(%)',
                'VGM Score']
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
        data = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("excel files", "*.xls"), ("all files", "*.*")))

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
    # add index column
    df['Index'] = range(1, len(df) + 1)

    # get rid of unnamed index
    df.drop(df.filter(regex="Unname"), axis=1, inplace=True)

    # parse dates
    # for col in date_cols:
    #     df[col] = pd.to_datetime(df[col], format='%b %d,%Y')

    # parse numbers
    for col in numerical_cols:
        df[col] = df[col].str.rstrip('%')
        df[col] = df[col].str.replace("[^0-9.\-]", "").astype(float)

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
        font = utils.fonts.calibri
        sf = StyleFrame(df)
        # sf.apply_column_style(cols_to_style=header_cols, styler_obj=Styler(bg_color=utils.colors.white),
        #                       style_header=True)
        sf.apply_column_style(cols_to_style=(header_cols + calculated_columns),
                              styler_obj=Styler(bg_color=utils.colors.white, wrap_text=False, font=font,
                                                font_size=12),
                              style_header=True)
        sf.apply_column_style(cols_to_style=odd_pair_columns,
                              styler_obj=Styler(bg_color='#ffe28a', wrap_text=False, font=font,
                                                font_size=12),
                              style_header=True)
        sf.apply_column_style(cols_to_style=even_pair_columns,
                              styler_obj=Styler(bg_color='#d6ffba', wrap_text=False, font=font,
                                                font_size=12),
                              style_header=True)
        sf.to_excel(
            excel_writer=excel_writer,
            best_fit=list(df.columns),
            columns_and_rows_to_freeze='C2',
            row_to_add_filters=0,
            index=False  # Index Column Added Seperately
        )
        try:
            excel_writer.save()
            print("Excel \"{}\" Saved.".format(processed_fname))
        except PermissionError as PE:
            print("Could not save file. Please make sure the file: \"{}\" is closed.".format(PE.filename))
            quit(-1)

        print("AutoFit Column Width for {}".format(processed_fname))
        excel = Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(os.getcwd(), processed_fname))

        # Activate second sheet
        # excel.Worksheets(2).Activate()

        # Autofit column in active sheet
        excel.ActiveSheet.Columns.AutoFit()

        # Save changes in a new file
        # wb.SaveAs("D:\\output_fit.xlsx")

        # Or simply save changes in a current file
        wb.Save()

        wb.Close()


if __name__ == '__main__':
    print(pyfiglet.figlet_format("Hello!!!"))
    button1 = tk.Button(root, text='Download and Score', command=lambda: execute(download=True), bg='#42b6f5',
                        fg='white')
    button1.pack(side=tk.TOP)
    button2 = tk.Button(root, text='Score Only', command=lambda: execute(download=False), bg='#42b6f5', fg='white')
    button2.pack(side=tk.BOTTOM)
    root.mainloop()
