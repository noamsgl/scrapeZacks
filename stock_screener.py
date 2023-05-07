#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Project: My Stock Screener
Software Developer: Noam Siegel <noamsi@post.bgu.ac.il>
"""
import traceback
import configparser
import os
import sys
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from io import StringIO
from tkinter import Tk, filedialog, messagebox

import pandas as pd
import pyfiglet
import requests
from styleframe import StyleFrame, Styler, utils

# TODO: FUTURE: replace this with xlwings
# from win32com.client import Dispatch

# Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
root = Tk()
canvas1 = tk.Canvas(root, width=300, height=300)

config = configparser.ConfigParser()
config.read('config.ini')
timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
history_base_dir = os.path.join('history', 'stock_screener')
if not os.path.exists(history_base_dir):
    os.makedirs(history_base_dir, exist_ok=True)
processed_fnames = ['stock_screener.xlsx',
                    os.path.join('history', 'stock_screener', 'stock_screener_{}.xlsx'.format(timestamp))]
# TODO: add 'Last Close'
header_cols = ['Index', 'Ticker', 'Company Name', 'Last Close']
numerical_cols = []
date_cols = []
rank_ascend = ['Zacks Rank', 'Momentum Score', "VGM Score",
               'Growth Score', 'Value Score', 'Zacks Industry Rank',
               'Current Avg Broker Rec']
rank_descend = ['Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                '% Price Change (YTD)']
drop_cols = []

# Defines the column orders
data_columns = ['Zacks Rank', 'Momentum Score', "VGM Score",
                'Growth Score', 'Value Score', 'Zacks Industry Rank',
                'Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                '% Price Change (YTD)', 'Current Avg Broker Rec']
results_data_columns = ['Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                        '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                        '% Price Change (YTD)']

score_columns = ['score to {}'.format(col) for col in data_columns]
calculated_columns = ['Total score', 'USA rankings', 'score to Results', 'Results rankings',
                      'Difference', 'Potential ranking']


pair_columns = list(zip(iter(data_columns), iter(score_columns)))


def flatten(lst):
    return [item for sublist in lst for item in sublist]


odd_pair_columns = flatten(pair_columns[1::2])
even_pair_columns = flatten(pair_columns[0::2])


def execute(download=False):
    if download:
        try:
            data_io = scrape()
            if data_io.getvalue() == 'Unauthorized access.':
                raise RuntimeError("Error: Unauthorized to access data.")
        except:
            messagebox.showinfo(title="Scraper and Scorer",
                                message='An error has occurred while trying to download data.\nExiting.')
            traceback.print_exc()
            sys.exit(-1)
        try:
            print("Reading data.")
            df = pd.read_csv(data_io, sep='\t')
        except Exception as e:
            messagebox.showinfo(title="Scraper and Scorer",
                                message='An error has occurred while trying to read data.\nExiting.')
            traceback.print_exc()
            sys.exit(-1)

    else:  # select and load csv
        print("Choose a *.csv file.")
        file_path: str = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                    filetypes=(("csv files", "*.csv"), ("all files", "*.*")))

        print("Loading Dataframe")
        df = pd.read_csv(file_path)

    try:
        print("Processing Dataframe")
        df = process_dataframe(df)
    except:
        messagebox.showinfo(title="Scraper and Scorer",
                            message='An error has occurred while processing the data. Exiting.')
        traceback.print_exc()
        sys.exit(-1)

    try:
        # save dataframe to styled excel workbook
        save(df)
    except Exception as e:
        messagebox.showinfo(title="Scraper and Scorer",
                            message='An error has occurred while saving the Excel. Exiting.')
        sys.exit(-1)

    messagebox.showinfo(title="Scraper and Scorer",
                        message='Scoring Completed.')
    sys.exit(0)


def scrape() -> StringIO:
    header = {
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:32.0) Gecko/20100101 Firefox/32.0', }
    urls = {"HOME": "https://www.zacks.com",
            "RANK1": "https://www.zacks.com/portfolios/rank/z2_1rank_tab.php?reference_id=all&sort_menu=show_date_added&tab_name=Full%20%231%20List&_=1610808269164",
            "RANK1EXCEL": "https://www.zacks.com/portfolios/rank/rank_excel.php?rank=1&reference_id=all",
            "STOCKSCREENER": "https://www.zacks.com/screening/stock-screener",
            "MYSCREEN": "https://screener-api.zacks.com/myscreen.php?screen_type=1&_=1610807751953",
            "ETFSCREENER": "https://www.zacks.com/screening/mutual-fund-screener",
            "EXPORT.PHP": "https://screener-api.zacks.com/export.php"}

    print(20 * "*" + " Begin Scraping " + 20 * "*")
    with requests.Session() as s:
        s.headers.update(header)
        print("Logging in to {}".format(urls["HOME"]))
        payload = {'username': config["APP"]["USER"],
                   'password': config["APP"]["PASSWORD"]}
        p = s.post(urls["HOME"], data=payload)
        print("Downloading CSV")
        # r = s.get("https://screener-api.zacks.com/export.php", cookies={"TS01af1a9c": "0128dc3c314e3a8771cef59eb7c63d55b65fd2de7dc678efe984e20eed3128808e1ae2ed8922dd9a7bd6b89a09d88dbfc37ca1f411",
        #                                                                 "visid_incap_2879622": "BwU+ZzsqS3Os+J8ublZzWEqpS2QAAAAAQUIPAAAAAAAFPRKwcjDkXbZt70pngjF0",
        #                                                                 "_vwo_uuid_v2": "DD95DFD4653D61E27B7B69C86F1EA955A|1c60d0ee5cf96e0ed67e10f819b9f69e",
        #                                                                 "AMCVS_3064401053DB594D0A490D4C%40AdobeOrg": "1",
        #                                                                 "_vis_opt_s": "1%7C",
        #                                                                 "_vis_opt_test_cookie": "1",
        #                                                                 "s_v17": "DEF", "s_cc": "true",
        #                                                                 "_vwo_uuid": "DD95DFD4653D61E27B7B69C86F1EA955A",
        #                                                                 "_vwo_ds": "3%241682680139%3A80.47738532%3A%3A",
        #                                                                 "_fbp": "fb.1.1682680140754.1256588369",
        #                                                                 "cookie": "7542dcdc-31ae-4937-a702-3b9efacc153a",
        #                                                                 "__qca": "P0-2018446936-1682680143486",
        #                                                                 "__gads": "ID=c05dafbe053f1108-22661f2eafdd0057:T=1682680140:S=ALNI_MYVpJ5mNveruYK7bsLTyOeLjrUzuw",
        #                                                                 "_au_1d": "AU1D-0100-001682680144-HJQUJ5J7-39QL",
        #                                                                 "_au_last_seen_pixels": "eyJhcG4iOjE2ODI2ODAxNDQsInR0ZCI6MTY4MjY4MDE0NCwicHViIjoxNjgyNjgwMTQ0LCJydWIiOjE2ODI2ODAxNDQsInRhcGFkIjoxNjgyNjgwMTQ0LCJhZHgiOjE2ODI2ODAxNDQsImdvbyI6MTY4MjY4MDE0NCwiYWRvIjoxNjgyNjgwMTQ0LCJzb24iOjE2ODI2ODAxNDQsImltcHIiOjE2ODI2ODAxNDR9",
        #                                                                 "_ga": "GA1.1.1697935371.1682680140",
        #                                                                 "PHPSESSID": "p1cuq3hfv5d5208f1nd195v51t",
        #                                                                 "visid_incap_2879609": "LgecmF6GTlWap3NufrL8hnKpS2QAAAAAQUIPAAAAAADK0n+SFzmc5TrhsX8b+jgH",
        #                                                                 "_tt_enable_cookie": "1",
        #                                                                 "_ttp": "dzXG97RnlLUO9AS-d39hiz-Juq4",
        #                                                                 "AMCV_3064401053DB594D0A490D4C%40AdobeOrg": "77933605%7CMCIDTS%7C19478%7CMCMID%7C58174398938772432401021054286524967623%7CMCAAMLH-1683477550%7C6%7CMCAAMB-1683477550%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1682879950s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C4.5.1",
        #                                                                 "s_vnum": "1682888400307%26vn%3D3",
        #                                                                 "s_invisit": "true",
        #                                                                 "undefined_s": "First%20Visit",
        #                                                                 "cf60519feb1344e434b8444b746a915b": "cdaeeeba9b4a4c5ebf042c0215a7bb0e",
        #                                                                 "incap_ses_1456_2879609": "d2dlX+rJrQk11MWH0sA0FFiaTmQAAAAAcZH8M5LGOO6TTREmyDqBNA==",
        #                                                                 "__gpi": "UID=00000bf1a572c226:T=1682680140:RT=1682872920:S=ALNI_MbVcZKY9RBWMS1KvayTOdUjWxWFiw",
        #                                                                 "s_gapv": "Logged%20In",
        #                                                                 "incap_ses_433_2879609": "8sj7LnhRdUZopgS0CFQCBlWcTmQAAAAAKybE03SzWNMk6zcIFxGN4g==",
        #                                                                 "incap_ses_1456_2879622": "vU2CKzX6oDxLcMqH0sA0FAShTmQAAAAA5lV5GL1S6dVPq/LONUlMqw==",
        #                                                                 "_vwo_sn": "194490%3A1",
        #                                                                 "s_p42": "screening%3A%20stock-screener",
        #                                                                 "s_sq": "%5B%5BB%5D%5D",
        #                                                                 "cto_bundle": "nxgQ3F84VVpxZUFLNTRwYldPJTJCekJVTkVrbmpPdlBOR1hFT0xyZ3lJODYlMkY2VCUyQkpQWm1WY1IxUEhCdERkWUVKWk5FU3pUbmhOazhWVXRJSjcxNHFkVGRqam4lMkJuRFhhaFRIQmlUMDljQTBLNWlzWUxXN2RIMG45SWtuNjdOanhxSXdGazRNVDNKbCUyQjY4NkhIS0lJMGVQSDNIenNBJTNEJTNE",
        #                                                                 "s_nr": "1682874634807-Repeat",
        #                                                                 "CURRENT_POS": "my_screen",
        #                                                                 "_ga_MXXMZ1PBF7": "GS1.1.1682872750.3.1.1682874640.0.0.0"})
        r = s.get("https://screener-api.zacks.com/export.php", cookies={"TS01af1a9c": "0128dc3c314e3a8771cef59eb7c63d55b65fd2de7dc678efe984e20eed3128808e1ae2ed8922dd9a7bd6b89a09d88dbfc37ca1f411", "visid_incap_2879622": "BwU+ZzsqS3Os+J8ublZzWEqpS2QAAAAAQUIPAAAAAAAFPRKwcjDkXbZt70pngjF0", "_vwo_uuid_v2": "DD95DFD4653D61E27B7B69C86F1EA955A|1c60d0ee5cf96e0ed67e10f819b9f69e", "AMCVS_3064401053DB594D0A490D4C%40AdobeOrg": "1", "_vis_opt_s": "1%7C", "_vis_opt_test_cookie": "1", "s_v17": "DEF", "s_cc": "true", "_vwo_uuid": "DD95DFD4653D61E27B7B69C86F1EA955A", "_vwo_ds": "3%241682680139%3A80.47738532%3A%3A", "_fbp": "fb.1.1682680140754.1256588369", "cookie": "7542dcdc-31ae-4937-a702-3b9efacc153a", "__qca": "P0-2018446936-1682680143486", "__gads": "ID=c05dafbe053f1108-22661f2eafdd0057:T=1682680140:S=ALNI_MYVpJ5mNveruYK7bsLTyOeLjrUzuw", "_au_1d": "AU1D-0100-001682680144-HJQUJ5J7-39QL", "_au_last_seen_pixels": "eyJhcG4iOjE2ODI2ODAxNDQsInR0ZCI6MTY4MjY4MDE0NCwicHViIjoxNjgyNjgwMTQ0LCJydWIiOjE2ODI2ODAxNDQsInRhcGFkIjoxNjgyNjgwMTQ0LCJhZHgiOjE2ODI2ODAxNDQsImdvbyI6MTY4MjY4MDE0NCwiYWRvIjoxNjgyNjgwMTQ0LCJzb24iOjE2ODI2ODAxNDQsImltcHIiOjE2ODI2ODAxNDR9", "_ga": "GA1.1.1697935371.1682680140", "PHPSESSID": "p1cuq3hfv5d5208f1nd195v51t", "visid_incap_2879609": "LgecmF6GTlWap3NufrL8hnKpS2QAAAAAQUIPAAAAAADK0n+SFzmc5TrhsX8b+jgH", "_tt_enable_cookie": "1", "_ttp": "dzXG97RnlLUO9AS-d39hiz-Juq4", "AMCV_3064401053DB594D0A490D4C%40AdobeOrg": "77933605%7CMCIDTS%7C19478%7CMCMID%7C58174398938772432401021054286524967623%7CMCAAMLH-1683477550%7C6%7CMCAAMB-1683477550%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1682879950s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C4.5.1",
                                                                        "cf60519feb1344e434b8444b746a915b": "cdaeeeba9b4a4c5ebf042c0215a7bb0e",
                                                                        "incap_ses_1456_2879609": "d2dlX+rJrQk11MWH0sA0FFiaTmQAAAAAcZH8M5LGOO6TTREmyDqBNA==",
                                                                        "__gpi": "UID=00000bf1a572c226:T=1682680140:RT=1682872920:S=ALNI_MbVcZKY9RBWMS1KvayTOdUjWxWFiw",
                                                                        "incap_ses_433_2879609": "8sj7LnhRdUZopgS0CFQCBlWcTmQAAAAAKybE03SzWNMk6zcIFxGN4g==",
                                                                        "incap_ses_1456_2879622": "vU2CKzX6oDxLcMqH0sA0FAShTmQAAAAA5lV5GL1S6dVPq/LONUlMqw==",
                                                                        "cto_bundle": "nxgQ3F84VVpxZUFLNTRwYldPJTJCekJVTkVrbmpPdlBOR1hFT0xyZ3lJODYlMkY2VCUyQkpQWm1WY1IxUEhCdERkWUVKWk5FU3pUbmhOazhWVXRJSjcxNHFkVGRqam4lMkJuRFhhaFRIQmlUMDljQTBLNWlzWUxXN2RIMG45SWtuNjdOanhxSXdGazRNVDNKbCUyQjY4NkhIS0lJMGVQSDNIenNBJTNEJTNE",
                                                                        "s_p42": "portfolios%3A%20my-stock-portfolio",
                                                                        "s_vnum": "1682888400307%26vn%3D4",
                                                                        "s_invisit": "true",
                                                                        "undefined_s": "First%20Visit",
                                                                        "_ga_MXXMZ1PBF7": "GS1.1.1682877320.4.1.1682877321.0.0.0",
                                                                        "s_gapv": "Logged%20In",
                                                                        "s_nr": "1682877322127-Repeat",
                                                                        "s_sq": "zacksprod%3D%2526c.%2526a.%2526activitymap.%2526page%253Dportfolios%25253A%252520my-stock-portfolio%2526link%253DStock%252520Screener%2526region%253Dscreener-tab%2526pageIDType%253D1%2526.activitymap%2526.a%2526.c%2526pid%253Dportfolios%25253A%252520my-stock-portfolio%2526pidt%253D1%2526oid%253Dhttps%25253A%25252F%25252Fwww.zacks.com%25252Fportfolios%25252Fmy-stock-portfolio%25252F%25253Ficid%25253Dportfolios-portfolios-nav_tracking-zcom-ma%2526ot%253DA",
                                                                        "CURRENT_POS": "my_screen"})

        # r = s.get(urls["RANK1EXCEL"])
        bytes_data = r.content
        string_data = str(r.content, 'utf-8')
        data_io = StringIO(string_data)
    return data_io


def process_dataframe(df):
    # add index column
    df['Index'] = range(1, len(df) + 1)

    # get rid of unnamed index
    df.drop(df.filter(regex="Unname"), axis=1, inplace=True)

    # parse dates
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], format='%b %d,%Y')

    # parse numbers
    for col in numerical_cols:
        df[col] = df[col].str.rstrip('%')
        df[col] = df[col].str.replace("[^0-9.\-]", "").astype(float)

    # rank ascending
    for col in rank_ascend:
        df['score to {}'.format(col)] = df[col].rank(
            ascending=True, method='min') - 1

    # rank descending
    for col in rank_descend:
        df['score to {}'.format(col)] = df[col].rank(
            ascending=False, method='min') - 1

    # add total column
    df['Total score'] = df[score_columns].sum(axis=1)
    df['USA rankings'] = df['Total score'].rank(
        ascending=True, method='min') - 1

    # add results
    df['score to Results'] = df[results_data_columns].sum(axis=1)
    df['Results rankings'] = df['score to Results'].rank(
        ascending=True, method='min') - 1

    # add potential
    df['Difference'] = df['Results rankings'] - df['USA rankings']
    df['Potential ranking'] = df['Difference'].rank(
        ascending=False, method='min') - 1

    # reorder columns
    df = df[header_cols +
            list(sum(list(zip(data_columns, score_columns)), ())) + calculated_columns]

    return df


def save(df):
    df.index = df.index + 1
    for processed_fname in processed_fnames:
        print(
            f"Styling Excel {processed_fname}. (Please wait, this may take a couple of minutes.)")
        excel_writer = StyleFrame.ExcelWriter(processed_fname)
        font = utils.fonts.calibri
        # font = 'Courier New'
        sf = StyleFrame(df)
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
        sf.apply_headers_style(cols_to_style=['Total score', 'USA rankings'],
                               styler_obj=Styler(bg_color='#ffc1f5', wrap_text=False, font=font,
                                                 font_size=12))
        sf.apply_headers_style(cols_to_style=['score to Results', 'Results rankings'],
                               styler_obj=Styler(bg_color='#ffc000', wrap_text=False, font=font,
                                                 font_size=12))
        sf.apply_headers_style(cols_to_style=['Difference', 'Potential ranking'],
                               styler_obj=Styler(bg_color='#00ff69', wrap_text=False, font=font,
                                                 font_size=12))

        print(f"Saving excel {processed_fname}")
        try:
            sf.to_excel(
                excel_writer=excel_writer,
                best_fit=list(df.columns),
                # best_fit=header_cols[:-1],
                columns_and_rows_to_freeze='E2',
                row_to_add_filters=0,
                index=False  # Index Column Added Seperately
            )
            excel_writer.save()
            print("Excel \"{}\" Saved.".format(processed_fname))
        except PermissionError as PE:
            print("Could not save file. Please make sure the file: \"{}\" is closed.".format(
                PE.filename))
            print(PE)
            quit(-1)

        # print("Adjust Column Width for {}".format(processed_fname))
        # excel = Dispatch('Excel.Application')
        # wb = excel.Workbooks.Open(os.path.join(os.getcwd(), processed_fname))

        # Activate second sheet
        # excel.Worksheets(2).Activate()

        # Autofit column in active sheet
        # excel.ActiveSheet.Columns.AutoFit()

        # Set column width to 25
        # excel.ActiveSheet.Range("A:ZZ").ColumnWidth = 25

        # print("Saving again.")
        # Save changes in a new file
        # wb.SaveAs("D:\\output_fit.xlsx")
        # Or simply save changes in a current file
        # TODO: insert path
        # wb.Save()
        # wb.SaveAs(processed_fname)

        # print("Closing ExcelWriter.")
        # wb.Close()

    # print("Excel Saved")


if __name__ == '__main__':
    print(pyfiglet.figlet_format("Hello!!!"))
    print("Please select a menu option.")
    button1 = ttk.Button(root, text='Download and Score',
                         command=lambda: execute(download=True))
    button1.pack(side=tk.TOP)
    button2 = ttk.Button(root, text='Score Only',
                         command=lambda: execute(download=False))
    button2.pack(side=tk.BOTTOM)
    root.mainloop()
