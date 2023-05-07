import configparser
from datetime import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import traceback
import pandas as pd
from selenium import webdriver
import styleframe
from styleframe import Styler, StyleFrame

class App:
    def __init__(self, master):
        self.master = master
        master.title("Data Downloader")

        self.label = tk.Label(master, text="Enter URL:")
        self.label.pack()

        self.entry = tk.Entry(master)
        self.entry.pack()

        self.select_and_process_button =\
            tk.Button(master, text="Select and Process",
                      command=self.select_and_process)
        self.select_and_process_button.pack()

        self.download_button = tk.Button(
            master, text="Download", command=self.download)
        self.download_button.pack()
        # TODO: make history base dir
        config = configparser.ConfigParser()
        config.read('config.ini')
        self.timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        self.output_filename = 'stock_screener.xlsx'
        self.header_cols = ['Index', 'Ticker', 'Company Name', 'Last Close']
        self.numerical_cols = []
        self.date_cols = []
        self.rank_ascend = ['Zacks Rank', 'Momentum Score', "VGM Score",
                    'Growth Score', 'Value Score', 'Zacks Industry Rank',
                    'Current Avg Broker Rec']
        self.rank_descend = ['Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                        '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                        '% Price Change (YTD)']
        self.drop_cols = []

        # Defines the column orders
        self.data_columns = ['Zacks Rank', 'Momentum Score', "VGM Score",
                        'Growth Score', 'Value Score', 'Zacks Industry Rank',
                        'Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                        '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                        '% Price Change (YTD)', 'Current Avg Broker Rec']
        self.results_data_columns = ['Market Cap (mil)', 'Avg Volume', '% Price Change (1 Week)',
                                '% Price Change (4 Weeks)', '% Price Change (12 Weeks)',
                                '% Price Change (YTD)']

        self.score_columns = ['score to {}'.format(col) for col in self.data_columns]
        self.calculated_columns = ['Total score', 'USA rankings', 'score to Results', 'Results rankings',
                            'Difference', 'Potential ranking']
        
        self.pair_columns = list(zip(iter(self.data_columns), iter(self.score_columns)))


        def flatten(lst):
            return [item for sublist in lst for item in sublist]


        self.odd_pair_columns = flatten(self.pair_columns[1::2])
        self.even_pair_columns = flatten(self.pair_columns[0::2])


    def download(self):

        # Process file with Pandas
        # data = pd.read_csv(file_path)
        # result = data.groupby(['column_name']).sum()

        # # Save result to disk
        # result.to_csv('result.csv')

        self.label.config(text="Done!")


    def handle_exception(self, error_code: str, e: Exception):
        messagebox.showinfo(title="Scraper and Scorer",
                            message=f"An error has occurred while processing the data. Exiting.\nError Code: {error_code}")
        traceback.print_exc()
        sys.exit(-1)


    def select_and_process(self):
        file_path: str = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                    filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
        df = pd.read_csv(file_path)

        try:
            print("Processing Dataframe")
            df = self.process_dataframe(df)
        except Exception as e:
            self.handle_exception('@err-processing', e)
        try:
            self.save(df)
        except Exception as e:
            self.handle_exception('@err-saving', e)

        messagebox.showinfo(title="Scraper and Scorer",
                            message='Scoring Completed.')

    def process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        # add index column
        df['Index'] = range(1, len(df) + 1)

        # get rid of unnamed index
        df.drop(df.filter(regex="Unname"), axis=1, inplace=True)

        # parse dates
        for col in self.date_cols:
            df[col] = pd.to_datetime(df[col], format='%b %d,%Y')

        # parse numbers
        for col in self.numerical_cols:
            df[col] = df[col].str.rstrip('%')
            df[col] = df[col].str.replace("[^0-9.\-]", "").astype(float)

        # rank ascending
        for col in self.rank_ascend:
            df['score to {}'.format(col)] = df[col].rank(
                ascending=True, method='min') - 1

        # rank descending
        for col in self.rank_descend:
            df['score to {}'.format(col)] = df[col].rank(
                ascending=False, method='min') - 1

        # add total column
        df['Total score'] = df[self.score_columns].sum(axis=1)
        df['USA rankings'] = df['Total score'].rank(
            ascending=True, method='min') - 1

        # add results
        df['score to Results'] = df[self.results_data_columns].sum(axis=1)
        df['Results rankings'] = df['score to Results'].rank(
            ascending=True, method='min') - 1

        # add potential
        df['Difference'] = df['Results rankings'] - df['USA rankings']
        df['Potential ranking'] = df['Difference'].rank(
            ascending=False, method='min') - 1

        # reorder columns
        df = df[self.header_cols +
                list(sum(list(zip(self.data_columns, self.score_columns)), ())) + self.calculated_columns]

        return df
    
    def save(self, df: pd.DataFrame):
        df.index = df.index + 1
        print(
            f"Styling Excel {self.output_filename}. (Please wait, this may take a couple of minutes.)")
        excel_writer = styleframe.ExcelWriter(self.output_filename)
        font = styleframe.utils.fonts.calibri
        # font = 'Courier New'
        sf = StyleFrame(df)
        sf.apply_column_style(cols_to_style=(self.header_cols + self.calculated_columns),
                            styler_obj=Styler(bg_color=styleframe.utils.colors.white, wrap_text=False, font=font,
                                                font_size=12),
                            style_header=True)
        sf.apply_column_style(cols_to_style=self.odd_pair_columns,
                            styler_obj=Styler(bg_color='#ffe28a', wrap_text=False, font=font,
                                                font_size=12),
                            style_header=True)
        sf.apply_column_style(cols_to_style=self.even_pair_columns,
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

        print(f"Saving excel {self.output_filename}")
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
            print("Excel \"{}\" Saved.".format(self.output_filename))
        except PermissionError as PE:
            print("Could not save file. Please make sure the file: \"{}\" is closed.".format(
                PE.filename))
            print(PE)
            quit(-1)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
