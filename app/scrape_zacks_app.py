import configparser
import logging
import os
import sys
import time
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox

import pandas as pd
import styleframe
from selenium import webdriver
from styleframe import StyleFrame, Styler

_logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        master.title("USA Model")

        self.label = tk.Label(master, text="Welcome!")
        self.label.pack()

        self.download_button = tk.Button(
            master, text="Fetch New Data", command=self.download)
        self.download_button.pack()

        self.select_and_process_button =\
            tk.Button(master, text="Use Existing Data",
                      command=self.select_and_process)
        self.select_and_process_button.pack()

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        self.timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        self.output_filename = f'USA-Model-{self.timestamp}.xlsx'
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

        self.score_columns = ['score to {}'.format(
            col) for col in self.data_columns]
        self.calculated_columns = ['Total score', 'USA rankings', 'score to Results', 'Results rankings',
                                   'Difference', 'Potential ranking']

        self.pair_columns = list(
            zip(iter(self.data_columns), iter(self.score_columns)))

        def flatten(lst):
            return [item for sublist in lst for item in sublist]

        self.odd_pair_columns = flatten(self.pair_columns[1::2])
        self.even_pair_columns = flatten(self.pair_columns[0::2])

    @staticmethod
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()

    def download(self):
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service as ChromeService
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.by import By
        from webdriver_manager.chrome import ChromeDriverManager

        options = Options()
        options.add_experimental_option('prefs', {
            'download.default_directory': '/path/to/download/directory',
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })
        _logger.info("Installing chrome driver")
        global driver
        driver = webdriver.Chrome(options=options, service=ChromeService(
            ChromeDriverManager().install()))

        _logger.info("Getting stock-screener homepage.")
        driver.get(
            "https://www.zacks.com/screening/stock-screener?icid=screening-screening-nav_tracking-zcom-main_menu_wrapper-stock_screener")

        current_dimension = driver.execute_script(
            "return [window.innerHeight, window.innerWidth];")
        height = current_dimension[0]
        width = current_dimension[1]
        _logger.info("Current height: {}".format(height))
        _logger.info("Current width: {}".format(width))
        # new_dimension = {'width': 1600, 'height': 600}
        # driver.set_window_size(new_dimension['width'], new_dimension['height'])

        _logger.info("Logging in.")
        self.label.config(text="Logging in.")
        login_popup = driver.find_element(
            By.XPATH, '//*[@id="mob_log_me_in"]/a')
        login_link = driver.find_element(By.XPATH, '//*[@id="log_me_in"]/a')
        if login_popup.is_displayed():
            login_popup.click()
        elif login_link.is_displayed():
            login_link.click()
        else:
            raise RuntimeError("No button was displayed")

        username_input = driver.find_element(By.ID, 'username')
        password_input = driver.find_element(By.ID, 'password')
        username_input.send_keys(self.config['APP']['USER'])
        password_input.send_keys(self.config['APP']['PASSWORD'])

        # submit the form
        login_button = driver.find_element(
            By.CSS_SELECTOR, 'input[type="submit"]')
        login_button.click()

        # TODO: raise exception if login fails
        if "account locked" in driver.find_element(By.XPATH, '/html/body').text.lower():
            raise RuntimeError("Account Locked.")

        _logger.info("Switching to My-Screen tab.")
        # switch to frame
        iframe = driver.find_element(
            By.XPATH, '/html/body/div[2]/div[3]/div/div/section/div/iframe')
        driver.switch_to.frame(iframe)

        # get my screen
        my_screen_button = driver.find_element(
            By.XPATH, '//*[@id="my-screen-tab"]')
        my_screen_button.click()

        # TODO: continue here
        _logger.info("Running AVI MODEL USA STOCK.")
        # switch to frame
        # iframe = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/section/div/iframe')
        # driver.switch_to.frame(iframe)

        run_button = driver.find_element(By.XPATH, '//*[@id="btn_run_99931"]')
        run_button.click()
        # create action chain object
        # action = ActionChains(driver)
        # perform the operation
        # action.move_to_element(run_button).click().perform()
        # run_button.click()

        _logger.info("Getting CSV.")
        csv_button = driver.find_element(
            By.XPATH, '//*[@id="screener_table_wrapper"]/div[1]/a[1]')
        csv_button.click()

        # wait for the file to finish downloading and get its path
        while not any(filename.endswith('.csv') for filename in os.listdir('/path/to/download/directory')):
            print("Waiting for file to download")
            time.sleep(1)
        filepath = os.path.join('/path/to/download/directory',
                                os.listdir('/path/to/download/directory')[0])
        driver.quit()
        # TODO: connect to processing script.
        # Process file with Pandas
        # data = pd.read_csv(file_path)
        # result = data.groupby(['column_name']).sum()

        # # Save result to disk
        # result.to_csv('result.csv')
        self.label.config(text="Downloaded!")
        quit(0)

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
                            message=f'Scoring Completed!\nFile saved to {self.output_filename}')
        quit(0)

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
            df[col] = df[col].str.replace(r"[^0-9.\-]", "").astype(float)

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
            quit(1)


if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root, width=300).pack(side="top", fill="both", expand=True)
    root.protocol("WM_DELETE_WINDOW", MainApplication.on_closing)
    root.mainloop()
