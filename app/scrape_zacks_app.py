import configparser
import glob
import logging
import os
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd
import styleframe
from styleframe import StyleFrame, Styler

_logger = logging.getLogger("USA Model")
logging.basicConfig(level=logging.INFO,
                    # format='[%(asctime)s] {%(pathname)s:%(lineno)d} %(levelname)s - %(message)s',
                    format='[%(asctime)s] %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S')


def highlight(element):
    """Highlights (blinks) a Selenium Webdriver element."""
    driver = element._parent

    def apply_style(s):
        driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",
                              element, s)
    original_style = element.get_attribute('style')
    apply_style("background: yellow; border: 2px solid red;")
    time.sleep(.3)
    apply_style(original_style)


class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        master.title("USA Model")

        self.label = tk.Label(master, text="Welcome!")
        self.label.pack()

        self.download_button = tk.Button(
            master, text="Fetch New Data", command=self.download_and_process)
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
        self.calculated_columns = ['Total score', 'USA rankings', 'score to Results',
                                   'Results rankings', 'Difference', 'Potential ranking']

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

    def download_and_process(self) -> None:
        filepath = self.download()
        self.read_csv_and_process(filepath)

    def download(self) -> Path:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service as ChromeService
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.wait import WebDriverWait
        from seleniumrequests.request import RequestsSessionMixin
        from webdriver_manager.chrome import ChromeDriverManager

        class RequestsChromeWebDriver(RequestsSessionMixin, webdriver.Chrome):
            """A Chrome webdriver with requests functionality."""
            pass

        options = Options()
        options.headless = self.config.getboolean('APP', 'HEADLESS')
        options.add_experimental_option('prefs', {
            'download.default_directory': os.getcwd(),
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True
        })
        _logger.info("Installing chrome driver")
        self.label.config(text="Running WebDriver")
        with RequestsChromeWebDriver(options=options,
                                     service=ChromeService(ChromeDriverManager().install())) as driver:  # noqa: E501
            # current_dimension = driver.execute_script(
            #     "return [window.innerHeight, window.innerWidth];")
            _logger.info("Set window size.")
            new_dimension = {'width': 1150, 'height': 1000}
            driver.set_window_size(
                new_dimension['width'], new_dimension['height'])

            _logger.info("Getting homepage.")
            driver.get("https://www.zacks.com/my_account/welcomeback.php")

            _logger.info("Logging in.")
            username_input = driver.find_element(By.XPATH,
                                                 '/html/body/div[2]/div[3]/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input')  # noqa: E501
            assert username_input.is_displayed()
            username_input.send_keys(self.config['APP']['USER'])  # type: ignore
            password_input = driver.find_element(By.XPATH,
                                                 '/html/body/div[2]/div[3]/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input')  # noqa: E501
            assert password_input.is_displayed()
            password_input.send_keys(self.config['APP']['PASSWORD'])  # type: ignore
            login_button = driver.find_element(By.ID, 'button')
            assert login_button.is_displayed()
            login_button.click()

            if "account locked" in driver.find_element(By.XPATH, '/html/body').text.lower():
                raise RuntimeError(
                    "Account Locked. Please try again in 30 minutes.")

            _logger.info("Switching to My-Screen tab.")
            driver.get('https://www.zacks.com/screening/stock-screener')

            # switch to frame
            _logger.info("Switching to screenerContent iframe.")
            iframe = driver.find_element(By.ID, 'screenerContent')
            driver.switch_to.frame(iframe)

            _logger.info("Getting my-screen-tab")
            # get my-screen tab
            my_screen_button = driver.find_element(By.ID, 'my-screen-tab')
            my_screen_button.send_keys(" ")

            _logger.info("Running AVI MODEL USA STOCK.")
            run_button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, 'btn_run_99931')))

            run_button.click()
            _logger.info("Getting CSV.")
            csv_button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="screener_table_wrapper"]/div[1]/a[1]'))
            )

            # driver.find_element(
            #     By.XPATH, '//*[@id="screener_table_wrapper"]/div[1]/a[1]')
            csv_button.click()

            self.label.config(text="Downloading")

            # wait for the file to finish downloading and get its path
            while not any(filename.endswith('.csv') for filename in os.listdir(os.getcwd())):
                print("Waiting for file to download..")
                time.sleep(1)
            # all_files = os.listdir(os.getcwd())
            csv_files = glob.glob(f"{os.getcwd()}/*.csv")
            latest_file = max(csv_files, key=os.path.getctime)
            # filepath = os.path.join(os.getcwd(),
            #                         os.listdir('/path/to/download/directory')[0])
            return Path(latest_file)

    def select_and_process(self):
        file_path = Path(filedialog.askopenfilename(initialdir="/", title="Select file",
                                                    filetypes=(("csv files", "*.csv"), ("all files", "*.*"))))
        assert file_path.exists(
        ), f"Error: expected {file_path} to exist but it does not."
        # self.label.config(text=f"Reading \"{file_path}\"")
        # self.label.config(text=f"Processing \"{file_path}\"")
        self.read_csv_and_process(file_path)

    def read_csv_and_process(self, filepath: Path) -> pd.DataFrame:
        _logger.info("Reading CSV")
        df = pd.read_csv(filepath)
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

        self.save(df)
        messagebox.showinfo(title="Scraper and Scorer",
                            message=f'Scoring Completed!\nFile saved to {self.output_filename}')
        quit(0)

    def save(self, df: pd.DataFrame) -> None:
        df.index = df.index + 1
        _logger.info("Styling Excel")
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

        _logger.info(f"Saving Excel to: {self.output_filename}")
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
            _logger.info("Excel \"{}\" Saved.".format(self.output_filename))
        except PermissionError as PE:
            _logger.info(
                f"Could not save file. Please make sure the file: \"{PE.filename}\" is closed")
            _logger.info(PE)
            raise


def show_error(exc, val, tb) -> None:
    messagebox.showerror(
        "Error!", message=f"An unexpected error occurred!\n\n{str(val)}")


if __name__ == "__main__":
    root = tk.Tk()
    root.report_callback_exception = show_error
    MainApplication(root, width=300).pack(side="top", fill="both", expand=True)
    root.protocol("WM_DELETE_WINDOW", MainApplication.on_closing)
    root.mainloop()
