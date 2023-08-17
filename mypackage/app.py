import glob
import logging
import os
import time
import tkinter as tk
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List

import pandas as pd
# import pandera as pa
import rootutils
import styleframe
from jsonargparse import CLI, ArgumentParser
from styleframe import StyleFrame, Styler

from mypackage import __version__, utils

_logger = logging.getLogger("USA Model")
logging.basicConfig(
    level=logging.INFO,
    # format='[%(asctime)s] {%(pathname)s:%(lineno)d} %(levelname)s - %(message)s',
    format="[%(asctime)s] %(levelname)s - %(message)s",
    datefmt="%H:%M:%S",
)


@dataclass
class PipelineConfig:
    user: str
    password: str
    headless: bool = True


class AbstractModelPipeline(ABC):
    @abstractmethod
    def download(self):
        pass

    @abstractmethod
    def style_as_excel_and_save(self, df: pd.DataFrame) -> None:
        pass

    @abstractmethod
    def read_csv_and_process(self, filepath: Path) -> pd.DataFrame:
        pass

    def download_and_process(self) -> None:
        """The main function of "Fetch New Data".
        """
        filepath = self.download()
        processed = self.read_csv_and_process(filepath)
        self.style_as_excel_and_save(processed)
        quit(0)

    def select_and_process(self):
        """The main function of "Use Existing Data".
        """
        file_path = Path(
            filedialog.askopenfilename(
                initialdir="/",
                title="Select file",
                filetypes=(("csv files", "*.csv"), ("all files", "*.*")),
            )
        )
        assert (
            file_path.exists()
        ), f"Error: expected {file_path} to exist but it does not."
        processed = self.read_csv_and_process(file_path)
        self.style_as_excel_and_save(processed)
        quit(0)


class USAModelPipeline(AbstractModelPipeline):
    def __init__(self, config: PipelineConfig) -> None:
        self.config = config
        self.timestamp = datetime.now().strftime("%d-%m-%Y %H-%M")
        # FUTURE: replace with this
        # self.output_filename = filedialog.askopenfile(filetypes=filetypes, initialdir="#Specify the file path")
        self.output_filename = f"USA-Model-{self.timestamp}.xlsx"
        self.header_cols = ["Index", "Ticker", "Company Name", "Last Close"]
        self.numerical_cols: List[str] = []
        self.date_cols: List[str] = []
        self.rank_ascend = [
            "Zacks Rank",
            "Momentum Score",
            "VGM Score",
            "Growth Score",
            "Value Score",
            "Zacks Industry Rank",
            "Current Avg Broker Rec",
        ]
        self.rank_descend = [
            "Market Cap (mil)",
            "Avg Volume",
            "% Price Change (1 Week)",
            "% Price Change (4 Weeks)",
            "% Price Change (12 Weeks)",
            "% Price Change (YTD)",
        ]
        self.drop_cols: List[str] = []

        # Defines the column orders
        self.data_columns = [
            "Zacks Rank",
            "Momentum Score",
            "VGM Score",
            "Growth Score",
            "Value Score",
            "Zacks Industry Rank",
            "Market Cap (mil)",
            "Avg Volume",
            "% Price Change (1 Week)",
            "% Price Change (4 Weeks)",
            "% Price Change (12 Weeks)",
            "% Price Change (YTD)",
            "Current Avg Broker Rec",
        ]

        # These are columns to be used in calculating fundamental statistics
        self.fundamentals_scores_columns = [
            f"score to {col}"
            for col in (
                "Market Cap (mil)",
                "Avg Volume",
                "% Price Change (1 Week)",
                "% Price Change (4 Weeks)",
                "% Price Change (12 Weeks)",
                "% Price Change (YTD)",
            )
        ]

        self.score_columns = ["score to {}".format(col) for col in self.data_columns]
        self.score_columns = ["score to {}".format(col) for col in self.data_columns]
        self.calculated_columns = [
            "Total score",
            "USA rankings",
            "Fundamentals Ranks Sum",
            "Results rankings",
            "Difference",
            "Potential ranking",
        ]

        self.pair_columns = list(zip(iter(self.data_columns), iter(self.score_columns)))

        self.odd_pair_columns = utils.flatten(self.pair_columns[1::2])
        self.even_pair_columns = utils.flatten(self.pair_columns[0::2])

    def download(self) -> Path:
        """Download data.

        Raises:
            RuntimeError: if downloading fails

        Returns:
            Path: path to downloaded csv file
        """
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
        options.headless = self.config.headless
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": os.getcwd(),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            },
        )
        _logger.info("Installing chrome driver")
        with RequestsChromeWebDriver(
            options=options, service=ChromeService(ChromeDriverManager(version="114.0.5735.16").install())
        ) as driver:  # noqa: E501
            # current_dimension = driver.execute_script(
            #     "return [window.innerHeight, window.innerWidth];")
            _logger.info("Set window size.")
            new_dimension = {"width": 1150, "height": 1000}
            driver.set_window_size(new_dimension["width"], new_dimension["height"])

            _logger.info("Getting homepage.")
            driver.get("https://www.zacks.com/my_account/welcomeback.php")

            _logger.info("Logging in.")
            username_input = driver.find_element(
                By.XPATH,
                "/html/body/div[2]/div[3]/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input",  # noqa: E501
            )  # noqa: E501
            assert username_input.is_displayed()
            username_input.send_keys(self.config.user)  # type: ignore
            password_input = driver.find_element(
                By.XPATH,
                "/html/body/div[2]/div[3]/div/div/div/div/div[2]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input",  # noqa: E501
            )  # noqa: E501
            assert password_input.is_displayed()
            password_input.send_keys(self.config.password)  # type: ignore
            login_button = driver.find_element(By.ID, "button")
            assert login_button.is_displayed()
            login_button.click()

            if (
                "account locked"
                in driver.find_element(By.XPATH, "/html/body").text.lower()
            ):
                raise RuntimeError("Account Locked. Please try again in 30 minutes.")

            if (
                "failed"
                in driver.find_element(By.XPATH, "/html/body").text.lower()
            ):
                raise RuntimeError("Sign in failed.")

            _logger.info("Switching to My-Screen tab.")
            driver.get("https://www.zacks.com/screening/stock-screener")

            # switch to frame
            _logger.info("Switching to screenerContent iframe.")
            iframe = driver.find_element(By.ID, "screenerContent")
            driver.switch_to.frame(iframe)

            _logger.info("Getting my-screen-tab")
            # get my-screen tab
            my_screen_button = driver.find_element(By.ID, "my-screen-tab")
            my_screen_button.send_keys(" ")

            _logger.info("Running AVI MODEL USA STOCK.")
            run_button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "btn_run_99931"))
            )

            run_button.click()
            _logger.info("Getting CSV.")
            csv_button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="screener_table_wrapper"]/div[1]/a[1]')
                )
            )

            # driver.find_element(
            #     By.XPATH, '//*[@id="screener_table_wrapper"]/div[1]/a[1]')
            pre_existing_csvs = glob.glob(f"{os.getcwd()}/*.csv")
            csv_button.click()

            # wait for the file to finish downloading and get its path
            while True:
                current_csvs = glob.glob(f"{os.getcwd()}/*.csv")
                if len(current_csvs) > len(pre_existing_csvs):
                    break
                print("Waiting for file to download..")
                time.sleep(1)
            latest_file = max(current_csvs, key=os.path.getctime)
            _logger.info(f"Got file {latest_file}.")
            # filepath = os.path.join(os.getcwd(),
            #                         os.listdir('/path/to/download/directory')[0])
            return Path(latest_file)

    def read_csv_and_process(self, filepath: Path) -> pd.DataFrame:
        """Read a CSV from a file path, perform main operations on it, and save it.

        Args:
            filepath (Path): a path to the CSV data.

        Returns:
            processed (pd.DataFrame): a processed dataframe
        """
        _logger.info("Reading CSV")
        df = pd.read_csv(filepath)
        # add index column
        df["Index"] = range(1, len(df) + 1)

        # get rid of unnamed index
        df.drop(df.filter(regex="Unname"), axis=1, inplace=True)

        # parse dates
        for col in self.date_cols:
            df[col] = pd.to_datetime(df[col], format="%b %d,%Y")

        # parse numbers
        for col in self.numerical_cols:
            df[col] = df[col].str.rstrip("%")
            df[col] = df[col].str.replace(r"[^0-9.\-]", "").astype(float)

        # rank ascending
        for col in self.rank_ascend:
            df["score to {}".format(col)] = (
                df[col].rank(ascending=True, method="min") - 1
            )

        # rank descending
        for col in self.rank_descend:
            df["score to {}".format(col)] = (
                df[col].rank(ascending=False, method="min") - 1
            )

        # add total column
        df["Total score"] = df[self.score_columns].sum(axis=1)
        df["USA rankings"] = df["Total score"].rank(ascending=True, method="min")

        # add results
        df["Fundamentals Ranks Sum"] = df[self.fundamentals_scores_columns].sum(axis=1)
        df["Results rankings"] = (
            df["Fundamentals Ranks Sum"].rank(ascending=True, method="min")
        )

        # add potential
        df["Difference"] = df["Results rankings"] - df["USA rankings"]
        df["Potential ranking"] = (
            df["Difference"].rank(ascending=False, method="min") - 1
        )

        # reorder columns
        df = df[
            self.header_cols
            + list(sum(list(zip(self.data_columns, self.score_columns)), ()))
            + self.calculated_columns
        ]
        return df

    def style_as_excel_and_save(self, df: pd.DataFrame) -> None:
        df.index = df.index + 1
        _logger.info("Styling Excel")
        excel_writer = styleframe.ExcelWriter(self.output_filename)
        font = styleframe.utils.fonts.calibri
        # font = 'Courier New'
        sf = StyleFrame(df)
        sf.apply_column_style(
            cols_to_style=(self.header_cols + self.calculated_columns),
            styler_obj=Styler(
                bg_color=styleframe.utils.colors.white,
                wrap_text=False,
                font=font,
                font_size=12,
            ),
            style_header=True,
        )
        sf.apply_column_style(
            cols_to_style=self.odd_pair_columns,
            styler_obj=Styler(
                bg_color="#ffe28a", wrap_text=False, font=font, font_size=12
            ),
            style_header=True,
        )
        sf.apply_column_style(
            cols_to_style=self.even_pair_columns,
            styler_obj=Styler(
                bg_color="#d6ffba", wrap_text=False, font=font, font_size=12
            ),
            style_header=True,
        )
        sf.apply_headers_style(
            cols_to_style=["Total score", "USA rankings"],
            styler_obj=Styler(
                bg_color="#ffc1f5", wrap_text=False, font=font, font_size=12
            ),
        )
        sf.apply_headers_style(
            cols_to_style=["Fundamentals Ranks Sum", "Results rankings"],
            styler_obj=Styler(
                bg_color="#ffc000", wrap_text=False, font=font, font_size=12
            ),
        )
        sf.apply_headers_style(
            cols_to_style=["Difference", "Potential ranking"],
            styler_obj=Styler(
                bg_color="#00ff69", wrap_text=False, font=font, font_size=12
            ),
        )

        _logger.info(f"Saving Excel to: {self.output_filename}")
        try:
            sf.to_excel(
                excel_writer=excel_writer,
                best_fit=list(df.columns),
                # best_fit=header_cols[:-1],
                columns_and_rows_to_freeze="E2",
                row_to_add_filters=0,
                index=False,  # Index Column Added Seperately
            )
            excel_writer.save()
            _logger.info('Excel "{}" Saved.'.format(self.output_filename))
        except PermissionError as PE:
            _logger.info(
                f'Could not save file. Please make sure the file: "{PE.filename}" is closed'
            )
            _logger.info(PE)
            raise
        messagebox.showinfo(
            title="Scraper and Scorer",
            message=f"Scoring Completed!\nFile saved to {self.output_filename}",
        )

# define schema (initial skeleton)
# usa_model_schema = pa.DataFrameSchema({
#     "column1": pa.Column(int, checks=pa.Check.le(10)),
#     "column2": pa.Column(float, checks=pa.Check.lt(-1.2)),
#     "column3": pa.Column(str, checks=[
#         pa.Check.str_startswith("value_"),
#         # define custom checks as functions that take a series as input and
#         # outputs a boolean or boolean Series
#         pa.Check(lambda s: s.str.split("_", expand=True).shape[1] == 2)
#     ]),
# })


class MainApplication(tk.Frame):
    def __init__(self, master: tk.Tk, config: PipelineConfig, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        master.title(f"USA Model - {__version__}")

        self.label = tk.Label(master, text="Welcome!")
        self.label.pack()

        self.pipeline = USAModelPipeline(config)

        self.download_button = tk.Button(master, text="Fetch New Data", command=self.pipeline.download_and_process)
        self.download_button.pack()

        self.select_and_process_button = tk.Button(
            master, text="Use Existing Data", command=self.pipeline.select_and_process)
        self.select_and_process_button.pack()

    @staticmethod
    def get_config_parser():
        parser = ArgumentParser()
        parser.add_dataclass_arguments(PipelineConfig, "pipeline")
        return parser

    @staticmethod
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()


def show_error(exc, val, tb) -> None:
    messagebox.showerror(
        "Error!", message=f"An unexpected error occurred!\n\n{str(val)}"
    )


if __name__ == "__main__":
    root_path = rootutils.find_root(search_from=__file__, indicator='.project-root')
    cfg: PipelineConfig = CLI(PipelineConfig,
                              as_positional=False,
                              default_config_files=[root_path / "config.yaml"])
    root = tk.Tk()
    root.report_callback_exception = show_error
    MainApplication(root, cfg, width=300).pack(side="top", fill="both", expand=True)
    root.protocol("WM_DELETE_WINDOW", MainApplication.on_closing)
    root.mainloop()
