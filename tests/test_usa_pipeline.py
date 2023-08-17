import pandas as pd
import rootutils

from mypackage.app import PipelineConfig, USAModelPipeline

root_path = rootutils.find_root(search_from=__file__, indicator=".project-root")


def test_read_csv_and_process():
    input_path = root_path / "tests" / "assets" / "input.csv"
    processed_path = root_path / "tests" / "assets" / "processed.csv"
    processed_true = pd.read_csv(processed_path, index_col=0)
    config = PipelineConfig("mock_u", "mock_p", False)
    pipeline = USAModelPipeline(config)
    processed = pipeline.read_csv_and_process(input_path)
    pd.testing.assert_frame_equal(processed, processed_true)
