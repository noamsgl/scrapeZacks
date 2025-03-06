from __future__ import annotations

import pandas as pd
import pytest
import rootutils

from mypackage.main import (AbstractModelPipeline, ETFModelPipeline,
                            PipelineConfig, USAModelPipeline)

root_path = rootutils.find_root(search_from=__file__, indicator=".project-root")


@pytest.mark.parametrize("pipeline_class,assets_folder", [(USAModelPipeline, "usa"), (ETFModelPipeline, "etf")])
def test_read_csv_and_process(pipeline_class: type[AbstractModelPipeline], assets_folder: str):
    input_path = root_path / "tests" / "assets" / assets_folder / "input.csv"
    config = PipelineConfig("mock_user", "mock_password", False)
    pipeline = pipeline_class(config)
    processed = pipeline.read_csv_and_process(input_path)
    processed_path = root_path / "tests" / "assets" / assets_folder / "processed.csv"
    processed_true = pd.read_csv(processed_path, index_col=0)
    pd.testing.assert_frame_equal(processed, processed_true)
