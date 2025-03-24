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
