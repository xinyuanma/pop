from re import L
from time import time
import time
import numpy as np
import pandas as pd
import warnings
import os
import xlwings as xw
import function
warnings.filterwarnings("ignore")  # 取消警告

data_path = os.getcwd() + "/Data"
PCMapping = pd.read_excel(data_path + "/Profit Center Hierarchy Flattened (MDG).xlsx")
popdata = data_path + "/POPData"
popyearlydata1 = data_path + "/Data1"
popyearlydata2 = data_path + "/Data2"

function.confirmversion()
pop_period = function.pop_period
pop_version = function.pop_version
period_list = function.period_list