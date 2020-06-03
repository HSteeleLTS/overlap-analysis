#!/usr/bin/env python3
import sys

import requests
import json
import os
import time
import csv
import re
import datetime

from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np

# from django.utils.encoding import smart_str

oDir = "./Output"
if not os.path.isdir(oDir) or not os.path.exists(oDir):
    os.makedirs(oDir)


proquest_input_filename = askopenfilename(title = "Select input file")
overlap_filename = askopenfilename(title = "Select master list file for overlap")





proquest_df = pd.read_csv(proquest_input_filename, dtype={'MMS Id': 'str', 'Title (Normalized)': 'str', 'Barcode': 'str'}, delimiter=',')

overlap_df = pd.read_csv(overlap_filename, dtype={'MMS Id': 'str', 'Title (Normalized)': 'str'}, delimiter=',')
#
# proquest_df.columns = ['MMS Id', 'TitleProQuest', 'Vendor Name']
#
# overlap_df.columns = ['MMS Id', 'TitleNonProQuest', 'Vendor Name']

master = pd.merge(proquest_df, overlap_df, left_on=['Title (Normalized)'], right_on=['Title (Normalized)'], how='outer', indicator=True)


proquest_overlap_df = master[master['_merge'] == 'both']



proquest_overlap_df.to_excel(oDir + '/Overlap Analsysis.xlsx', index=False)
