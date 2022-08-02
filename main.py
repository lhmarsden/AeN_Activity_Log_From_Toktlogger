#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22nd 2022

Standalone program to create an activity log
without the rest of the Nansen Legacy logging system

@author: Luke Marsden
"""

import xlsxwriter
import requests
import json
import numpy as np
from datetime import datetime as dt
import make_xlsx as mx
import toktlogger_json_to_df as tl
import pandas as pd
import os.path

filename = 'activity_log'

i = 1
while os.path.exists(f"{filename}_{i}.xlsx"):
    print(i)
    i += 1

path = f"{filename}_{i}.xlsx"

toktlogger = 'toktlogger-bonnevie.hi.no' # My laptop VM toktlogger

print('\nPulling data from toktlogger')
data = tl.json_to_df(toktlogger)
metadata_df = tl.pull_metadata(toktlogger)

print('\nWriting XLSX file')
terms = list(data.columns)
field_dict = mx.make_dict_of_fields()
metadata = True
conversions = True # Include metadata sheet and conversions sheet
mx.write_file(path,terms,field_dict,metadata,conversions,data, metadata_df)

print('\nGenerated file:', path,'\n')
