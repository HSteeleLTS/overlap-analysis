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

import xml.etree.cElementTree as et
import pymarc as pym
import io
# from django.utils.encoding import smart_str


oDir = "./Output"
if not os.path.isdir(oDir) or not os.path.exists(oDir):
    os.makedirs(oDir)


reserves_filename = askopenfilename(title = "Select reserves filename")
# electronic_filename = askopenfilename(title = "Select electronic inventory comparison filename")
# covid_filename = askopenfilename(title = "Select physical titles with temporary COVID electronic portfolios filename")

start = datetime.datetime.now()



#reserves_df = pd.read_csv(reserves_filename, quotechar='"', dtype={'MMS Id': 'str', 'Title (Complete)': 'str', 'Barcode': 'str'}, delimiter=',')

reserves_df = pd.read_excel(reserves_filename, dtype={'Barcode': 'str'})
pd.set_option('display.max_columns', None)


# electronic_df = pd.read_csv(electronic_filename, quotechar='"', dtype={'MMS Id': 'str', 'Title (Complete)': 'str'}, delimiter=',')
#

### KEEP
reserves_df['title_only'] = reserves_df['title'].apply(lambda x: re.sub(r'([^\/]+).+', r'\1', x).lower())
reserves_df['title_only'] = reserves_df['title_only'].apply(lambda x: re.sub(r'([^\s\w])', '', x).lower())
reserves_df['title_only'] = reserves_df['title_only'].apply(lambda x: re.sub(r'\s{2,}', r' ', x))

print(reserves_df.head())
### ###
sru_url_prefix = "https://tufts.alma.exlibrisgroup.com/view/sru/01TUN_INST?version=1.2&operation=searchRetrieve&recordSchema=marcxml&query=alma.title="


# print("Normalized reserves df\n")

# electronic_df['Title (Complete)'] = electronic_df['Title (Complete)'].apply(lambda x: re.sub(r'\W', '', x))

# covid_managed_set = requests.get()
# covid_df = pd.read_excel(covid_filename)


#covid_xml_file = open(covid_filename, "r+", encoding='utf-8')


#covid_xml = covid_xml_file.read()

#unicode_covid_xml = covid_xml.encode('unicode').decode('unicode')

# columns = ['Title', 'MMS ID', 'Provider', 'URL']
# covid_df = pd.DataFrame(columns=columns)

#### KEEP
x = 0
print("Starting reading through RESERVES file\n")
# for event, elem in et.iterparse(covid_filename):
while x < len(reserves_df.index):
    if x > 1:
        break

    title = reserves_df.iloc[x]['title_only']

    title = re.sub(r'\s', '%20', title)
    title = '%22' + title + '%22'
    record = requests.get(sru_url_prefix + title)
    print(title)
    print(record.text)
    x += 1
#### END KEEP

    #     covid_df = covid_df.append({'Title': title, 'MMS ID': mms_id, 'Provider': provider, 'URL': url}, ignore_index=True)
    #
    #
    #     # for sub_record in elem:
    #     # print(title)
    #     # bib_record = pym.parse_xml_to_array(io.StringIO(elem))
    #     # title = bib_record.get_fields('245')
    #     #
    #     elem.clear()
    #
    #
    # # tree = et.ElementTree(et.fromstring(covid_xml))
    #


# root = tree.get_root()
#
# for record in root.findall('record'):


# reserves_df.to_excel(oDir + '/Titles from Reserves List.xlsx', index=False)
# electronic_df.to_excel(oDir + 'Titles from Electronic List.xlsx', index=False)
# covid_df.to_excel(oDir + '/Parsed COVID Titles - Sample.xlsx', index=False)
#
#
# end = datetime.datetime.now() - start
#
#
# print("Execution time for creating COVID dataframe: " + str(end) + "\n")

#     x += 1
#

#
#
#
# for bib_record in bib_records:
#     if x > 5:
#         break
#     title = bib_record.get_fields('245')
#     print("Title: "  + str(title) +"\n")
#
#     x += 1
#
#
#
#
# # reserves_df['Title (Complete)'] = reserves_df['Title (Complete)'].apply(lambda x: x.lower())
# # electronic_df['Title (Complete)'] = electronic_df['Title (Complete)'].apply(lambda x: x.lower())
#
# print(reserves_df['Title (Complete)'])
# print(electronic_df['Title (Complete)'])
# print(covid_df['Title'])
# # #
# print(covid_df['title'])

#covid_sample = covid_df.head()

#covid_sample.to_excel(oDir + '/test.xlsx', index=False)
#
# proquest_df.columns = ['MMS Id', 'TitleProQuest', 'Vendor Name']
#
# overlap_df.columns = ['MMS Id', 'TitleNonProQuest', 'Vendor Name']


#### KEEP - MAYBE
# start2 = datetime.datetime.now()
# master = pd.merge(reserves_df, electronic_df, left_on=['Title (Complete)'], right_on=['Title (Complete)'], how='outer', indicator=True)
# end2 = datetime.datetime.now() - start2
#
#
# print("Execution time for merging with electronic: " + str(end2) + "\n")
# #
# electronic_overlap_df = master[master['_merge'] == 'both']
#
# start3 = datetime.datetime.now()
# master_2 = pd.merge(reserves_df, covid_df, left_on=['Title (Complete)'], right_on=['Title'], how='outer', indicator=True)
#
# end3 = datetime.datetime.now() - start3
# print("Execution time for merging with COVID: " + str(end3) + "\n")
# covid_overlap_df = master[master['_merge'] == 'both']
# #
# start4 = datetime.datetime.now()
# final_df = pd.merge(electronic_overlap_df, covid_overlap_df, on=['Title (Complete)'], how='outer')
# end4 = datetime.datetime.now() - start4
# print("Execution time for merging merged files: " + str(end4) + "\n")
#
# final_df.to_excel(oDir + '/Overlap Analsysis - Reserves.xlsx', index=False)


#### END KEEP - MAYBE
