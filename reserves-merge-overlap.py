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
electronic_filename = askopenfilename(title = "Select electronic inventory comparison filename")
covid_filename = askopenfilename(title = "Select physical titles with temporary COVID electronic portfolios filename")

start = datetime.datetime.now()



reserves_df = pd.read_csv(reserves_filename, quotechar='"', dtype={'MMS Id': 'str', 'Title (Complete)': 'str', 'Barcode': 'str'}, delimiter=',')

electronic_df = pd.read_csv(electronic_filename, quotechar='"', dtype={'MMS Id': 'str', 'Title (Complete)': 'str'}, delimiter=',')

reserves_df['Title (Complete)'] = reserves_df['Title (Complete)'].apply(lambda x: re.sub(r'\W', '', x))
print("Normalized reserves df\n")
#electronic_df['Title (Complete)'] = electronic_df['Title (Complete)'].apply(lambda x: re.sub(r'\W', '', x))

# covid_managed_set = requests.get()
# covid_df = pd.read_excel(covid_filename)


#covid_xml_file = open(covid_filename, "r+", encoding='utf-8')


#covid_xml = covid_xml_file.read()

#unicode_covid_xml = covid_xml.encode('unicode').decode('unicode')

columns = ['Title', 'MMS ID', 'Provider', 'URL']
covid_df = pd.DataFrame(columns=columns)
x = 0
print("Starting parsing COVID file\n")
for event, elem in et.iterparse(covid_filename):
    if x > 100:
        break
    if elem.tag == "record":
        # tree = et.parse(elem)
        # print(et.tostring(elem))
        bib_record = pym.parse_xml_to_array(io.StringIO(et.tostring(elem).decode()))
        a = ""
        b = ""
        c = ""

        mms_id = bib_record[0]['001']
        print("MMS ID: " + str(mms_id) + "\n")
        if 'a' in bib_record[0]['245']:
            a = bib_record[0]['245']['a']
        if 'b' in bib_record[0]['245']:
            b = bib_record[0]['245']['b']
        if 'c' in bib_record[0]['245']:
            c = bib_record[0]['245']['c']

        title = a + b + c
        #print(str(bib_record[0]) + "\n")
        # print("Title: " + str(title))
        mms_id = ""
        if '1' in bib_record[0]:
            mms_id = bib_record[0]['1']
        provider = ""
        url = ""
        if '856' in bib_record[0]:
            if 'u' in bib_record[0]['856']:
                url = bib_record[0]['856']['u']
            if 'z' in bib_record[0]['856']:
                provider = bib_record[0]['856']['z']

        covid_df = covid_df.append({'Title': title, 'MMS ID': mms_id, 'Provider': provider, 'URL': url}, ignore_index=True)


        # for sub_record in elem:
        # print(title)
        # bib_record = pym.parse_xml_to_array(io.StringIO(elem))
        # title = bib_record.get_fields('245')
        #
        elem.clear()
        x += 1

    # tree = et.ElementTree(et.fromstring(covid_xml))
    #


# root = tree.get_root()
#
# for record in root.findall('record'):


reserves_df.to_excel(oDir + '/Titles from Reserves List.xlsx', index=False)
electronic_df.to_excel(oDir + 'Titles from Electronic List.xlsx', index=False)
covid_df.to_excel(oDir + '/Parsed COVID Titles - Sample.xlsx', index=False)


end = datetime.datetime.now() - start


print("Execution time for creating COVID dataframe: " + str(end) + "\n")

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
start2 = datetime.datetime.now()
master = pd.merge(reserves_df, electronic_df, left_on=['Title (Complete)'], right_on=['Title (Complete)'], how='outer', indicator=True)
end2 = datetime.datetime.now() - start2


print("Execution time for merging with electronic: " + str(end2) + "\n")
#
electronic_overlap_df = master[master['_merge'] == 'both']

start3 = datetime.datetime.now()
master_2 = pd.merge(reserves_df, covid_df, left_on=['Title (Complete)'], right_on=['Title'], how='outer', indicator=True)

end3 = datetime.datetime.now() - start3
print("Execution time for merging with COVID: " + str(end3) + "\n")
covid_overlap_df = master[master['_merge'] == 'both']
#
start4 = datetime.datetime.now()
final_df = pd.merge(electronic_overlap_df, covid_overlap_df, on=['Title (Complete)'], how='outer')
end4 = datetime.datetime.now() - start4
print("Execution time for merging merged files: " + str(end4) + "\n")
# #
# #
# #
final_df.to_excel(oDir + '/Overlap Analsysis - Reserves.xlsx', index=False)
