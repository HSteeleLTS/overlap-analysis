import os
import requests
import json

import sys
import time
import csv
import re
import datetime

from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np

import urllib


def link_check_test(link):
    print(link)
    link_check_result = os.system('linkchecker -q "' + link + '"')
    if link_check_result == 0:
        print("Link check success")


    else:
        print("Link check failure")

def link_curl(link):
    print(link)

    #link = urllib.parse.quote_plus(link)
    link = re.sub(r'\'', '\\"', link)
    link = re.sub(r'([\[\]\{\}])', r'\\\1', link)
    curl_result = os.system('curl -g -s -o nul "' + link + '"')
    #curl_result = os.system('wget --no-check-certificate -q "' + link + '"')
    print("\n\n" + str(curl_result) + "\n\n")
    # print(curl_result)
    if curl_result != 0 and curl_result != 5:
        print("Fail")
        return("Fail")
    else:
        print("Success")
        return("Success")

        # print("Link: " + str(link))
        # print("curl result:\n" + str(curl_result) + "\n")
        # print ("Fail")
def link_check(row, link, broken_links, link_success_counter, link_broken_counter):

    print(link)
    #link = quote(link)
    link_check_result = os.system('linkchecker -q "' + link + '"')

    if link_check_result == 0:
        print("Link check success")
        row = row.set_value('Link Status', 'Success')
        link_success_counter += 1
    else:
        print("Link check failure")
        link_broken_counter += 1
        row = row.set_value('Link Status', 'Fail')
        broken_links = broken_links.append(row)


    return([broken_links, row, link_success_counter, link_broken_counter])
#status = os.system('linkchecker -q \"https://tufts.userservices.exlibrisgroup.com/view/uresolver/01TUN_INST/openurl?ctx_enc=info:ofi/enc:UTF-8&ctx_id=10_1&ctx_tim=2020-07-11T09%3A57%3A38IST&ctx_ver=Z39.88-2004&url_ctx_fmt=info:ofi/fmt:kev:mtx:ctx&url_ver=Z39.88-2004&rfr_id=info:sid/primo.exlibrisgroup.com-01TUN_ALMA&req_id=&rft_dat=ie=01TUN_INST/51258506960003851,ie=01TUN_INST:51258506960003851,language=eng,view=01TUN&svc_dat=viewit&u.ignore_date_coverage=true&env_type=&rft.local_attribute=&rft.format=1%20online%20resource%201%20volume..&rft.kind=\"')

#print(status)
#os.system('linkchecker -q https%3A%2F%2Ftufts.userservices.exlibrisgroup.com%2Fview%2Furesolver%2F01TUN_INST%2Fopenurl%3Fctx_enc%3Dinfo%3Aofi%2Fenc%3AUTF-8%26ctx_id%3D10_1%26ctx_tim%3D2020-07-11T09%253A57%253A38IST%26ctx_ver%3DZ39.88-2004%26url_ctx_fmt%3Dinfo%3Aofi%2Ffmt%3Akev%3Amtx%3Actx%26url_ver%3DZ39.88-2004%26rfr_id%3Dinfo%3Asid%2Fprimo.exlibrisgroup.com-01TUN_ALMA%26req_id%3D%26rft_dat%3Die%3D01TUN_INST%2F51258506960003851%2Cie%3D01TUN_INST%3A51258506960003851%2Clanguage%3Deng%2Cview%3D01TUN%26svc_dat%3Dviewit%26u.ignore_date_coverage%3Dtrue%26env_type%3D%26rft.local_attribute%3D%26rft.format%3D1%2520online%2520resource%25201%2520volume..%26rft.kind%3D')


reserves_filename = askopenfilename(title = "Select reserves filename")

start = datetime.datetime.now()

reserves_df = pd.read_excel(reserves_filename, dtype={'MMS Id': 'str', 'Publication Date': 'str', 'Title (Normalized)': 'str', 'Reading List Id': 'str', 'Citation Id': 'str', 'Citation Source': 'str'})

source_list = reserves_df['Citation Source'].to_list()

x = 0
fail_counter = 0
success_counter = 0
for source in source_list:
    # if x < 10:
    #     # link_check_test(source)
    print(str(type(source)))
    result = link_curl(str(source))

    if result == "Fail":
        fail_counter += 1
    # else:
    elif result == "Success":
        success_counter += 1

    #     break

    x += 1



print("Number of Failures: " + str(fail_counter))
print("Number of Sucesses: " + str(success_counter))
end = datetime.datetime.now() - start
# test_file.close()

print("Execution time: " + str(end) + "\n")
