import sys

import requests
import json
import os
import sys
import time
import csv
import re
import datetime

from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np

course_code = "2208-80313"
course_record_url = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey=l7xx3a793454bd9b40e3ae9b433bea5709d0&q=code~{" + course_code + "}&format=json"
course_record = requests.get(course_record_url).text

course_record = json.loads(course_record)
#print(course_record)
course_id = course_record['course'][0]['id']
reading_list_id = "9664352350003851"
citation_id =     "9874611310003851"
citation_url = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses/" + str(course_id) + "/reading-lists/" + str(reading_list_id) + "/citations/" + str(citation_id) + "?apikey=l7xx3a793454bd9b40e3ae9b433bea5709d0&format=json"

citation_record = requests.get(citation_url).text

print(citation_record)
citation_record = json.loads(citation_record)

citation_title = citation_record['metadata']['title']
citation_author = citation_record['metadata']['author']

#print("Title: " + citation_title + "; Author: " + citation_author + "\n" )
