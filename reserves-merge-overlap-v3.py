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

# from openpyxl import load_workbook



oDir = "./Output"
if not os.path.isdir(oDir) or not os.path.exists(oDir):
    os.makedirs(oDir)



# electronic_filename = askopenfilename(title = "Select electronic inventory comparison filename")
# covid_filename = askopenfilename(title = "Select physical titles with temporary COVID electronic portfolios filename")



semester = input("Enter the course code prefix for the semester you'd like analyze: ")

semester = str(semester)

reserves_filename = askopenfilename(title = "Select reserves filename")

start = datetime.datetime.now()
#reserves_df = pd.read_csv(reserves_filename, quotechar='"', dtype={'MMS Id': 'str', 'Title (Complete)': 'str', 'Barcode': 'str'}, delimiter=',')

reserves_df = pd.read_excel(reserves_filename, dtype={'MMS Id': 'str', 'Publication Date': 'str', 'Title (Normalized)': 'str'})

reserves_df_total = reserves_df.copy()
reserves_df_total = reserves_df_total.dropna(subset=['MMS Id'])
# reserves_df_total = reserves_df_total.dropna()

reserves_df_total['Title (Normalized)'] = reserves_df_total['Title (Normalized)'].apply(lambda x: x.lower())

# print(reserves_df_total)

pd.set_option('display.max_columns', None)

# print(reserves_df_total)
reserves_df_total['Semester'] = reserves_df_total['Course Code'].apply(lambda x: re.sub(r'^(\d{4}).+', r'\1', x))
reserves_df_selected = reserves_df_total[reserves_df_total['Semester'] == semester]

reserves_df_selected = reserves_df_selected.sort_values(by=['Processing Department', 'Course Code'])
# print(reserves_df_selected)
column_list = reserves_df_selected.columns
proc_dept_df_list = []
proc_dept_list = []
for proc_dept, proc_dept_df in reserves_df_selected.groupby('Processing Department'):

    proc_dept_df_list.append(proc_dept_df)
    proc_dept_list.append(proc_dept)

d = 0

print("got to line 68\n")
test_file = open(oDir + '/Test File.txt', 'w+', encoding='utf-8')
for dept in proc_dept_df_list:
    processing_dept = proc_dept_list[d]
    date = datetime.datetime.now().strftime("%m-%d-%Y")
    filename = oDir + '/Electronic Copy Counts for - ' + processing_dept + ' - ' + date + '.xlsx'

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    workbook = writer.book
    # counts_sheet = workbook.add_worksheet('Counts')
    #
    #
    #
    # in_repo = workbook.add_worksheet('InRepo')
    # in_repo_different_year = workbook.add_worksheet('InRepoDifferentYear')
    #
    # temporary_collection_sheet = workbook.add_worksheet('TempColl')
    # temporary_collection_diff_year_sheet = workbook.add_worksheet('TempCollDiffYear')
    #
    # to_order_sheet = workbook.add_worksheet('ForPurchase')

    if d >= 1:
        break
    course_df_list = []
    # reserves_df_selected = reserves_df_selected[reserves_df_selected['Processing Department'] == 'Hirsh Health Sciences Reserves']
    for course, course_df in proc_dept_df_list[d].groupby('Course Code'):
        course_df_list.append(course_df)

    ## testing


    x = 0

    counts_df = pd.DataFrame(columns=['Processing Department', 'Course Name', 'Course Code', 'Books on Course', 'Physical Books on Course', 'No Ebook in Collection', 'Electronic Copies for Physical on Course', 'COVID Temporary Electronic Copies for Physical on Course - Subset', 'Electronic Copies for Physical on Course - Different Year', 'Ebook Copy for Physical in Repo - Add', 'Ebook Copy for Physical in Repo - Potentially Add - Different Year'])
    ebooks_to_add = pd.DataFrame(columns=['MMS ID', 'Title', 'Author', 'Publication Year', 'URL or Collection'])
    ebooks_to_add_different_year = pd.DataFrame(columns=['MMS ID', 'Title', 'Author', 'Publication Year', 'URL or Collection'])
    ebooks_we_need = pd.DataFrame(columns=column_list)
    #temporary_physical_record_ebooks = pd.DataFrame(columns=column_list)

    sru_url_prefix = "https://tufts.alma.exlibrisgroup.com/view/sru/01TUN_INST?version=1.2&operation=searchRetrieve&recordSchema=marcxml&query="
    # covid_column_list = column_list
    # covid_column_list = covid_column_list.append(['URL or Collection'])

    covid_e_books_df = pd.DataFrame(columns=column_list)
    covid_e_books_df['URL or Collection'] = ""
    covid_e_books_near_match_df = covid_e_books_df.copy()
    while x < len(course_df_list):

        print("Got into course_df_list loop")
        # print(course_df_list)
        # print("\n\n\n\n\n")
        # print(course_df_list[x])
        course_name = course_df_list[x]['Course Name']

        if x >= 1:
            break
        y = 0
        temporary_collections_portfolio_counter = 0
        temporary_collections_portfolio_counter_on_course = 0
        temporary_collections_portfolio_counter_near_match = 0
        ebook_counter = 0
        ebook_for_physical_counter = 0
        different_year_ebook_for_physical_counter = 0
        ebook_match_on_list_counter_without_year = 0

        ebook_match_on_list_counter = 0
        ebooke_match_on_list_counter_without_year = 0
        no_match_counter = 0
        number_of_books_on_course = len(course_df_list[x])


        date = datetime.datetime.now().strftime("%m/%d/%Y")

        #counts_file = open(oDir + '/Counts - ' + str(course_name) + ' - ' + date + '.txt', 'w+')
        counts_per_course_list = []

        # print("Items on course: " + str(len(course_df_list[x])))
        course_ebook_count = 0
        course_code = ""


        physical_list = course_df_list[x][course_df_list[x]['Resource Type'] == 'Book - Physical']


        electronic_list = course_df_list[x][course_df_list[x]['Resource Type'] == 'Book - Electronic']

        number_of_physical_books_on_course = len(physical_list)
        # print(physical_list)
        if len(physical_list) > 0:

            while y < len(physical_list):
                match = False
                course_df = physical_list.copy()
                # print(course_df)
                # if pd.isnull(course_df['Title (Normalized)']):
                #     print("Null: " + course_df[y]['Title (Normalized)'])
                # print(physical_list)
                # if x > 20:
                #     break
                print(course_df.iloc[y])
                test_file.write(str(course_df.iloc[y]) + "\n")

                # if y > 2:
                #     break
                # print("Resource Type: " + str(course_df.iloc[y]['Resource Type']))


                title = course_df.iloc[y]['Title (Normalized)']
                title_for_query = '\"' + re.sub(r'\s', '%20', title) + '\"'
                author = ""
                # print(str(author) + "Type: " + str(type(author)))
                if course_df.iloc[y]['Author'] != "" and course_df.iloc[y]['Author'] is not np.nan:
                    # author = re.sub(r'([A-Za-z0-9\.]+,\s+[A-Za-z0-9\.]+).+', r'\1', course_df.iloc[y]['Author'])
                    author = course_df.iloc[y]['Author'].lower()
                elif course_df.iloc[y]['Author (contributor)'] != "" and course_df.iloc[y]['Author (contributor)'] is not np.nan:
                    # author = re.sub(r'([A-Za-z0-9\.]+,\s+[A-Za-z0-9\.]+).+', r'\1', course_df.iloc[y]['Author (contributor)'])
                    author = course_df.iloc[y]['Author (contributor)'].lower()
                #author_for_query = '\"' + re.sub(r'\s', '%20', author) + '\"'
                year = course_df.iloc[y]['Publication Date']
                year = re.sub(r'\D', '', year)
                mms_id = course_df.iloc[y]['MMS Id']
                # if course_df.iloc[y]['Resource Type'] == 'Book - Electronic':
                #     ebook_counter += 1
                #     ebooks_on_this_course.loc[len(ebooks_on_this_course)] = course_df.iloc[y]
                #
                #     course_ebook_count += 1
                #     continue
                # elif course_df.iloc[y]['Resource Type'] == 'Book - Physical':
                #     if len(course_df[(course_df['Title (Normalized)'] == title) and ([course_df['Author'] = author) and (course_df['Publication Year'] == 'year')] > 1:
                #         ebook_for_physical_counter += 1
                #         continue
                #     elif len(course_df[(course_df['Title (Normalized)'] == title) and ([course_df['Author'] = author)] > 1:
                #         different_year_ebook_for_physical_counter += 1
                electronic_match = electronic_list[(electronic_list['Title (Normalized)'] == title) & ((electronic_list['Author'] == author) | (electronic_list['Author (contributor)'] == author))  & (electronic_list['Publication Date'] == year)]
                electronic_match_without_year = electronic_list[(electronic_list['Title (Normalized)'] == title) & ((electronic_list['Author'] == author) | (electronic_list['Author (contributor)'] == author)) ]
                quasi_match_bool = False

                if len(electronic_match) >= 1:
                    ebook_match_on_list_counter += 1
                    match = True
                    y += 1
                    continue
                elif len(electronic_match_without_year) >= 1:
                     ebook_match_on_list_counter_without_year += 1
                     different_year_ebook_for_physical_counter
                     quasi_match_bool = True
                     match = True


                else:

                    record_result = requests.get(sru_url_prefix + 'alma.title=' + title_for_query)
                    #print(record_result.content)
                    tree = et.ElementTree(et.fromstring(record_result.content))
                    root = tree.getroot()
                    # print(tree)
                    # print(root)
                    z = 0
                    # result_df = pd.DataFrame(columns=['MMS Id', 'Title', 'Author', 'Year'])
                    for elem in root.iter():


                        if re.match(r'.*record$', elem.tag):
                            # print(elem.tag)
                            # print(et.tostring(elem))
                            bib_record = pym.parse_xml_to_array(io.StringIO(et.tostring(elem).decode('utf-8')))
                            print(bib_record[0])
                            test_file.write(str(bib_record[0]) + "\n")
                            # print(course_df.iloc[y])
                            # print(bib_record[0])
                            if '001' in bib_record[0]:
                                # print("got into if statement")

                                # if bib_record[0]['001'] == mms_id:
                                #     z += 1
                                #     continue


                                # if '655' not in bib_record[0]:
                                #     z += 1
                                #     continue
                                # elif 'Electronic books' not in bib_record[0]['655']['a']:
                                #     continue
                                xml_mms_id = ""

                                # print("Data: " + str(bib_record[0]['001']['data']))
                                # print("Data value: " + bib_record[0]['001'].value())
                                # for t in bib_record[0]['001']:
                                #     print("T: " + str(t))
                                # if 'data' in bib_record[0]['001']:
                                #     for s in bib_record[0]['001']['data']:
                                #         print("S: " + str(s))

                                xml_mms_id = bib_record[0]['001'].value()
                                print('Analytics ID: ' + str(mms_id))
                                print('XML MMS ID:   ' + str(xml_mms_id))


                                xml_url = ""

                                if '856' in bib_record[0]:
                                    if 'u' in bib_record[0]['856']:
                                        xml_url = bib_record[0]['856']['u']
                                a = ""
                                b = ""
                                if 'a' in bib_record[0]['245']:
                                    a = bib_record[0]['245']['a']
                                if 'b' in bib_record[0]['245']:
                                    b = bib_record[0]['245']['b']

                                xml_title = a + b
                                xml_title = xml_title.lower()
                                xml_title = re.sub(r'\s\/$', '', xml_title)
                                xml_title = re.sub(r'\'', ' ', xml_title)
                                xml_title = re.sub(r'[^a-z0-9 ]', '', xml_title)
                                xml_title = re.sub(r'\.$', '', xml_title)

                                xml_author = ""

                                if '100' in bib_record[0]:
                                    if 'a' in bib_record[0]['100']:
                                        xml_author = bib_record[0]['100']['a']
                                    if 'd' in bib_record[0]['100']:
                                        xml_author += " " + bib_record[0]['100']['d']
                                    elif 'e' in bib_record[0]['100']:
                                        xml_author += " " + bib_record[0]['100']['e']
                                elif '700' in bib_record[0]:
                                    if 'a' in bib_record[0]['700']:
                                        xml_author = bib_record[0]['700']['a']
                                    if 'd' in bib_record[0]['700']:
                                        xml_author += " " + bib_record[0]['700']['d']
                                    elif 'e' in bib_record[0]['700']:
                                        xml_author += " " + bib_record[0]['700']['e']
                                xml_author = xml_author.lower()
                                xml_author = re.sub(r',$', '', xml_author)


                                xml_year = ""

                                if '260' in bib_record[0]:
                                    if 'c' in bib_record[0]['260']:
                                        xml_year = bib_record[0]['260']['c']
                                        xml_year = re.sub(r'\D', '', xml_year)
                                elif '264' in bib_record[0]:
                                    if 'c' in bib_record[0]['264']:
                                        xml_year = bib_record[0]['264']['c']
                                        xml_year = re.sub(r'\D', '', xml_year)

                                collection = ""
                                type = ""
                                if 'AVE' in bib_record[0]:
                                    if 'm' in bib_record[0]['AVE']:
                                        collection = bib_record[0]['AVE']['m']
                                        type = 'temp'


                                elif '655' in bib_record[0] and "Electronic books".lower() in bib_record[0]['655']['a'].lower():
                                    type = 'ebook'
                                print("\n\n\nXML:        " + str(xml_title) + "|" + str(xml_author) + "|" + str(xml_year))
                                print(      "Analytics:  " + str(title) + "|" + str(author) + "|" + str(year) + "|" + type + "\n\n\n")
                                test_file.write("\n\n\nXML:        " + str(xml_title) + "|" + str(xml_author) + "|" + str(xml_year) + "\n")
                                test_file.write(      "Analytics:  " + str(title) + "|" + str(author) + "|" + str(year) + "|" + type + "\n\n\n")

                                if ('AVE' in bib_record[0] and str(title) == str(xml_title) and xml_author in author and str(year) == str(xml_year) and str(mms_id) == str(xml_mms_id)):
                                    ebook_match_on_list_counter += 1
                                    temporary_collections_portfolio_counter_on_course += 1
                                    match = True
                                    break
                                elif ('655' in bib_record[0] and "Electronic books".lower() in bib_record[0]['655']['a'].lower() and str(title) == str(xml_title) and xml_author in author and str(year) == str(xml_year)):
                                    ebooks_to_add = ebooks_to_add.append({'MMS ID': xml_mms_id, 'Title': xml_title, 'Author': xml_author, 'Publication Year': xml_year, 'URL or Collection': xml_url}, ignore_index=True)
                                    ebook_for_physical_counter += 1
                                    match = True
                                    break
                                elif ('655' in bib_record[0] and "Electronic books".lower() in bib_record[0]['655']['a'].lower() and str(title) == str(xml_title) and xml_author in author and quasi_match_bool == False):
                                    ebooks_to_add_different_year = ebooks_to_add_different_year.append({'MMS ID': xml_mms_id, 'Title': xml_title, 'Author': xml_author, 'Publication Year': xml_year, 'URL or Collection': xml_url}, ignore_index=True)
                                    different_year_ebook_for_physical_counter += 1
                                    match = True
                                    break
                                elif ('AVE' in bib_record[0] and str(title) == str(xml_title) and xml_author in author and str(year) == str(xml_year)):
                                    # print("got into AVE")
                                    ebooks_to_add = ebooks_to_add.append({'MMS ID': xml_mms_id, 'Title': xml_title, 'Author': xml_author, 'Publication Year': xml_year}, ignore_index=True)
                                    temporary_collections_portfolio_counter += 1
                                    ebook_for_physical_counter += 1
                                    #ebook_match_on_list_counter += 1


                                    covid_e_books_df = covid_e_books_df.append(course_df.iloc[y], ignore_index=True)
                                    covid_e_books_df.iloc[len(covid_e_books_df) - 1]['URL or Collection'] = collection
                                    match = True
                                    break

                                elif ('AVE' in bib_record[0] and title == xml_title and xml_author in author and quasi_match_bool == False):
                                    temporary_collections_portfolio_counter_near_match += 1
                                    ebooks_to_add_different_year = ebooks_to_add_different_year.append({'MMS ID': xml_mms_id, 'Title': xml_title, 'Author': xml_author, 'Publication Year': xml_year}, ignore_index=True)
                                    covid_e_books_near_match_df = covid_e_books_near_match_df.append(course_df.iloc[y], ignore_index=True)
                                    different_year_ebook_for_physical_counter += 1
                                    covid_e_books_near_match_df = covid_e_books_near_match_df.append(course_df.iloc[y], ignore_index=True)
                                    covid_e_books_near_match_df.iloc[len(covid_e_books_near_match_df) - 1]['URL or Collection'] = collection
                                    match=True
                                    break

                                z += 1
                if match == False:
                    ebooks_we_need = ebooks_we_need.append(course_df.iloc[y], ignore_index=True)
                    no_match_counter += 1
                #print(ebooks_on_this_course)
                y += 1
        print("Course ebook count:               " + str(course_ebook_count))
        print("Different:                        " + str(different_year_ebook_for_physical_counter))
        print("Temporary:                        " + str(temporary_collections_portfolio_counter_near_match))
        print("Ebook for physical:               " + str(ebook_for_physical_counter))


        # print("Number of books on course:                                                " + str(number_of_books_on_course))
        # print("Number of books for which there's an electronic copy on reading list:     " + str(ebook_match_on_list counter))
        # print("     Ebooks on list under temporary COVID Collection")
        #
        course_name = course_df_list[x].iloc[0]['Course Name']
        course_code = course_df_list[x].iloc[0]['Course Code']
        processing_dept = course_df_list[x].iloc[0]['Processing Department']


        #counts_file = open(oDir + '/Counts - ' + str(course_name) + ' - ' + date + '.txt', 'w+')

        counts_df = counts_df.append({'Processing Department': processing_dept, 'Course Name': course_name, 'Course Code': course_code, 'Books on Course': number_of_books_on_course, 'No Ebook in Collection': no_match_counter, 'Physical Books on Course': number_of_physical_books_on_course, 'Electronic Copies for Physical on Course': ebook_match_on_list_counter, 'COVID Temporary Electronic Copies for Physical on Course - Subset': temporary_collections_portfolio_counter_on_course,  'Electronic Copies for Physical on Course - Different Year': ebook_match_on_list_counter_without_year, 'Ebook Copy for Physical in Repo - Add': ebook_for_physical_counter, 'Ebook Copy for Physical in Repo - Potentially Add - Different Year': different_year_ebook_for_physical_counter}, ignore_index=True)




        # processing_dept_master = proc_dept_df_list[d].iloc[0]['Processing Department']
        #

        # date = datetime.datetime.now().strftime("%m-%d-%Y")
        # filename = oDir + '/Electronic Copy Counts for - ' + processing_dept + ' - ' + date + '.xlsx'
        #
        # book = load_workbook(filename)
        # writer = pandas.ExcelWriter(filename, engine='xlsxwriter')
        # writer.book = book
        # counts_sheet writer.sheets['Sheet'] = dict((ws.title, ws) for ws in book.worksheets



        # counts_sheet = sheet_1
        #
        # in_repo = workbook.add_worksheet('InRepo')
        # in_repo_different_year = workbook.add_worksheet('InRepoDifferentYear')
        #
        # temporary_collection_sheet = workbook.add_worksheet('TempColl')
        # temporary_collection_sheet = workbook.add_worksheet('TempCollDiffYear')
        #
        # to_order_sheet = workbook.add_worksheet('ForPurchase')



        # counts_df.drop(['Processing Department', 'Ebook Copy for Physical in Repo - Potentially Add - Different Year'])
        if d == 0:
            counts_df.to_excel(writer, sheet_name='Counts', startrow=0, index=False)
            ebooks_to_add.to_excel(writer, sheet_name='InRepo',  startrow=0, index=False)
            ebooks_to_add_different_year.to_excel(writer, sheet_name='InRepoDifferentYear', startrow=0, index=False)
            covid_e_books_df.to_excel(writer, sheet_name='TempColl', startrow=0, index=False)
            covid_e_books_near_match_df.to_excel(writer, sheet_name='TempCollDiffYear', startrow=0, index=False)
            ebooks_to_add.to_excel(writer, sheet_name='ForPurchase', startrow=0, index=False)

        else:

            counts_max_row = workbook.sheet['Counts'].dim_rowmax
            in_repo_max_row = workbook.sheet['InRepo'].dim_rowmax
            in_repo_diff_max_row = workbook.sheet['InRepoDifferentYear'].dim_rowmax
            temp_coll_max_row = workbook.sheet['TempColl'].dim_rowmax
            temp_coll_diff_year_max_row = workbook.sheet['InRepoDifferentYear'].dim_rowmax
            to_order_max_row = workbook.sheet['ForPurchase'].dim_rowmax


            counts_df.to_excel(writer, sheet_name='Counts', startrow=counts_max_row, index=False, header=False)
            ebooks_to_add.to_excel(writer, sheet_name='InRepo', startrow=in_repo_max_row, index=False, header=False)
            ebooks_to_add_different_year.to_excel(writer, sheet_name='InRepoDifferentYear', startrow=in_repo_diff_max_row, index=False, header=False)
            covid_e_books_df.to_excel(writer, sheet_name='TempColl', startrow=temp_coll_max_row, index=False, header=False)
            covid_e_books_near_match_df.to_excel(writer, sheet_name='TempCollDiffYear', startrow=temp_coll_diff_year_max_row, index=False, header=False)
            ebooks_to_add.to_excel(writer, sheet_name='ForPurchase', startrow=to_order_max_row, index=False, header=False)



        writer.save()

        x += 1


    # print(proc_dept_df_list)
    # print(proc_dept_df_list[d])

    d += 1
    # print("All Ebooks: " + str(all_ebooks))
    # print("Total ebook count: " + str(ebook_counter) + "\n")
    #sru_url_prefix = "https://tufts.alma.exlibrisgroup.com/view/sru/01TUN_INST?version=1.2&operation=searchRetrieve&recordSchema=marcxml&query=alma.title="


# x = 0
# print("Starting reading through RESERVES file\n")
# # for event, elem in et.iterparse(covid_filename):
# while x < len(reserves_df.index):
#     if x > 1:
#         break
#
#     title = reserves_df.iloc[x]['title_only']
#
#     title = re.sub(r'\s', '%20', title)
#     title = '%22' + title + '%22'
#     record = requests.get(sru_url_prefix + title)
#     print(title)
#     print(record.text)
#     x += 1

    # tree = et.ElementTree(et.fromstring(covid_xml))
    # root = tree.get_root()
    #
    # for record in root.findall('record'):

    #bib_record = pym.parse_xml_to_array(io.StringIO(elem))










# reserves_df.to_excel(oDir + '/Titles from Reserves List.xlsx', index=False)
# electronic_df.to_excel(oDir + 'Titles from Electronic List.xlsx', index=False)
# covid_df.to_excel(oDir + '/Parsed COVID Titles - Sample.xlsx', index=False)
#
#
end = datetime.datetime.now() - start
test_file.close()

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
