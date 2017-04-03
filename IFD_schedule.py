#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Fri Mar  3 15:18:22 2017

@author: Chen
"""
#%% 
import os
import sys
import fnmatch
from xlrd import open_workbook
import re 
import docx2txt

#%% 
def get_path():  
    """ 
    INPUT:  Ask user for location of IFD schedules 
    OUTPUT: input_path
    
    """  
    print('Please enter full path of IFD schdules')
    
    path = raw_input('>> ')
    
    if not os.path.isdir(path):
        print('This is not a directory')
        sys.exit(0)
    return path 
    
#%%
def get_keyword():  
    """ 
    INPUT:  Ask user for keyword
    OUTPUT: keyword
    
    """  
    print('Please enter your search keyword')
    keyword = raw_input('>>') 
    return keyword 

#%%
def get_files(path):
    """
    INPUT:  file path of input folder
    OUTPUT: list of excel files within input folder 
            list of word files within input folder 
    
    """
    word_list = []
    excel_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if fnmatch.fnmatch(file, '*.docx') and not ("~") in file:
                rel_file_path = os.path.join(root, file)
                word_list.append(rel_file_path)
            elif fnmatch.fnmatch(file, '*.xlsx') and not ("~") in file:
                rel_file_path = os.path.join(root, file)
                excel_list.append(rel_file_path)

    return word_list, excel_list
#%%
def search_keyword_in_word(word_list):
    print 'we are running search_keyword_in_word'
    for word in word_list:
        text = docx2txt.process(word)
        if keyword in text:
            print word
            
#%%

def search_keyword_in_excel(excel_list):
    print 'we are running search_keyword_in_excel'
##    cols = ['Day','Date','Patient','Draw Time','Status','Location','Volume',
#    'Application','Responsible for run','Responsible for run2',
#    'Responsible for run3','responsible for product','responsible for product2',
#    'responsible for product3','notes','','tech','request not','or','or2']   
##    composite = pd.DataFrame(columns=cols)
#    print composite 
#    composite_row_count = 0 
    for excel in excel_list:
        book = open_workbook(excel)
        sheet_names = book.sheet_names()
        data_sheet_names = []
        for element in sheet_names:
            m = re.match('\d+\-\d+\-\d+',element)
            if m:
                data_sheet_names.append(element)
        patient_sample_col = 2
        for sheet_name in data_sheet_names:
            sheet = book.sheet_by_name(sheet_name)
            for row in range(sheet.nrows):
                if sheet.cell(row,patient_sample_col).value == keyword:
                    print 'You can find patient', keyword, 'in row', (row+1), 'of excel sheet', sheet_name
#                    print composite_row_count
#                    new_row = []
#                    for column in range(sheet.ncols):
#                        print sheet.cell(row,column).value
#                        new_row.append(sheet.cell(row,column).value)    
#                        print 'new_row is', new_row
#                    composite.loc[composite_row_count] = new_row
#                    print 'composite is', composite
#                    composite_row_count += 1

#    fname = keyword + 'history.csv'
#    composite.to_csv(path + '/' + fname, index=False)



#%% ----- MAIN ----- 
path = get_path()
keyword = get_keyword()
word_list, excel_list = get_files(path)
search_keyword_in_word(word_list)
search_keyword_in_excel(excel_list)

