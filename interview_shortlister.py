#importing required library

import pandas as pd
import numpy as np
import seaborn as sns
import xlrd
import csv
import matplotlib.pyplot as plt
import docx2txt
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.linear_model import LogisticRegression
from sklearn.externals import joblib




# find your sheet name at the bottom left of your excel file and assign 
# it to my_sheet 
# Program extracting first column

#read excel file having data 

loc = ("mavoix_ml_sample_dataset.xlsx") 

wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
#sheet.cell_value(0, 0)


#print(sheet.cell_value(2,14))


web_developer_skills=docx2txt.process("web_development_skills.docx")
data_science_skills=docx2txt.process("data_science_skills.docx")
data=[web_developer_skills, sheet.cell_value(0, 16)]

#defining required functions


def find_match_percentage(x):
    vectorizer = CountVectorizer()      
    count_matrix = vectorizer.fit_transform(x)
    match_percentage=cosine_similarity(count_matrix)[1][0]*100
    match_percentage=round(match_percentage,2)
    return match_percentage

def fl_value(x,y):

    try:

        val=sheet.cell_value(x,y)
        if val=='' or val=='N/A' or val==' ':
            return 0.0
        else:
            return float(int(val))
    except ValueError:
        return 0.0
        

def string_value(x):

    val=sheet.cell_value(x,16)
    return val



def skill_point(x):
    score=0.00
    for cell in range(2,15):
        #print(sheet.cell_value(x,cell))
        score=score+fl_value(x,cell)

    return score

def convert_to_float(frac_str):
    val=frac_str
    if val=='' or val=='N/A':
        return 0.0
    else:
    
        try:
            return float(frac_str)
        except ValueError:
            num, denom = frac_str.split('/')
            try:
                leading, num = num.split(' ')
                whole = float(leading)
            except ValueError:
                whole = 0
            frac = float(num) / float(denom)
            return whole - frac if whole < 0 else whole + frac



def accademic_scor(x):
    acc_score=0.0
    for cell in range(21,23):
        normalized_score=convert_to_float(sheet.cell_value(x,cell))*100
        acc_score=acc_score+normalized_score
    return acc_score


        
def web_development(x):

    data=[web_developer_skills, sheet.cell_value(x, 16)]
    match_value=find_match_percentage(data)
    match_value=match_value+skill_point(x)
    match_value=match_value+accademic_scor(x)
    return match_value


def data_scientist(x):

    data=[data_science_skills, sheet.cell_value(x, 16)]
    match_value=find_match_percentage(data)
    match_value=match_value+skill_point(x)
    match_value=match_value+accademic_scor(x)
    return match_value






# create dictionary to store results

data_key_value ={}
web_key_value ={}
def data_dictionairy(x,y):  
  
       
         
    data_key_value[x] = y

def web_dictionairy(x,y):  
       
           
    web_key_value[x] = y

    
shr = wb.sheet_by_name('Sheet1')
for rownum in range(1,shr.nrows):
    web_development_match_value=web_development(rownum)
    data_scientist_match_value=data_scientist(rownum)
    if (data_scientist_match_value>web_development_match_value):
        y=data_scientist_match_value
        x=sheet.cell_value(rownum,0)
        data_dictionairy(x,y)

    else:
        y=data_scientist_match_value
        x=sheet.cell_value(rownum,0)
        web_dictionairy(x,y)


#sorting classified dictionaries


data_key_value_sorted=sorted(data_key_value.items(), key =lambda kv:(kv[1], kv[0]))
web_key_value_sorted=sorted(web_key_value.items(), key =lambda kv:(kv[1], kv[0]))


#printing top 10 ranked applicant from sorted dictionary values

print("Top 10 data science ranked candidates in order with score:")
print("( applicant    score )")
for i in range(10):
    print(data_key_value_sorted[10-i])


print("Top 10 web development ranked candidates in order with score:")
print("( applicant    score )")
for i in range(10):
    print(web_key_value_sorted[10-i])










        

        

        
        
    
    



    
#w=data_scientist(2)
#print(sc)
    
    
    
    

    

#sc=accademic_scor(2)
#print(sc)
    



















