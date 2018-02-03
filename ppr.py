# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import numpy
import openpyxl
import pandas
import json
import base64
import smtplib,ssl
from io import BytesIO
import colorsys


json_query = '/Users/Looker/python/tenx ppr/webhook_looker_simple.json'
#with open(json_query) as data_file:    
#  d = json.load(data_file) 
df = pandas.read_json(json_query) 
df_data = pandas.read_json(df.attachment['data'])




json_query = '/Users/Looker/python/tenx ppr/webhook_looker_xlsx.json'
df = pandas.read_json(json_query) 
# df_query = pandas.DataFrame(df.scheduled_plan['query'])
#df.scheduled_plan['query']

#df.attachment['data']
#df.attachment['mimetype']

#base64.b64decode(df.attachment['data'])
#file = open("testfile.xlsx","w")
#file.write(base64.b64decode(df.attachment['data']))
#file.close()

#wb = openpyxl.load_workbook('testfile.xlsx')

# OPEN WORKBOOK IN MEMORY BY CALLING THE ATTACHMENT>>DATA AND DECODING

bg_rgb = ['68BD45', '329bd6','ffe300']
bg_color = ('FF68BD45', 'FF329bd6','FFffe300')

def colorFader(n,)

def checkValues(value1,value2):
    return value1 == value2

wb = openpyxl.load_workbook(filename=BytesIO(base64.b64decode(df.attachment['data'])))
ws = wb.active

rows = ws.max_row
column = ws.max_column

cells = ws[2]
for i in range(2,column):
    cells[i].font = openpyxl.styles.Font(italic=True, bold=True)

cells = ws['A']
color_counter = 0




breaker = [2]

for i in range(2,rows):
    cells[i].fill = openpyxl.styles.PatternFill(start_color=bg_color[color_counter], end_color=bg_color[color_counter], fill_type='solid')
    try:
        if not checkValues(cells[i].value, cells[i+1].value):
            color_counter = (color_counter+1) % 3
            breaker.append()
    except:
        print('last row')

    #print(cells[i].fill)


wb.save('output_test2.xlsx')
