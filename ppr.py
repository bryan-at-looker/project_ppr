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



def checkValues(value1,value2):
    return value1 == value2

bg_rgb = ['68BD45', '329bd6','ffe300']
bg_rgb_len = len(bg_rgb)
bg_argb = ["FF" + str(x) for x in bg_rgb]

wb = openpyxl.load_workbook(filename=BytesIO(base64.b64decode(df.attachment['data'])))
# OPEN WORKBOOK IN MEMORY BY CALLING THE ATTACHMENT>>DATA AND DECODING
ws = wb.active

rows = ws.max_row
column = ws.max_column
first_data_row = {"index":2, "cell":3} # index, cell
first_data_col = {"index":3, "cell":'D'}
last_data_row = {'index':68,'cell':69}
last_data_col = {'index':74,'cell':'BW'}

cells = ws[2]
for i in range(first_data_col['index'],column):
    cells[i].font = openpyxl.styles.Font(italic=True, bold=True)


## MANIPULATE COLUMNS

cells = ws['A']
color_counter = 0
breaker = [first_data_row['cell']]

for i in range(first_data_row['index'],rows):
    cells[i].fill = openpyxl.styles.PatternFill(start_color=bg_argb[color_counter], end_color=bg_argb[color_counter], fill_type='solid')
    try:
        if not checkValues(cells[i].value, cells[i+1].value):
            color_counter = (color_counter+1) % bg_rgb_len
            breaker.append(i+1+1) # add 1 for the next cell add 1 for index
    except:
        print('last row')

breaker.append(i+1+1) # add the last row + 1


for i in range(len(breaker)-1):
    breaker_start = "A" + str(breaker[i])
    breaker_end =   "A" + str(breaker[i+1]-1)
    rng = breaker_start+':'+breaker_end
    ws.merge_cells(rng)
    ws.cell('breaker_start').alignment = openpyxl.styles.Alignment(horizontal='center')

wb.save('output_test2.xlsx')
