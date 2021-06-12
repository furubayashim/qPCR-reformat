#!/usr/bin/env python

import pandas as pd
import numpy as np
import sys

if len(sys.argv) <2:
    print("Need filename as argument")
    exit(0)

filename = sys.argv[1]
outputfilename = "output.xlsx"

df = pd.read_csv(filename, delimiter='\t')
df['row'] = df.apply(lambda x:x['Position'][0],axis=1)
df['col'] = df.apply(lambda x:int(x['Position'][1:]),axis=1)

# Choose items to export from below list
# ['File Name', 'Analysis Name', 'Plate Id', 'Position', 'Sample Name',
#  'Gene Name', 'Cq', 'Concentration', 'Call', 'Excluded', 'Sample Type',
#  'Standard', 'Cq Mean', 'Cq Error', 'Concentration Mean',
#  'Concentration Error', 'Replicate Group', 'Dye', 'Edited Call', 'Slope',
#  'EPF', 'Failure', 'Notes', 'Sample Prep Notes', 'Number', 'row', 'col']
columns_to_export = ['Concentration','Gene Name','Sample Name','Position']

# use this to reformat
row96 = list('ABCDEFGH')
column96 = np.arange(1,13,1)

# make df for each item
dfs =[]
for c in columns_to_export:
    newdf = df.pivot(values=c,index='row',columns='col')
    newdf = newdf.reindex(index=row96, columns=column96)
    dfs.append(newdf)

# export in xls
writer = pd.ExcelWriter(outputfilename,engine='xlsxwriter')
workbook=writer.book
worksheet=workbook.add_worksheet('Result')
writer.sheets['Result'] = worksheet

posi = 0
for n,d in enumerate(dfs):
    worksheet.write_string(posi, 0, columns_to_export[n])
    d.to_excel(writer,sheet_name='Result',startrow=posi+1 , startcol=0)
    posi = posi + d.shape[0]+3
writer.save()
