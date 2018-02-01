# coding=utf-8
#!/usr/bin/env python3

"""
print正常输出中文
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# input
import xlrd

filename_rd = 'calendar.xls'
workbook_rd = xlrd.open_workbook(filename_rd)

try:
    sheet_rd = workbook_rd.sheet_by_name('Sheet1')
except:
    print("no sheet in %s named Sheet1" % filename_rd)

output_data = []

for week in range(21):
    for i in range(4, 10):
        for j in range(1, 8):
            cell_value = sheet_rd.cell_value(10 * week + i, j)
            if cell_value != '':
                # subject
                subject = ''.join(cell_value.split('\r'))
                # date
                date = sheet_rd.cell_value(10 * week + 3, j).split('-')
                date.append('2018')
                date = '/'.join(date)
                # time
                startTime = {
                    '第一大节': '08:00 AM',
                    '第二大节': '09:50 AM',
                    '第三大节': '02:30 PM',
                    '第四大节': '04:15 PM',
                    '第五大节': '07:30 PM',
                    '第六大节': '09:15 PM',
                }
                endTime = {
                    '第一大节': '09:25 AM',
                    '第二大节': '12:00 PM',
                    '第三大节': '03:55 PM',
                    '第四大节': '04:55 PM',
                    '第五大节': '08:55 PM',
                    '第六大节': '09:55 PM',
                }
                # format
                formated = (
                    subject,  # Subject
                    date,  # Start Date
                    startTime[sheet_rd.cell_value(
                        10 * week + i, 0)],  # Start Time
                    date,  # End Date
                    endTime[sheet_rd.cell_value(10 * week + i, 0)],  # End Time
                )
                output_data.append(formated)

# output
import csv

filename_wt = 'output.csv'
file_wt = open(filename_wt,'w+')
writer = csv.writer(file_wt)
writer.writerow(['Subject','Start Date','Start Time','End Date','End Time'])
for i in range(len(output_data)):
    writer.writerow(output_data[i])
file_wt.close()
