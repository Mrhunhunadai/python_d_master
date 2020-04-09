#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlrd
import math
import xlwt
import os

def read_water_data():
    #wb = xlrd.open_workbook('test.xlsx') #open file
    #sheet = wb.sheet_by_name('water') #table water by read
    wb = xlrd.open_workbook('数据A 20200331.xlsx') #open file
    sheet = wb.sheet_by_name('水') #table water by read
#    dat = [] #create null list
    for a in range(1,sheet.nrows):
        cells = sheet.row_values(a)
        data = int(cells[0])
        num_list.append(data)

        data = float(cells[1])
        cod_list.append(data)

        data = str(cells[2])
        cod_state_list.append(data)

        data = float(cells[3])
        andan_list.append(data)

        data = str(cells[4])
        andan_state_list.append(data)

        data = float(cells[5])
        ph_list.append(data)

        data = str(cells[6])
        ph_state_list.append(data)

        data = float(cells[7])
        flow_list.append(data)

        data = str(cells[8])
        flow_state_list.append(data)
    return


