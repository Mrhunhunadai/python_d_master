#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlrd
import math
import xlwt
import os
import numpy

'''
-------------------------
-------------------------
计算水数据
-------------------------
-------------------------
一、注意单位
      日均值的单位，均为日值，不转换为小时。

三、ph的计算
    采用两种方法计算均可，推荐使用考虑水电离平衡的算法。

四、小数点取舍
    按照规范保留，保留后计算的误差不超过1%。

    1、水的排放标准为ph  6-9，COD100，氨氮8
'''
cod_max = 100.0
andan_max = 8.0
ph_max = 9.0
ph_min = 6.0

num_list = []
cod_list = []
cod_state_list = []
andan_list = []
andan_state_list = []
ph_list = []
ph_state_list = []
flow_list = []
flow_state_list = []


'''
    读取水类数据
'''
def read_water_data():
    wb = xlrd.open_workbook('数据B.xlsx') #open file
    sheet = wb.sheet_by_name('水数据') #table water by read
    for a in range(1,sheet.nrows):#从1开始,避开第一行表头
        cells = sheet.row_values(a)
        data = int(cells[0])
        num_list.append(data)

        if type(cells[1]) == float:
            data = float(cells[1])
        else:
            #print(a,type(cells[1]),cells[1])
            data = 0.0
        cod_list.append(data)

        data = str(cells[2])
        cod_state_list.append(data)

        if type(cells[3]) == float:
            data = float(cells[3])
        else:
            #print(a,type(cells[3]),cells[3])
            data = 0.0
        andan_list.append(data)

        data = str(cells[4])
        andan_state_list.append(data)

        if type(cells[5]) == float:
            data = float(cells[5])
        else:
            #print(a,type(cells[5]),cells[5])
            data = 0.0
        ph_list.append(data)

        data = str(cells[6])
        ph_state_list.append(data)

        if type(cells[7]) == float:
            data = float(cells[7])
        else:
            #print(a,type(cells[7]),cells[7])
            data = 0.0
        flow_list.append(data)

        data = str(cells[8])
        flow_state_list.append(data)
    return

def printList():
    for i in range(0,1440):
        print(num_list[i],cod_list[i],cod_state_list[i],andan_list[i],andan_state_list[i],ph_list[i],ph_state_list[i],flow_list[i],flow_state_list[i])


'''
初始化全局 水类小时list
'''
hour_flow = []
hour_flow_state = []
hour_cod = []
hour_cod_sum = []
hour_cod_state = []
hour_andan = []
hour_andan_sum = []
hour_andan_state = []
hour_ph = []
hour_ph_state = []


'''
#屏蔽打印函数
def printHour(hour_flow,hour_flow_state):
    for i in range(0,24):
        print(i,hour_flow[i],hour_flow_state[i])
'''

'''
N状态不小于45min的，小时流量标记为N，否则为D。整小时都是D的，全部累计，标记为D。
'''
def statics_flow():
    #计算小时流量
    for i in range(24):
        data_N_total = 0
        data_D_total = 0
        has_N = 0
        has_D = 0
        for j in range(60):
            if(flow_state_list[i*60+j]=='N'):
                data_N_total += flow_list[i*60+j]
                has_N += 1
            else:
                data_D_total += flow_list[i*60+j]
                has_D += 1
        if(has_N>0):#L/s -> t/h: 累加/60*3600/1000
            data_N_total = data_N_total*60.0/1000.0
            hour_flow.append(data_N_total)
        elif(has_D==60):#整小时都是D的，全部累计，标记为D
            data_D_total = data_D_total
            hour_flow.append(data_D_total)
        else:
            data_N_total = data_N_total*60.0/1000.0
            hour_flow.append(data_N_total)

        if(has_D>15):
            data = 'D'
        else:
            data = 'N'
        hour_flow_state.append(data)
        #print("water jour: N ",i+1, data_N_total, data_N_total/60.0*1000.0, has_N)
        #print("water jour: D ",i+1, data_D_total, data_D_total/60.0*1000.0, has_D)

    #计算cod 小时均值 排放量 数据标记
    for i in range(24):
        the_formula = 0
        data = 0
        data_count = 0
        data_state = 'N'
        data_sum = 0
        flow_sum = 0

        data_other = 0
        other_count = 0

        data_D = 0

        N = 0
        O = 0
        D = 0
        M = 0
        C = 0
        T = 0
        if(hour_flow[i]==0):#该小时流量为0
            the_formula = 1
            data_state = 'F'

        for j in range(60):
            if flow_state_list[i*60+j]=='N' and cod_state_list[i*60+j]!='B':
                data += cod_list[i*60+j]*flow_list[i*60+j]
                flow_sum += flow_list[i*60+j]
                data_count += 1

            #排除B的数据
            if cod_state_list[i*60+j]!='B':
                data_other += cod_list[i*60+j]
                other_count += 1
            #计算流量为D的数量
            if flow_state_list[i*60+j]=='D':
                data_D += 1

            if(cod_state_list[i*60+j]=='N'):
                N += 1
            elif(cod_state_list[i*60+j]=='O'):
                O += 1
            elif(cod_state_list[i*60+j]=='D'):
                D += 1
            elif(cod_state_list[i*60+j]=='M'):
                M += 1
            elif(cod_state_list[i*60+j]=='C'):
                C += 1
            elif(cod_state_list[i*60+j]=='T'):
                T += 1

        if(data_state!='F'):
            if(D>0):
                data_state = 'D'
            elif(M>0):
                data_state = 'M'
            elif(C>0):
                data_state = 'C'
            elif(T>0):
                data_state = 'T'
            elif(O>0):
                data_state = 'O'
            else:
                data_state = 'N'

        if the_formula==1 or data_D==60:#流量为0或者流量全为D
            data = data_other/other_count
            #print("COD:",i+1,"data_sum/avg_count:",data)
        else:
            data = data/(flow_sum)
            #print("COD:",i+1,"data/(flow_sum):",data)

        hour_cod.append(data)

        if (data_state=='O' or data_state=='N') and data_D!=60:
            data_sum = data*hour_flow[i]*60/data_count
        else:
            data_sum = 0

        if data>cod_max and data_state=='N':
            data_state = 'O'
        hour_cod_state.append(data_state)
        hour_cod_sum.append(data_sum/1000)
        #print("COD:",i+1,hour_cod_state[i],hour_cod[i],hour_cod_sum[i],data_count)
        #print("")

    #计算andan 小时均值 排放量 数据标记
    for i in range(24):
        the_formula = 0
        data = 0
        data_count = 0
        data_state = 'N'
        data_sum = 0
        flow_sum = 0

        data_other = 0
        other_count = 0
        data_D = 0

        N = 0
        O = 0
        D = 0
        M = 0
        C = 0
        T = 0
        if(hour_flow[i]==0):#该小时流量为0
            the_formula = 1
            data_state = 'F'
        for j in range(60):
            if flow_state_list[i*60+j]=='N' and andan_state_list[i*60+j]!='B':
                data += andan_list[i*60+j]*flow_list[i*60+j]
                flow_sum += flow_list[i*60+j]
                data_count += 1

            #排除B的数值计算均值
            if andan_state_list[i*60+j]!='B':
                data_other += andan_list[i*60+j]
                other_count += 1
            #计算流量为D的数量
            if flow_state_list[i*60+j]=='D':
                data_D += 1

            if(andan_state_list[i*60+j]=='N'):
                N += 1
            elif(andan_state_list[i*60+j]=='O'):
                O += 1
            elif(andan_state_list[i*60+j]=='D'):
                D += 1
            elif(andan_state_list[i*60+j]=='M'):
                M += 1
            elif(andan_state_list[i*60+j]=='C'):
                C += 1
            elif(andan_state_list[i*60+j]=='T'):
                T += 1

        #print(i,"N:",N,"O:",O,"D:",D,"M:",M,"C:",C,"T:",T)
        if(data_state!='F'):
            if(D>0):
                data_state = 'D'
            elif(M>0):
                data_state = 'M'
            elif(C>0):
                data_state = 'C'
            elif(T>0):
                data_state = 'T'
            elif(O>0):
                data_state = 'O'
            else:
                data_state = 'N'

        if the_formula==1 or data_D==60:#流量为0或者流量全为D
            data = data_other/other_count
            #print("ANDAN:",i+1,"data_sum/avg_count:",data)
        else:
            data = data/(flow_sum)
            #print("ANDAN:",i+1,"data/(flow_sum):",data)
        hour_andan.append(data)

        if (data_state=='O' or data_state=='N') and data_D!=60:
            data_sum = data*hour_flow[i]*60/data_count
        else:
            data_sum = 0

        if data>andan_max and data_state=='N':
            data_state = 'O'
        hour_andan_state.append(data_state)
        hour_andan_sum.append(data_sum/1000)
        
        #print("ANDAN:",i+1,hour_andan_state[i],hour_andan[i],hour_andan_sum[i],data_count)
        #print("")

    ##计算ph 小时均值 数据标记
    for i in range(24):
        the_formula = 0
        data_list = []
        data_count = 0
        data_state = 'N'
        data = 0
        data_H = 0
        data_HO = 0
        flow_sum = 0
        data_D = 0

        N = 0
        if(hour_flow[i]==0):
            the_formula = 1
            data_state = 'F'

        for j in range(60):
            if flow_state_list[i*60+j]=='N' and \
                 (ph_state_list[i*60+j]=='N' or ph_state_list[i*60+j]=='O'):
                data_count += 1
                if ph_list[i*60+j] < 7:
                    data_H += math.pow(10,-ph_list[i*60+j])*flow_list[i*60+j]
                else:
                    data_H -= math.pow(10,ph_list[i*60+j]-14)*flow_list[i*60+j]
                flow_sum += flow_list[i*60+j]

            #排除B的数值来计算中位数
            if ph_state_list[i*60+j]!='B':
                data_list.append(ph_list[i*60+j])
            #计算流量D的数量
            if flow_state_list[i*60+j]=='D':
                data_D += 1

            if(ph_state_list[i*60+j]=='N'):
                N += 1
            elif(ph_state_list[i*60+j]=='O'):
                N += 1

        if(data_state!='F'):
            if(N>=45):
                data_state='N'
            else:
                data_state='D'

        if the_formula==1 or data_D==60:#流量为0或者流量全为D
            data_list.sort()                
            data = numpy.median(data_list)
        else:
            data = data_H/flow_sum
            if data_H>0:
                data = math.log10(1/data)
            else:
                data = 14-math.log10(-1/data)
        hour_ph.append(data)
        if (data<ph_min or data>ph_max) and data_state=='N':
            data_state = 'O'
        hour_ph_state.append(data_state)
        #print(i+1,hour_ph[i],hour_ph_state[i],N)
        #print('PH:',i+1,hour_ph[i])
        
    return

'''
初始化 全局 水类日数据参数list
'''
day_flow = []
day_flow_state = []
day_cod = []
day_cod_state = []
day_cod_sum = []
day_andan = []
day_andan_state = []
day_andan_sum = []
day_ph = []
day_ph_state = []

water_day =[]

def statics_day_flow():
    #计算流量 日均值、数据标记
    data = 0
    data_count = 0
    for i in range(24):
        if(hour_flow_state[i]=='N'):
            data += hour_flow[i]
            data_count += 1
    data = data/data_count
    day_flow.append(data)
    if(data_count/24>0.75):
        data_state = 'N'
    else:
        data_state = 'D'
    day_flow_state.append(data_state)
    #print("flow_day:",day_flow,day_flow_state)

    #计算cod 日均值、数据标记和排放量
    data = 0
    flow = 0
    flow_N = 0
    data_count = 0
    flow_count = 0
    for i in range(24):
        if(hour_flow_state[i]!='N'):
            continue
        elif(hour_cod_state[i]=='N'):
            data += hour_flow[i]*hour_cod[i]
            flow += hour_flow[i]
            data_count += 1
        elif(hour_cod_state[i]=='O'):
            data += hour_flow[i]*hour_cod[i]
            flow += hour_flow[i]
            data_count += 1
        elif(hour_cod_state[i]=='F'):
            data += hour_flow[i]*hour_cod[i]
            flow += hour_flow[i]
            data_count += 1
        flow_N += hour_flow[i]
        flow_count += 1
    if(data_count/24.0>0.75):
        data_state = 'N'
    else:
        data_state = 'U'
    data = data/flow
    #print(data,day_flow,data_count)
    data_sum = data*flow_N*24/flow_count/1000.0

    day_cod.append(data)
    day_cod_state.append(data_state)
    day_cod_sum.append(data_sum)
    #print("cod_day:",day_cod[0],day_cod_sum[0],day_cod_state[0],flow_count)

    #计算andan 日均值、数据标记和排放量
    data = 0
    flow = 0
    flow_N = 0
    data_count = 0
    flow_count = 0
    for i in range(24):
        if(hour_flow_state[i]!='N'):
            continue
        elif(hour_andan_state[i]=='N'):
            data += hour_flow[i]*hour_andan[i]
            flow += hour_flow[i]
            data_count += 1
        elif(hour_andan_state[i]=='O'):
            data += hour_flow[i]*hour_andan[i]
            flow += hour_flow[i]
            data_count += 1
        elif(hour_andan_state[i]=='F'):
            data += hour_flow[i]*hour_andan[i]
            flow += hour_flow[i]
            data_count += 1
        flow_count += 1
        flow_N += hour_flow[i]
    if(data_count/24.0>0.75):
        data_state = 'N'
    else:
        data_state = 'U'
    data = data/flow
    data_sum = data*flow_N*24/flow_count/1000.0

    day_andan.append(data)
    day_andan_state.append(data_state)
    day_andan_sum.append(data_sum)
    #print("andan_day:",day_andan,day_andan_sum,day_andan_state,flow_count)

    #计算ph 日均值和排放量
    the_formula = 0
    data_list = []
    data_count = 0
    data_state = 'N'
    data = 0
    data_H = 0

    N = 0
    for i in range(24):
        if(hour_flow_state[i]=='N'):
            data_count += 1
            data_H += math.pow(10,-hour_ph[i])*hour_flow[i]

        if(hour_ph_state[i]=='N'):
            N += 1
        elif(hour_ph_state[i]=='O'):
            N += 1

    if(data_state!='F'):
        if(N/24>0.75):
            data_state='N'
        else:
            data_state='D'

    #print(data_H,day_flow[0],(data_H)/(day_flow[0]*24))
    data = (data_H)/(day_flow[0]*24)
    data = math.fabs(math.log10(data))
        
    day_ph.append(data)
    day_ph_state.append(data_state)
    #print("ph_day",day_ph[0],day_ph_state[0],N)

    water_day.append(str('日均值'))
    water_day.append(float(day_cod[0]))
    water_day.append(str(day_cod_state[0]))
    water_day.append(float(day_andan[0]))
    water_day.append(str(day_andan_state[0]))
    water_day.append(float(day_ph[0]))
    water_day.append(str(day_ph_state[0]))
    water_day.append(float(day_flow[0]*24))
    water_day.append(str(day_flow_state[0]))
    water_day.append(float(day_cod_sum[0]))
    water_day.append(float(day_andan_sum[0]))
    
    return


print("--------------READ-----------------")
read_water_data()
#printList()


print("--------------HOUR START-----------------")
statics_flow()
print("--------------HOUR END-----------------")


print("--------------DAY START-----------------")
statics_day_flow()
print("--------------DAY END-----------------")


'''
-------------------------
-------------------------
计算烟气
-------------------------
-------------------------
一、注意单位
      日均值的单位，均为日值，不转换为小时。

二、气的计算场景
      烟尘为抽取法直读数据。氧量为氧化锆方法测定。基准氧量为6%。烟道面积为20平方米。
气的排放标准为烟尘20，二氧化硫 50，氮氧化物 200

四、小数点取舍
按照规范保留，保留后计算的误差不超过1%。
'''
gas_area = 20.0
base_O2 = 6.0
a34013_max = 20.0
SO2_max = 50.0
NOX_max = 200.0

list_min = []
a34013_min = []
a34013_min_state = []

SO2_min = []
SO2_min_state = []

NOX_min = []
NOX_min_state = []

O2_min = []
O2_min_state = []

temp_min = []
press_min = []
humi_min = []
rate_min = []
state_min = []

#读取分钟数据源
def read_gas_data():
    wb = xlrd.open_workbook('数据B.xlsx') #open file
    sheet = wb.sheet_by_name('气数据') #table water by read
#    dat = [] #create null list
    cells = sheet.row_values(0)
    for a in cells:
        #print(a)
        data = str(a)
        list_min.append(data)

    for a in range(1,sheet.nrows):
        cells = sheet.row_values(a)

        if type(cells[1]) == float:
            data = float(cells[1])
        else:
            print(a,type(cells[1]),cells[1])
            data = 0.0
        a34013_min.append(data)
        #data = str(cells[2])
        #a34013_min_state.append(data)

        if type(cells[2]) == float:
            data = float(cells[2])
        else:
            print(a,type(cells[2]),cells[2])
            data = 0.0
        SO2_min.append(data)
        #data = str(cells[4])
        #SO2_min_state.append(data)

        if type(cells[3]) == float:
            data = float(cells[3])
        else:
            print(a,type(cells[3]),cells[3])
            data = 0.0
        NOX_min.append(data)
        #data = str(cells[6])
        #NOX_min_state.append(data)

        if type(cells[4]) == float:
            data = float(cells[4])
        else:
            print(a,type(cells[4]),cells[4])
            data = 0.0
        O2_min.append(data)
        #data = str(cells[8])
        #O2_min_state.append(data)

        if type(cells[5]) == float:
            data = float(cells[5])
        else:
            print(a,type(cells[5]),cells[5])
            data = 0.0
        temp_min.append(data)

        if type(cells[6]) == float:
            data = float(cells[6])
        else:
            print(a,type(cells[6]),cells[6])
            data = 0.0
        press_min.append(data)

        if type(cells[7]) == float:
            data = float(cells[7])
        else:
            print(a,type(cells[7]),cells[7])
            data = 0.0
        humi_min.append(data)

        if type(cells[8]) == float:
            data = float(cells[8])
        else:
            print(a,type(cells[8]),cells[8])
            data = 0.0
        rate_min.append(data)

        data = str(cells[9])
        state_min.append(data)
        a34013_min_state.append(data)
        SO2_min_state.append(data)
        NOX_min_state.append(data)
        O2_min_state.append(data)
    return

#打印分钟数据源
def print_min_data():
    print(list_min)
    for i in range(0,1440):
        print(i+1,a34013_min[i],SO2_min[i],NOX_min[i],O2_min[i],temp_min[i],press_min[i],humi_min[i],rate_min[i],state_min[i])
    
    return

def overload_state():
    for i in range(1440):
        if a34013_min_state[i]=='N' and a34013_min[i] > a34013_max:
            a34013_min_state[i] = 'O';
        if SO2_min_state[i]=='N' and SO2_min[i]>SO2_max:
            SO2_min_state[i] = 'O';
        if NOX_min_state[i]=='N' and NOX_min[i]>NOX_max:
            NOX_min_state[i] = 'O';
    return

#计算小时数据
state_hour = []
a34013_state_hour = []
SO2_state_hour = []
NOX_state_hour = []
def statics_hour_state(src_state):
    hour = []
    for i in range(24):
        F = 0
        D = 0
        M = 0
        C = 0
        T = 0
        O = 0
        N = 0
        data_count = 0
        data = 0
        state = 'N'
        
        for j in range(60):
            if src_state[i*60+j]=='F':
                F += 1
            elif src_state[i*60+j]=='D':
                D += 1
            elif src_state[i*60+j]=='M':
                M += 1
            elif src_state[i*60+j]=='C':
                C += 1
            elif src_state[i*60+j]=='T':
                T += 1
            elif src_state[i*60+j]=='B':
                B += 1
            elif src_state[i*60+j]=='O':
                O += 1
            elif src_state[i*60+j]=='N':
                N += 1

        if F >= 45:
            state = 'F'
            print(F)
        elif D > 15:
            state = 'D'
            print(D)
        elif M > 15:
            state = 'M'
            print(M)
        elif C > 15:
            state = 'C'
            print(C)
        elif T >= 45:
            state = 'T'
            print(T)
        elif O >= 45:
            state = 'O'
            print(O)
        elif N >= 45:
            state = 'N'
            print(N)
        elif (D+M+C+B) > 15:
            if D > 0:
                state = 'D'
            elif M > 0:
                state = 'M'
            elif C > 0:
                state = 'C'
            elif B > 0:
                state = 'B'
            print(i+1,'return:',state,'\t',',D:',D,',M:',M,',C:',C,',B:',B)
        elif (F+T+O+N) >= 45:
            if F > 0:
                state = 'F'
            elif T > 0:
                state = 'T'
            elif O > 0:
                state = 'O'
            elif N > 0:
                state = 'N'
            print(i+1,'return:',state,'\t',',F:',F,',T:',T,',O:',O,',N:',N)
        else:
            state = 'N'
            print(i+1,'return:',state,'\t',',F:',F,',D:',D,',M:',M,',C:',C,',T:',T,',O:',O,',N:',N)

        hour.append(state)
        print(i+1,'return:',hour[i],state)
        print()
    return hour

def statics_hour_data(name,src_data,src_state):
    hour = []
    for i in range(24):
        F = 0
        D = 0
        M = 0
        C = 0
        T = 0
        O = 0
        N = 0
        data_count = 0
        data = 0

        for j in range(60):
            if src_state[i*60+j]=='F' or src_state[i*60+j]=='O' or \
               src_state[i*60+j]=='N' or src_state[i*60+j]=='T':
                data += src_data[i*60+j]
                data_count += 1
        if data_count>0:
            #print(name,i+1,data/data_count, "=", data ,"/", data_count)
            #print(name,i+1,data/data_count)
            hour.append(float(data/data_count))
        else:
            #print(name,i+1,data, "=", data ,"/", data_count)
            #print(name,i+1,data)
            hour.append(float(data))
    return hour

def statics_hour_zhesuan(name, src_data):
    hour = []
    for i in range(24):
        data = src_data[i] *(21.0 - 6.0)/(21 - O2_hour[i])
        hour.append(float(data))
        #print(name,hour[i])
    
    return hour

def statics_hour_sum(name, src_data, src_state, flow):
    sum = []
    for i in range(24):
        if src_state[i]=='F' or src_state[i]=='O' or \
               src_state[i]=='N' or src_state[i]=='T':
            data = src_data[i]*flow[i]/1000000.0
        else:
            data = 0

        sum.append(float(data))
    return sum

def statics_day_avg(src_data, src_state):
    data = 0
    count = 0
    for i in range(24):
        if src_state[i]=='F' or src_state[i]=='O' or \
                src_state[i]=='N' or src_state[i]=='T':
            data += src_data[i]
            count += 1
    if count > 0:
        data = data/count
    #print(data,count,data*count)
    return data

print("##############start read gas source data###############")
read_gas_data()
#overload_state()
#print_min_data()
print("##############end read gas source data###############")
state_hour=statics_hour_state(state_min)

print()
print("a34013_state_hour")
a34013_state_hour=statics_hour_state(a34013_min_state)

print()
print("SO2_state_hour")
SO2_state_hour=statics_hour_state(SO2_min_state)

print()
print("NOX_state_hour")
NOX_state_hour=statics_hour_state(NOX_min_state)


#a34013_hour = []
a34013_hour = statics_hour_data("a34013_hour",a34013_min,a34013_min_state)
print()
#SO2_hour = []
SO2_hour = statics_hour_data("SO2_hour:",SO2_min,SO2_min_state)
print()
#NOX_hour = []
NOX_hour = statics_hour_data("NOX_hour:",NOX_min,NOX_min_state)
print()
#O2_hour = []
O2_hour = statics_hour_data("O2_hour:",O2_min,state_min)
print()
#temp_hour = []
temp_hour = statics_hour_data("temp_hour:",temp_min,state_min)
print()
#press_hour = []
press_hour = statics_hour_data("press_hour:",press_min,state_min)
print()
#humi_hour = []
humi_hour = statics_hour_data("humi_hour:",humi_min,state_min)
print()
#rate_hour = []
rate_hour = statics_hour_data("rate_hour:",rate_min,state_min)

a34013z_hour = statics_hour_zhesuan("a34013z:",a34013_hour)
SO2z_hour = statics_hour_zhesuan("SO2z:",SO2_hour)
NOXz_hour = statics_hour_zhesuan("NOXz:",NOX_hour)

def overload_state_hour():
    for i in range(24):
        if a34013_state_hour[i]=='N' and (a34013_hour[i]>a34013_max or a34013z_hour[i]>a34013_max):
            a34013_state_hour[i] = 'O'
            SO2_state_hour[i] = 'O'
            NOX_state_hour[i] = 'O'
            state_hour[i] = 'O'
        if SO2_state_hour[i]=='N' and (SO2_hour[i]>SO2_max or SO2z_hour[i]>SO2_max):
            a34013_state_hour[i] = 'O'
            SO2_state_hour[i] = 'O'
            NOX_state_hour[i] = 'O'
            state_hour[i] = 'O'
        if NOX_state_hour[i]=='N' and (NOX_hour[i]>NOX_max or NOXz_hour[i]>NOX_max):
            a34013_state_hour[i] = 'O'
            SO2_state_hour[i] = 'O'
            NOX_state_hour[i] = 'O'
            state_hour[i] = 'O'
    return
overload_state_hour()

a00000_hour = []
for i in range(24):
    data = rate_hour[i]*20.0*3600*273/(273+temp_hour[i])*(press_hour[i]*1000+101325)/101325*(1-humi_hour[i]/100.0)
    #print("标态干流量[",i,"]:",data)
    a00000_hour.append(float(data))

a34013z_sum = statics_hour_sum("a34013z:",a34013_hour, state_hour, a00000_hour)
SO2z_sum = statics_hour_sum("SO2z:",SO2_hour, state_hour, a00000_hour)
NOXz_sum = statics_hour_sum("NOXz:",NOX_hour, state_hour, a00000_hour)

day_avg = []
data = '日均值'
day_avg.append(str(data))
data = statics_day_avg(a34013_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(a34013z_hour, state_hour)
day_avg.append(float(data))
day_avg.append('N')

data = statics_day_avg(SO2_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(SO2z_hour, state_hour)
day_avg.append(float(data))
day_avg.append('N')

data = statics_day_avg(NOX_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(NOXz_hour, state_hour)
day_avg.append(float(data))
day_avg.append('N')

data = statics_day_avg(O2_hour, state_hour)
day_avg.append(float(data))
day_avg.append('N')

data = statics_day_avg(temp_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(press_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(humi_hour, state_hour)
day_avg.append(float(data))
data = statics_day_avg(rate_hour, state_hour)
day_avg.append(float(data))
day_avg.append('N')
data = statics_day_avg(a00000_hour, state_hour)
day_avg.append(float(data*24))
data = statics_day_avg(a34013z_sum, state_hour)
day_avg.append(float(data*24))
data = statics_day_avg(SO2z_sum, state_hour)
day_avg.append(float(data*24))
data = statics_day_avg(NOXz_sum, state_hour)
day_avg.append(float(data*24))
day_avg.append('N')



'''
**************************************
**************************************
**************导出表格
**************************************
**************************************
'''

title = ['小时', 'COD（mg/l）', '状态', '氨氮（mg/l）', '状态',\
         'ph均值', '状态', '流量t/h', '状态', 'COD排放量kg',\
         '氨氮排放量kg']

title1 = ['小时', '烟尘（mg/m3）', '烟尘折算浓度（mg/m3）', '状态', '二氧化硫（mg/m3）', '二氧化硫折算浓度（mg/m3）', '状态',\
         '氮氧化物（mg/m3）', '氮氧化物折算浓度（mg/m3）', '状态', '氧量（%）', '状态', '烟气温度(℃)', '烟气压力(KPa)',\
         '烟气湿度(%)', '烟气流速(m/s)', '状态','烟气流量（m3/h）', '烟尘排放量（kg/h）', '二氧化硫排放量（kg/h）', \
         '氮氧化物排放量（kg/h）', '状态']
def write_execl_xls():
    workbook = xlwt.Workbook()
    #sheet = workbook.add_sheet('water')
    sheet = workbook.add_sheet('计算结果')

    '''
    导入水数据结果
    '''
    count = 0
    for i in title:
        sheet.write(0, count, i)
        count += 1
    for i in range(24):
        sheet.write(i+1, 0, i+1)
        '''sheet.write(i+1, 1, hour_cod[i])
        sheet.write(i+1, 2, hour_cod_state[i])
        sheet.write(i+1, 3, hour_andan[i])
        sheet.write(i+1, 4, hour_andan_state[i])
        sheet.write(i+1, 5, hour_ph[i])
        sheet.write(i+1, 6, hour_ph_state[i])
        sheet.write(i+1, 7, hour_flow[i])
        sheet.write(i+1, 8, hour_flow_state[i])
        sheet.write(i+1, 9, hour_cod_sum[i])
        sheet.write(i+1, 10, hour_andan_sum[i])'''

        sheet.write(i+1, 1, round(hour_cod[i],1))
        sheet.write(i+1, 2, hour_cod_state[i])
        sheet.write(i+1, 3, round(hour_andan[i],2))
        sheet.write(i+1, 4, hour_andan_state[i])
        sheet.write(i+1, 5, round(hour_ph[i],2))
        sheet.write(i+1, 6, hour_ph_state[i])
        sheet.write(i+1, 7, round(hour_flow[i],1))
        sheet.write(i+1, 8, hour_flow_state[i])
        sheet.write(i+1, 9, round(hour_cod_sum[i],3))
        sheet.write(i+1, 10, round(hour_andan_sum[i],3))

    for i in range(len(water_day)):
        if type(water_day[i]) == float:
            sheet.write(25, i, round(water_day[i],3))
        else:
            sheet.write(25, i, water_day[i])

    '''
    导入气数据结果
    '''
    #sheet = workbook.add_sheet('gas')
    init_row = 30
    count = 0
    for i in title1:
        sheet.write(0+init_row, count, i)
        count += 1
    for i in range(24):
        sheet.write(i+1+init_row, 0, i+1)
        sheet.write(i+1+init_row, 1, round(a34013_hour[i],1))
        sheet.write(i+1+init_row, 2, round(a34013z_hour[i],1))
        sheet.write(i+1+init_row, 3, a34013_state_hour[i])

        sheet.write(i+1+init_row, 4, round(SO2_hour[i],1))
        sheet.write(i+1+init_row, 5, round(SO2z_hour[i],1))
        sheet.write(i+1+init_row, 6, SO2_state_hour[i])

        sheet.write(i+1+init_row, 7, round(NOX_hour[i],1))
        sheet.write(i+1+init_row, 8, round(NOXz_hour[i],1))
        sheet.write(i+1+init_row, 9, NOX_state_hour[i])

        sheet.write(i+1+init_row, 10, round(O2_hour[i],1))
        sheet.write(i+1+init_row, 11, state_hour[i])

        sheet.write(i+1+init_row, 12, round(temp_hour[i],1))
        sheet.write(i+1+init_row, 13, round(press_hour[i],1))
        sheet.write(i+1+init_row, 14, round(humi_hour[i],1))
        sheet.write(i+1+init_row, 15, round(rate_hour[i],1))
        sheet.write(i+1+init_row, 16, state_hour[i])

        sheet.write(i+1+init_row, 17, round(a00000_hour[i],1))
        sheet.write(i+1+init_row, 18, round(a34013z_sum[i],3))
        sheet.write(i+1+init_row, 19, round(SO2z_sum[i],3))
        sheet.write(i+1+init_row, 20, round(NOXz_sum[i],3))
        sheet.write(i+1+init_row, 21, state_hour[i])
    for i in range(len(day_avg)):
        if type(day_avg[i]) == float:
            sheet.write(25+init_row, i, round(day_avg[i],3))
        else:
            sheet.write(25+init_row, i, day_avg[i])

    
    dirPath = "./计算结果C.xlsx"
    #print('移除前test目录下有文件:',os.listdir(dirPath))
    #判断文件是否存在
    if(os.path.exists(dirPath)):
        os.remove(dirPath)
        #print('移除后test 目录下有文件：',os.listdir(dirPath))
    else:
        print("要删除的文件不存在！")
    workbook.save(dirPath)

    return
write_execl_xls()
