# -*- coding: utf-8 -*-
# import pandas as pd
# import numpy as np
# import math
# from sklearn import preprocessing



# first_row = 10


# def GetLen(eachtable_firstrow):
#     each_row_value = []
#     loc_location = 2
#     while not(np.isnan(df.iat[eachtable_firstrow,loc_location])):

#         value = df.iat[eachtable_firstrow,loc_location]
#         each_row_value.append(value)
#         loc_location += 1
#     return len(each_row_value)

# def GetValue(eachtable_firstrow,end):
#     each_row_value = []
#     for i in range(2,end+2):
#         value = df.iat[eachtable_firstrow,i]
#         each_row_value.append(value)
#     return each_row_value


# def GetValueLog(eachtable_firstrow,end):
#     each_row_value = []
#     for i in range(2,end+2):
#         value = df.iat[eachtable_firstrow,i]
#         if value > 0:
#             value = math.log(value,10)
#             each_row_value.append(value)
#         elif value == 0:
#             each_row_value.append(value)
#         else:
#             value = math.log(abs(value),10)
#             each_row_value.append(-value)
#     return each_row_value


# def Catch_Name(eachtable_firstrow):
#     name = df.iat[int(eachtable_firstrow)-9,0]
#     return name


# def creatFile():
#     excel = pd.ExcelFile("./input/data.xlsx")
#     sheet_names = excel.sheet_names

#     df = pd.DataFrame()
#     df.to_excel("./output/combo_char.xlsx")
#     writer = pd.ExcelWriter("./output/combo_char.xlsx")

#     for each_sheet_name in sheet_names:
#         df.to_excel(writer, sheet_name=each_sheet_name)

#     writer.save()


# creatFile()

# writer = pd.ExcelWriter("./output/combo_char.xlsx", engine = 'xlsxwriter')

# excel = pd.ExcelFile("./input/data.xlsx")
# sheet_names = excel.sheet_names

# for each_sheet_name in sheet_names:

#     df = pd.read_excel("./input/data.xlsx", header=None,sheet_name=each_sheet_name)

#     loc_col = 0


#     df.to_excel(writer, sheet_name = '{}'.format(each_sheet_name), startcol=loc_col, index=False, header=False)

#     workbook  = writer.book
#     worksheet = writer.sheets[each_sheet_name]

    
#     for eachtable_firstrow in range(first_row, df.shape[0], 15):
#     # for eachtable_firstrow in range(first_row, 30, 15):
#        index = eachtable_firstrow
#        each_table_name=Catch_Name(eachtable_firstrow)
#        Vs_V =GetValue(eachtable_firstrow,GetLen(eachtable_firstrow))
#        fires_len = len(Vs_V)
#        index += 1

#        Is_A =GetValue(index,fires_len)

#        Is_A_normalization =preprocessing.scale(Is_A)
#        worksheet.write_row('Y{}'.format(index+1), Is_A_normalization)

#        index += 1

#        Vd_v =GetValue(index,fires_len)
#        index+= 1

#        Id_a =GetValue(index,fires_len)

#        Id_a_normalization =preprocessing.scale(Id_a)
#        worksheet.write_row('Y{}'.format(index+1), Id_a_normalization)

#        col_chart = workbook.add_chart({'type': 'line'})
#        col_chart.add_series({
#                               'categories': [each_sheet_name,eachtable_firstrow,2,eachtable_firstrow,fires_len+1],
#                               'values': [each_sheet_name,eachtable_firstrow+1,2,eachtable_firstrow+1,fires_len+1],
#                               'marker':{'type':'circle','size':7},

#                             })

#        col_chart.add_series({
#                               'categories': [each_sheet_name,eachtable_firstrow+2,2,eachtable_firstrow+2,fires_len+1],
#                               'values': [each_sheet_name,eachtable_firstrow+3,2,eachtable_firstrow+3,fires_len+1],
#                               'marker':{'type':'circle','size':7},

#                               })


#        col_chart_2 = workbook.add_chart({'type': 'line'})
#        col_chart_2.add_series({
#                             'categories': [each_sheet_name, eachtable_firstrow, 2, eachtable_firstrow, fires_len+1],
#                             'values': [each_sheet_name, eachtable_firstrow + 1, 24, eachtable_firstrow + 1, 23+fires_len],
#                             'marker': {'type': 'circle', 'size': 7},
#                             'data_labels':{'value':True}
#                             })

#        col_chart_2.add_series({
#                               'categories': [each_sheet_name, eachtable_firstrow + 2, 2, eachtable_firstrow + 2, fires_len+1],
#                               'values': [each_sheet_name, eachtable_firstrow + 3, 24, eachtable_firstrow + 3, 23+fires_len],
#                               'marker': {'type': 'circle', 'size': 7},

#                               })
#        Vsd = Vs_V + Vd_v
#        Isd = Is_A + Id_a

#        Isd_normalization = Is_A_normalization + Id_a_normalization

#        col_chart.set_x_axis({
#            'name':'Index',
#            'min':min(Vsd),
#            'max':max(Vsd),
#        })
#        col_chart.set_y_axis({
#            'name':'Value',
#            'major_gridlines':{'visible':False},
#            'min':min(Isd),
#            'max':max(Isd),
#        })
#        col_chart.set_legend({'position':'none'})

#        col_chart_2.set_x_axis({
#            'name':'Index',
#            'min':min(Vsd),
#            'max':max(Vsd),
#        })
#        col_chart_2.set_y_axis({
#            'name':'Value',
#            'major_gridlines':{'visible':False},
#            'min':min(Isd_normalization),
#            'max':max(Isd_normalization),
#        })
#        col_chart_2.set_legend({'position':'none'})

#        worksheet.insert_chart('O{}'.format(eachtable_firstrow-7), col_chart)
#        worksheet.insert_chart('W{}'.format(eachtable_firstrow-7), col_chart_2)

# writer.close()

import pandas as pd
import numpy as np
import xlsxwriter
import math
from sklearn import preprocessing
import string
import re
import xlrd

first_row = 10

def GetLen(eachtable_firstrow):
    each_row_value = []
    loc_location = 2
    while not(np.isnan(float(df.iat[eachtable_firstrow,loc_location]))):
        value = df.iat[eachtable_firstrow,loc_location]
        each_row_value.append(float(value))
        loc_location += 1
        if loc_location == df.shape[1]:
            break
    return len(each_row_value)

def GetValue(eachtable_firstrow,end):
    each_row_value = []
    for i in range(2,end+2):
        value = df.iat[eachtable_firstrow,i]
        if type(value).__name__ == 'str': #處理科學符號
            if 'm' in value:
                value = value.replace('m', '')
                each_row_value.append(float(value)*float(0.001))
            elif 'u' in value:
                value = value.replace('u', '')
                each_row_value.append(float(value)*float(0.000001))
            elif 'n' in value:
                value = value.replace('n', '')
                each_row_value.append(float(value)*0.000000001)
            elif 'p' in value:
                value = value.replace('p', '')
                each_row_value.append(float(value)*0.000000000001)
            elif 'f' in value:
                value = value.replace('f', '')
                each_row_value.append(float(value)*0.000000000000001)
            elif 'a' in value:
                value = value.replace('a', '')
                each_row_value.append(float(value)*0.000000000000000001)
            else:
                each_row_value.append(float(value))
        else:
            each_row_value.append(value)
    return each_row_value

def Catch_Name(eachtable_firstrow):
    name = df.iat[int(eachtable_firstrow)-9,0]
    return name



df = pd.read_excel("./input/data.xlsx", header = None)
df = df.replace(r'[{}]'.format('?'), '', regex = True) #處理特殊字元
df = df.replace(r'[{}]'.format('*'), '', regex = True)
df = df.iloc[1:]
df = df.reset_index(drop = True)

writer = pd.ExcelWriter("./output/combo_char.xlsx", engine = 'xlsxwriter')

excel = pd.ExcelFile("./input/data.xlsx")
sheet_names = excel.sheet_names

for each_sheet_name in sheet_names:

    df = pd.read_excel("./input/data.xlsx", header=None, sheet_name=each_sheet_name)
    
    loc_col = 0

    df.to_excel(writer, sheet_name='{}'.format(each_sheet_name), startcol=loc_col, index=False, header=False)

    workbook = writer.book
    worksheet = writer.sheets[each_sheet_name]

    for eachtable_firstrow in range(first_row, df.shape[0], 15):
       index = eachtable_firstrow
       each_table_name=Catch_Name(eachtable_firstrow)
       Vs_V =GetValue(eachtable_firstrow,GetLen(eachtable_firstrow))
       fires_len = len(Vs_V)

       index += 1

       Is_A =GetValue(index,fires_len)
       # worksheet.write_row(f'Y{index}', Is_A)
       worksheet.write_row('Y{}'.format(index), Is_A)
       Is_A_normalization =preprocessing.scale(Is_A)
       worksheet.write_row('Y{}'.format(index+1), Is_A_normalization)

       index += 1

       Vd_v =GetValue(index,fires_len)
       index+= 1

       Id_a =GetValue(index,fires_len)
       worksheet.write_row('Y{}'.format(index), Id_a)

       Id_a_normalization =preprocessing.scale(Id_a)
       worksheet.write_row('Y{}'.format(index+1), Id_a_normalization)

       col_chart = workbook.add_chart({'type': 'line'})
       col_chart.add_series({
                              'categories': [each_sheet_name,eachtable_firstrow,2,eachtable_firstrow,fires_len+1],
                              'values': [each_sheet_name, eachtable_firstrow, 24, eachtable_firstrow, 23+fires_len],
                              'marker':{'type':'circle','size':7},

                            })

       col_chart.add_series({
                              'categories': [each_sheet_name,eachtable_firstrow+2,2,eachtable_firstrow+2,fires_len+1],
                              'values': [each_sheet_name, eachtable_firstrow + 2, 24, eachtable_firstrow + 2, 23+fires_len],
                              'marker':{'type':'circle','size':7},

                              })


       col_chart_2 = workbook.add_chart({'type': 'line'})
       col_chart_2.add_series({
                            'categories': [each_sheet_name, eachtable_firstrow, 2, eachtable_firstrow, fires_len+1],
                            'values': [each_sheet_name, eachtable_firstrow + 1, 24, eachtable_firstrow + 1, 23+fires_len],
                            'marker': {'type': 'circle', 'size': 7},
                            'data_labels':{'value':True}
                            })

       col_chart_2.add_series({
                              'categories': [each_sheet_name, eachtable_firstrow + 2, 2, eachtable_firstrow + 2, fires_len+1],
                              'values': [each_sheet_name, eachtable_firstrow + 3, 24, eachtable_firstrow + 3, 23+fires_len],
                              'marker': {'type': 'circle', 'size': 7},

                              })
       Vsd = Vs_V + Vd_v
       Isd = Is_A + Id_a

       Isd_normalization = Is_A_normalization + Id_a_normalization

       col_chart.set_x_axis({
           'name':'Index',
           'min':min(Vsd),
           'max':max(Vsd)
       })
       col_chart.set_y_axis({
           'name':'Value',
           'major_gridlines':{'visible':False},
           'min':min(Isd),
           'max':max(Isd),
           
       })
       col_chart.set_legend({'position':'none'})

       col_chart_2.set_x_axis({
           'name':'Index',
           'min':min(Vsd),
           'max':max(Vsd),
       })
       col_chart_2.set_y_axis({
           'name':'Value',
           'major_gridlines':{'visible':False},
           'min':min(Isd_normalization),
           'max':max(Isd_normalization),
       })
       col_chart_2.set_legend({'position':'none'})

       worksheet.insert_chart('O{}'.format(eachtable_firstrow-7), col_chart)
       worksheet.insert_chart('W{}'.format(eachtable_firstrow-7), col_chart_2)

    writer.save()
