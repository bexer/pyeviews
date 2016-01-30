# -*- coding: utf-8 -*-
from comtypes.client import CreateObject
import numpy as np       
import pandas as pa
#import Lib/string.y

#def ConvertToEViews(df) :
#    if type(df) == pa.core.frame.DataFrame:
#        coln = list(df.columns.values)
#        datam = df.as_matrix()
#        freq = df.index.freqstr
#        dts = df.index.values
#        vec2 = ["python.DataFrame", coln, datam, freq, dts]
#        return vec2
#    else:
#        return df    

def __BuildWFFromPython(objlen):
    result = "create u " + str(objlen)
    return result

def __BuildWFFromPandas(obj) :
    elem_cnt = len(obj)
    freq_str_all = obj.freqstr
    pos = freq_str_all.find("-")
    freq_str = freq_str_all
    freq_str_sp = ' '
    if (pos > 0) :
        freq_str = freq_str_all[:pos]
        freq_str_sp = '(' + freq_str_all[pos+1:] + ') '
    result = "create "
    # construct the time period strings
    yr_begin = str(obj[0].year)
    yr_end = str(obj[elem_cnt - 1].year)
    mo_begin = str(obj[0].month)
    mo_end = str(obj[elem_cnt - 1].month)
    day_begin = str(obj[0].day)
    day_end = str(obj[elem_cnt - 1].day)
    date_begin = mo_begin + '/' + day_begin + '/' + yr_begin + ' '
    date_end = mo_end + '/' + day_end + '/' + yr_end + ' '
    dow_begin = str(obj[0].dayofweek + 1) # EV d.o.w. has Mon = 1
    dow_end = str(obj[elem_cnt - 1].dayofweek + 1)
    time_begin = str(obj[0].strftime('%H')) + ':' + \
                 str(obj[0].strftime('%M')) + ':' + \
                 str(obj[0].strftime('%S')) + ' '
    time_end = str(obj[elem_cnt - 1].strftime('%H')) + ':' + \
               str(obj[elem_cnt - 1].strftime('%M')) + ':' + \
               str(obj[elem_cnt - 1].strftime('%S'))
    time_min = str(obj.hour.min().item()) + ':' + \
               str(obj.minute.min().item()) + ':' + \
               str(obj.second.min().item()) 
    time_max = str(obj.hour.max().item()) + ':' + \
               str(obj.minute.max().item()) + ':' + \
               str(obj.second.max().item())    
    # yearly
    if (freq_str == 'AS' or freq_str == 'A' or freq_str == 'BAS' or 
            freq_str == 'BA'):
        result = result + 'a' + freq_str_sp + date_begin + date_end
    # quarterly
    elif (freq_str == 'QS' or freq_str == 'Q' or freq_str == 'BQS' or 
            freq_str == 'BQ'):
        result = result + 'q' + freq_str_sp + date_begin + date_end
    # monthly
    elif (freq_str == 'MS' or freq_str == 'M' or freq_str == 'BMS' or 
            freq_str == 'BM' or freq_str == 'CBMS' or freq_str == 'CBM'): 
        result = result + 'm' + freq_str_sp  + date_begin + date_end
    # weekly
    elif (freq_str == 'W'): 
        result = result + 'w' + freq_str_sp + date_begin + date_end
    # daily
    elif (freq_str=='D'):
        result = result + 'd7' + freq_str_sp + date_begin + date_end
    # daily (business days)
    elif (freq_str == 'B'):
        result = result + 'd5' + freq_str_sp + date_begin + date_end
    # daily (custom days)
    elif (freq_str == 'C'):
        result = result + 'd(' + dow_begin + ',' + dow_end + ') ' + \
                 freq_str_sp + date_begin + date_end    
    # hourly
    elif (freq_str == 'H' or freq_str == 'BH'):
        result = result + 'h(' + dow_begin + '-' + dow_end + ', ' + \
                 time_min + '-' + time_max + ') ' + \
                 date_begin + time_begin + date_end + time_end
    # minutes
    elif (freq_str == 'T' or freq_str == 'min'):
        result = result + 'Min(' + dow_begin + '-' + dow_end + ', ' + \
                 time_min + '-' + time_max + ') ' + \
                 date_begin + time_begin + date_end + time_end
    # seconds
    elif (freq_str == 'S'):
        result = result + 'Sec(' + dow_begin + '-' + dow_end + ', ' + \
                 time_min + '-' + time_max + ') ' + \
                 date_begin + time_begin + date_end + time_end
    # EViews doesn't support these frequencies
    elif (freq_str == 'L' or freq_str == 'ms' or freq_str == 'U' or 
            freq_str == 'us' or freq_str == 'N'):
        raise ValueError('Frequency ' + freq_str + ' is unsupported.')
    else :
        raise ValueError('Unrecognized frequency: ' + freq_str)
    return result   
    
def __CheckReservedNames(names):
    if 'c' in names:
        raise ValueError('c is not an allowed name for a series.')
    if 'resid' in names:
        raise ValueError('resid is not an allowed name for a series.')

def __CreateEViewsWF(app, obj) :
    # which python data structure is obj?
    if type(obj) == pa.core.frame.DataFrame:
        #create a new EViews workfile with the right frequency
        create = __BuildWFFromPandas(obj.index)
        app.Run(create)
        __CheckReservedNames(obj.columns)
        # loop through all columns and push each into EViews as series objects
        for col in obj.columns:
            col_data = obj.as_matrix(columns=[col])
            app.PutSeries(col, col_data)
    elif type(obj) == pa.core.series.Series:
        #create a new EViews workfile with the right frequency
        create = __BuildWFFromPandas(obj.index)
        app.Run(create)
        # push the data into EViews as a series object
        name = "series"
        if (obj.name):
            name = obj.name
            __CheckReservedNames([name])
        data = obj.as_matrix()
        app.PutSeries(name, data)
    elif type(obj) == pa.tseries.index.DatetimeIndex:
        #create a new EViews workfile with the right frequency
        create = __BuildWFFromPandas(obj)
        app.Run(create)
    elif type(obj) == pa.core.panel.Panel:
        #create a new EViews workfile with the right frequency
        create = __BuildWFFromPandas(obj.major_axis)
        create = create + str(len(obj.items))
        app.Run(create)
        # concatenate items into single dataframe
        # loop through and push each column into EViews as a series object
        result = pa.concat([obj[item] for item in obj.items])
        __CheckReservedNames(result.columns)
        for col in result.columns:
            #col_name = obj.columns[col_index]
            #col_data = list(obj.values[col_index])
            col_data = result.as_matrix(columns=[col])
            app.PutSeries(col, col_data)  
    elif type(obj) == list:
        # create a new undated workfile with the right length
        create = __BuildWFFromPython(obj)
        app.Run(create)
        # push the data into EViews as a series object
        data = np.asarray(obj)
        app.PutSeries("series", data)
    elif type(obj) == dict:
        # create a new undated workfile with the right length
        length = max(len(item) for item in obj.values())
        create = __BuildWFFromPython(length)
        app.Run(create)
        __CheckReservedNames(obj.keys())
        # loop through the dict and push the data into EViews as series objects
        for key in obj.keys():
            data = np.asarray(obj[key])
            app.PutSeries(str(key), data)
    elif type(obj) == np.ndarray:
        # create a new undated workfile with the right length
        # is it a structured array?
        if (obj.dtype.names):
            raise ValueError('Structured arrays are not supported.')
        create = __BuildWFFromPython(obj.shape[1])
        app.Run(create)
       # loop through the array, push the data into EViews as series objects
        for col_num in range(len(obj)):
            name = "series" + str(col_num)
            data = obj[:, col_num]
            app.PutSeries(name, data)        
    else:
        raise ValueError('Unsupported type: ' + str(type(obj)))

def NewEViewsWF(obj):
    #get manager object
    mgr = CreateObject("EViews.Manager")
    #get application object
    app = mgr.GetApplication(1)
    #show EViews window
    app.Show() #optional
    __CreateEViewsWF(app, obj)
    
#def EViewsCommand(commandstr):
#    app.Run(commandstr)

#create some data
#dts = pa.date_range('20000101', periods=20, freq='A')
#dts = pa.date_range('20000203', periods=20, freq='AS')
#dts = pa.date_range('20000203', periods=20, freq='BAS')
#dts = pa.date_range('20000203', periods=20, freq='BA')
#dts = pa.date_range('20000203', periods=20, freq='Q')
#dts = pa.date_range('20000203', periods=20, freq='QS')
#dts = pa.date_range('20000203', periods=20, freq='BQS')
#dts = pa.date_range('20000101', periods=20, freq='BQ')
#dts = pa.date_range('20000203', periods=20, freq='M')
#dts = pa.date_range('20000101', periods=20, freq='MS')
#dts = pa.date_range('20000101', periods=20, freq='BM')
#dts = pa.date_range('20000101', periods=20, freq='BMS')
#dts = pa.date_range('20000101', periods=20, freq='CBMS')
#dts = pa.date_range('20160101', periods=20, freq='W-THU')
#dts = pa.date_range('20000101', periods=20, freq='D')
#dts = pa.date_range('20000101', periods=20, freq='B')
weekmask_egypt = 'Sun Mon Tue Wed Thu'
bday_egypt = pa.tseries.offsets.CustomBusinessDay(weekmask=weekmask_egypt)
dts = pa.date_range('20160101', periods=20, freq=bday_egypt) # freq 'C'
#dts = pa.date_range('20160101', periods=20, freq='H')
#dts = pa.date_range('20000101', periods=20, freq='BH')
#dts = pa.date_range('20000101', periods=20, freq='min')
#dts = pa.date_range('20000101', periods=20, freq='S')
s = pa.Series(np.random.randn(20), index=dts, name = 'egypt')
#df = pa.DataFrame(np.random.randn(20, 3), index=dts, 
#                  columns=("col1", "col2", "resid"))
#pl = pa.Panel(np.random.randn(2, 20, 3), items = ['Item1', 'Item2'], 
#              major_axis = dts, minor_axis = ("col1", "col2", "c"))
#l = [9,8,7,6,5,4,3,2,1,0]
#d = {'log': l, 'y': l*2, 'z': l*3}
#a1 = np.array([(1,2,3),(4,5,6),(7,8,9)])
#a2 = np.array([(1,2,3),(2,3,4),(7,8,9)], 
#               dtype=[('hi','i4'),('bye','i4'),('why','i4')])
#print d

#CreateEViewsWF(app, dts)
#CreateEViewsWF(app, s)
#CreateEViewsWF(app, df)
#CreateEViewsWF(app, pl)
#CreateEViewsWF(app, l)
#CreateEViewsWF(app, d)
#CreateEViewsWF(app, a1)
#CreateEViewsWF(app, a2)
"""
app.Run("wfopen c:\\files\\src.wf1")
col1 = app.GetSeries("col1")
grp_manual = app.GetGroup("col1 col2 col3")
grp_auto = app.GetGroup("grp")
"""

#release the COM objects so they can close
app = None
mgr = None

if __name__ == "__main__":
    NewEViewsWF(s)