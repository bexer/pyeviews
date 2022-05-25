"""pyeviews: Use EViews directly from python."""
import fnmatch
import gc
import re
import logging
from pkg_resources import get_distribution
from comtypes.client import CreateObject
from comtypes import automation
import numpy as np
import pandas as pa

_dist = get_distribution('pyeviews')
__version__ = _dist.version

# default app if users don't want to specify their own
globalevapp = None

def _BuildFromPython(objlen, newwf=True):
    """Construct EViews command for undated wf"""
    result = "create "  if newwf else "pagecreate "
    result = result + "u " + str(objlen)
    return result

def _BuildFromPandas(obj, newwf=True):
    """Construct EViews command for dated wf"""
    elem_cnt = len(obj)
    # parse the frequency string
    freq_str_all = obj.freqstr  # dated
    if freq_str_all == None:
        freq_str_all = pa.infer_freq(obj)
    # check frequency string for custom spacing
    spacing = None
    search_obj = re.search(r'(\d*)(.*)', freq_str_all)
    if search_obj:
        spacing = search_obj.group(1)
        freq_str_all = search_obj.group(2)
    pos = freq_str_all.find("-")
    freq_str = freq_str_all
    freq_str_sp = ' '
    if pos > 0:
        freq_str = freq_str_all[:pos]
        freq_str_sp = '(' + freq_str_all[pos+1:] + ') '
    # - if necessary, convert index to be start of period since 
    # EViews uses the beginning of the period and 
    # we'll have misaligned dates in EViews otherwise
    # - need to separate business freqs from regular otherwise get 
    # misaligned subtractions/additions of dates
    # - note that the DateOffset calculations (sometimes?) give performance warnings
    # fix with resample?
    if freq_str == 'A':
        dtr = pa.DatetimeIndex(obj + pa.DateOffset(years=-1, days=1), 
                               freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    elif freq_str == 'BA':
        dtr = pa.DatetimeIndex(obj - pa.tseries.offsets.BYearEnd() 
                               + pa.tseries.offsets.BDay(), freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    elif freq_str == 'Q':
        dtr = pa.DatetimeIndex(obj - pa.tseries.offsets.QuarterEnd() 
                               + pa.DateOffset(days=1), freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    elif freq_str == 'BQ':
        dtr = pa.DatetimeIndex(obj - pa.tseries.offsets.BQuarterEnd() 
                               + pa.tseries.offsets.BDay(), freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    elif freq_str == 'M':
        dtr = pa.DatetimeIndex(obj - pa.tseries.offsets.MonthEnd() 
                               + pa.DateOffset(days=1), freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    elif freq_str == 'BM':
        dtr = pa.DatetimeIndex(obj - pa.tseries.offsets.BMonthEnd() 
                               + pa.tseries.offsets.BDay(), freq = "infer")
        return _BuildFromPandas(dtr, newwf)
    # first part of the EViews command
    result = "create " if newwf else "pagecreate "
    # construct the time period strings
    yr_begin = str(obj[0].year)
    yr_end = str(obj[elem_cnt - 1].year)
    mo_begin = str(obj[0].month)
    mo_end = str(obj[elem_cnt - 1].month)
    day_begin = str(obj[0].day)
    day_end = str(obj[elem_cnt - 1].day)
    date_begin = mo_begin + '/' + day_begin + '/' + yr_begin + ' '
    date_end = mo_end + '/' + day_end + '/' + yr_end + ' '
    # may need some extra processing for custom business days
    missingbdays = _MissingElements(list(set(obj.dayofweek)))
    # +1 for first custom day in sequence, +1 bc EV d.o.w. has Mon = 1
    if missingbdays:
        dow_begin = str(missingbdays[-1] + 1 + 1) 
        dow_end = str(missingbdays[0])
    else:
        dow_begin = str(obj.dayofweek.min() + 1)
        dow_end = str(obj.dayofweek.max() + 1)
    time_begin = str(obj[0].strftime('%H')) + ':' + \
                 str(obj[0].strftime('%M')) + ':' + \
                 str(obj[0].strftime('%S')) + ' '
    time_end = str(obj[elem_cnt - 1].strftime('%H')) + ':' + \
               str(obj[elem_cnt - 1].strftime('%M')) + ':' + \
               str(obj[elem_cnt - 1].strftime('%S')) + ' '
    time_min = str(obj.hour.min()) + ':' + \
               str(obj.minute.min()) + ':' + \
               str(obj.second.min())
    time_max = str(obj.hour.max()) + ':' + \
               str(obj.minute.max()) + ':' + \
               str(obj.second.max())    
    # yearly
    if (freq_str in ['AS', 'A', 'BAS', 'BA'] and \
        spacing in ['2', '3', '4', '5', '6', '7', '8', '9', '10', '20']):
            # month alignment not allowed for multi-year freqs in EViews
            result = result + spacing + 'a ' + date_begin + date_end
    elif (freq_str in ['AS', 'A', 'BAS', 'BA'] and not spacing):
        result = result + 'a' + freq_str_sp + date_begin + date_end
    # quarterly
    elif (freq_str in ['QS', 'Q', 'BQS', 'BQ'] and not spacing):
        result = result + 'q' + freq_str_sp + date_begin + date_end
    # monthly
    elif (freq_str in ['MS', 'M', 'BMS', 'BM', 'CBMS', 'CBM'] and not spacing): 
        result = result + 'm' + freq_str_sp  + date_begin + date_end
    # weekly
    elif (freq_str == 'W' and not spacing): 
        result = result + 'w' + freq_str_sp + date_begin + date_end
    # daily
    elif (freq_str == 'D' and not spacing):
        result = result + 'd7' + freq_str_sp + date_begin + date_end
    # daily (business days)
    elif (freq_str == 'B' and not spacing):
        result = result + 'd5' + freq_str_sp + date_begin + date_end
    # daily (custom days)
    elif (freq_str == 'C' and not spacing):
        result = result + 'd(' + dow_begin + ',' + dow_end + ') ' + \
                 freq_str_sp + date_begin + date_end    
    # hourly
    elif (freq_str in ['H', 'BH'] and (not spacing or \
        spacing in ['2', '4', '6', '8', '12'])):
            result = result + spacing + 'h(' + dow_begin + '-' + dow_end + \
                     ', ' + time_min + ' - ' + time_max + ') ' + \
                     date_begin + time_begin + date_end + time_end
            if spacing:
                logging.warning("Hourly pandas DatetimeIndex may not be exactly \
                                replicated in EViews.")
    # minutes
    elif (freq_str in ['T', 'min'] and (not spacing or \
        spacing in ['2', '5', '10', '15', '20', '30'])):
            result = result + spacing + 'Min(' + dow_begin + '-' + dow_end + \
                    ', ' + time_min + ' - ' + time_max + ') ' + \
                     date_begin + time_begin + date_end + time_end
    # seconds
    elif (freq_str == 'S' and (not spacing or spacing in ['5', '15', '30'])):
            result = result + spacing + 'Sec(' + dow_begin + '-' + dow_end + \
                    ', ' + time_min + ' - ' + time_max + ') ' + \
                     date_begin + time_begin + date_end + time_end
    # EViews doesn't support these frequencies
    elif (freq_str in ['L', 'ms', 'U', 'us', 'N']):
        raise ValueError('Frequency ' + freq_str
                         + ' is not supported in EViews.')
    else:
        raise ValueError('Frequency ' + spacing
                         + freq_str + ' is not supported in EViews.')
    return result   
    
def _CheckReservedNames(names):
    """Check for disallowed names in EViews."""
    if 'c' in names:
        raise ValueError('c is not an allowed name for a series.')
    if 'resid' in names:
        raise ValueError('resid is not an allowed name for a series.')

def _GetApp(app=None):
    """Return app (user-defined or default)"""
    global globalevapp
    if app is not None:
        return app
    if globalevapp is None:
        globalevapp = GetEViewsApp(instance='new')
    return globalevapp
    
def Cleanup():
    """Clean up"""
    global globalevapp
    if globalevapp is not None:
        globalevapp = None
    gc.collect()

def PutPythonAsWF(obj, app=None, newwf=True):
    """Push Python data to EViews."""
    app = _GetApp(app)
    # which python data structure is obj?
    if (isinstance(obj, pa.core.frame.DataFrame) and not isinstance(obj.index, pa.MultiIndex)):
        #create a new EViews workfile with the right frequency
        if obj.index.inferred_type == "datetime64": # dated
            create = _BuildFromPandas(obj.index, newwf)  
        else:                                       # undated
            create = _BuildFromPython(len(obj.index), newwf)
        app.Run(create)
        _CheckReservedNames(obj.columns)
        # loop through all columns, push each into EViews as a list
        for col in obj.columns:
            col_data = automation.tagVARIANT(obj[col].tolist())
            app.PutSeries(col, col_data)
            # check if df has attributes, and if so, copy into each series
            if obj.attrs:   
                for key, value in obj.attrs.items():    
                    app.Run(str(col) + '.setattr(' + str(key) + ') ' + str(value))
    elif (isinstance(obj, pa.core.series.Series) and not isinstance(obj.index, pa.MultiIndex)):
        #create a new EViews workfile with the right frequency
        if obj.index.inferred_type == "datetime64": # dated
            create = _BuildFromPandas(obj.index, newwf)
        else:                                       # undated    
            create = _BuildFromPython(len(obj.index), newwf)
        app.Run(create)
        # push the data into EViews as a list
        name = "series"
        if obj.name:
            name = obj.name
            _CheckReservedNames([name])
        data = automation.tagVARIANT(obj.tolist())
        app.PutSeries(name, data)
        if obj.attrs:
            for key, value in obj.attrs.items():
                app.Run(str(name) + '.setattr(' + str(key) + ') ' + str(value))
    elif isinstance(obj, pa.core.indexes.datetimes.DatetimeIndex):
        #create a new EViews workfile with the right frequency
        create = _BuildFromPandas(obj, newwf)
        app.Run(create)
    elif isinstance(obj, pa.DataFrame) and isinstance(obj.index, pa.MultiIndex):
        collevels = obj.columns.nlevels
        idxlevels = obj.index.nlevels
        # only one level of cols and two levels of rows allowed
        if idxlevels != 2:
            raise ValueError("Only two index (row) levels are allowed. You have " + idxlevels + ".")
        if collevels > 1:
            #raise Warning("Multiple columns levels found. Collapsing ...")
            #newcollabels = obj.columns.map(lambda x: '_'.join([str(i) for i in x]))
            raise ValueError("Only one column level is allowed. You have " + collevels + ".")
        # create a new EViews workfile with the right frequency
        if obj.index.get_level_values(1).inferred_type == "datetime64":                # dated
            create = _BuildFromPandas(obj.index.get_level_values(1), newwf)            # total rows
        else:                                                                          # undated    
            create = _BuildFromPython(len(obj.index.get_level_values(1)), newwf)       # total rows
        app.Run(create)
        # concatenate items into single dataframe
        result = pa.concat([obj.loc[item] for item in obj.index.get_level_values(0).unique()])
        result.columns = result.columns.str.replace(" ", "_")
        _CheckReservedNames(result.columns)
        # loop through and push each column into EViews as a list
        for col in result.columns:
            col_data = automation.tagVARIANT(result[col].tolist())
            app.PutSeries(col, col_data)  
        # handle index names
        if obj.index.names:
            groupname = obj.index.names[0]
            cellname = obj.index.names[1]
        else:
            groupname = "groupid"
            cellname = "cellid"
        # structure the workfile
        app.PutSeries(groupname, automation.tagVARIANT(obj.index.get_level_values(0).tolist()))
        app.PutSeries(cellname, automation.tagVARIANT(obj.index.get_level_values(1).tolist()))
        app.Run('pagestruct(bal=m) ' + groupname + ' ' + cellname)
    elif isinstance(obj, pa.core.indexes.range.RangeIndex):
        length = len(obj)
        create = _BuildFromPython(length, newwf)
        app.Run(create) # non-standard start/stop/step information currently being lost
    elif isinstance(obj, list):
        # create a new undated workfile with the right length
        length = len(obj)
        create = _BuildFromPython(length, newwf)
        app.Run(create)
        # push the data into EViews as a list
        data = automation.tagVARIANT(obj)
        app.PutSeries("series", data)
    elif isinstance(obj, dict):
        # create a new undated workfile with the right length
        maxlength = 0
        for item in obj.values():
            if type(item) == int or type(item) == float:
                length = 1
            else:
                length = len(item)
            if length > maxlength:
                maxlength = length
        create = _BuildFromPython(maxlength, newwf)
        app.Run(create)
        _CheckReservedNames(obj.keys())
        # loop through the dict and push the data into EViews as a list
        for key in obj:
            data = automation.tagVARIANT(obj[key])
            app.PutSeries(str(key), data)
    elif isinstance(obj, np.ndarray):
        # create a new undated workfile with the right length
        create = _BuildFromPython(obj.shape[0], newwf)
        app.Run(create)
        # loop through the array, push the data into EViews as a list
        # note that nested arrays may not be transferred properly
        # is it a structured array?
        if obj.dtype.names:
            for name in obj.dtype.names:
                data = automation.tagVARIANT(obj[name].tolist())
                app.PutSeries(name, data)
        else:
            for col_num in range(obj.shape[1]):
                name = "series" + str(col_num)
                data = automation.tagVARIANT(obj[:, col_num].tolist())
                app.PutSeries(name, data)        
    else:
        raise ValueError('Unsupported type: ' + str(type(obj)))

def GetWFAsPython(app=None, wfname='', pagename='', namefilter='*'):
    """Move EViews data to Python."""
    app = _GetApp(app)
    # EViews : pandas
    dt_map = {'D5':'B', '5':'B', 'D7':'D', '7':'D', 'D':'D',
              'W':'W', 'T':'10D', 'F':'2W', 'M':'MS', 'Q':'QS',
              'S':'6M', 'A':'AS', 'Y':'AS',
              'H':'H', 'Min':'T', 'Sec':'S'} # also 'D7':'D', 'Min':'min'
    # load the workfile 
    if wfname != '':
        app.Run("wfuse " + '"' + wfname + '"') # needs full pathname
    if pagename:
        #change workfile page to specified page name
        if app.Get('=@pageexist("' + str(pagename) + '")') == 1:
            app.Run("pageselect " + str(pagename))
        else:
            raise ValueError('Invalid pagename: ' + str(pagename))
    # get workfile frequency
    pgfreq = app.Get("=@pagefreq")
    # reset sample range to all
    app.Run("smpl @all")
    # is the workfile a panel?
    ispanel = app.Get("=@ispanel")
    # get series names as a 1-dim array
    snames = app.Lookup(namefilter, "series", 1)
    anames = app.Lookup(namefilter, "alpha", 1)
    names = snames + anames
    wflength = app.Get("=@obsrange")
    #if not snames:
    #    raise ValueError('No series objects found.')
    # build *Index object
    search_obj = re.search(r'(\d*)(\w*)', pgfreq)
    if search_obj:
        spacing = search_obj.group(1)
        freq_str = search_obj.group(2)
    if (freq_str in ['F', 'M', 'Q', 'S', 'A', 'Y']):
        #get index series
        dates = app.GetSeries("@date")
        idx = pa.DatetimeIndex(dates, freq=spacing + dt_map[freq_str])
    elif freq_str == 'W': # might be improved w/ w(*) parsing, ok for now
        #get index series
        dates = app.GetSeries("@date")
        idx = pa.DatetimeIndex(dates)
    elif freq_str in ['H', 'Min', 'Sec']: 
        #get index series
        dates = app.GetSeries("@date")
        idx = pa.DatetimeIndex(dates)
        #idx.freq = pa.tseries.frequencies.to_offset(spacing + 'H')
        if spacing and fnmatch.fnmatch(pgfreq, '*H(*)'):
            logging.warning("Custom hourly frequencies in EViews may not be exactly \
replicated in pandas.")
        if spacing and fnmatch.fnmatch(pgfreq, '*Min(*)'):
            logging.warning("Custom minute frequencies in EViews may not be exactly \
replicated in pandas.")
        if spacing and fnmatch.fnmatch(pgfreq, '*Sec(*)'):
            logging.warning("Custom seconds frequencies in EViews may not be exactly \
replicated in pandas.")           
    elif (pgfreq in ['D5', '5', 'D7', '7'] or
                    fnmatch.fnmatch(pgfreq, 'D(*)')):
        #get index series
        dates = app.GetSeries("@date")
        if fnmatch.fnmatch(pgfreq, 'D(*)'):
            idx = pa.DatetimeIndex(dates)
            # unable to set following? - freq left as None!
            #idx.freq = pa.tseries.frequencies.to_offset('C')
        else:
            idx = pa.DatetimeIndex(dates, freq=dt_map[pgfreq])
    elif pgfreq == '20Y':
        raise ValueError('Frequency ' + pgfreq + ' is not supported in pandas.')
    elif pgfreq == 'U':
        idx = pa.RangeIndex(start = 0, stop = wflength)
    else:
        raise ValueError('Unsupported workfile frequency: ' + pgfreq)
    # convert series names to a space delimited string
    names_str = ' '.join(names)
    # retrieve all series+alpha data as a single call
    if len(names) != 0: 
        grp = app.GetGroup(automation.tagVARIANT(names_str), "@all")
    else:   # wf is empty
        grp = None
    if pgfreq == 'U' and len(names) != 0:
        dfr = pa.DataFrame(index = idx, columns=list(names))
        # for each series name, extract the data from our grp array
        for sindex in range(len(names)):
            dfr[names[sindex]] = [col[sindex] for col in grp]
        data = dfr
    elif pgfreq == 'U' and len(names) == 0:
        dfr = pa.DataFrame(index = idx)
        data = dfr
    elif pgfreq != 'U' and len(names) != 0:
        # for dated workfiles create the dataframe
        # build dataframe with empty columns 
        dfr = pa.DataFrame(index=idx, columns=list(snames))
        #for each series name, extract the data from our grp array
        for sindex in range(len(snames)):
            dfr[snames[sindex]] = [col[sindex] for col in grp]
        data = dfr
    elif pgfreq != 'U' and len(names) == 0:
        dfr = pa.DataFrame(index = idx)
        data = dfr
    if ispanel:
        panelids = app.Get("=@pageids").split()
        if len(panelids) != 2:
            raise ValueError("EViews panel must have two id values, not " + len(panelids) + ".")
        data = pa.DataFrame(data = dfr.drop(panelids, axis = 'columns').to_numpy(),\
                            index = pa.MultiIndex.from_product([dfr[panelids[0]].unique(), dfr[panelids[1]].unique()],\
                            names = panelids), columns = dfr.columns.drop(panelids))
    # get all attribute names
    for var in names:
        attrnames = app.Get('=@attrnames("*", "' + str(var) + '")').split()
        if attrnames and not ispanel:
            attrvals = []
            for attr in attrnames:
                tempvar = app.Get('=@getnextname("TEMP")') 
                app.Run('string ' + str(tempvar) + ' = ' + str(var).strip() + '.@attr("' + str(attr) + '")')
                attrvals.append(app.Get(str(tempvar)))
                app.Run('delete ' + str(tempvar))
            data.attrs.update(dict(zip(attrnames, attrvals)))
    # close the workfile
    #app.Run("wfclose")   
    return data
    
def GetEViewsApp(version='EViews.Manager', instance='either', showwindow=False): 
    """Allow for more control of the app"""
    # get manager object
    # this is an optional function for greater control of the app object
    # otherwise can just use the global app object
    try:
        mgr = CreateObject(version)
    except WindowsError:
        #if mgr is None:
        raise WindowsError(version + " not found.")
    if mgr is None:
        raise WindowsError(version + " not found.")
    # dictionary for type of EViews instance
    # 0 = new EViews, 1 = new or existing, 2 = existing
    table = {'new':0, 'either':1, 'existing':2}
    # get application object
    try:
        app = mgr.GetApplication(table[instance])
    except Exception:
        raise WindowsError("Problem with " + version + " EViews installation.\
                           Verify that EViews runs and is properly licensed.")
    if app is None:
        raise WindowsError("Problem with " + version + " EViews installation.\
                           Verify that EViews runs and is properly licensed.")
    # release manager object
    mgr = None
    # show EViews window
    if showwindow:
        app.Show() #optional
    return app
    
def Run(command, app=None):
    """Send commands to EViews."""
    app = _GetApp(app)
    app.Run(command)   
    
def Get(objname, app=None):
    """Retrieve the results of EViews commands."""
    app = _GetApp(app)
    return app.Get(objname)

def _MissingElements(listy):
    """Find missing elements in a sorted integer sequence.
    From: https://stackoverflow.com/questions/16974047/efficient-way-to-find-missing-elements-in-an-integer-sequence"""
    start, end = listy[0], listy[-1]
    return sorted(set(range(start, end + 1)).difference(listy))
