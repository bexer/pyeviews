========
pyeviews
========

Since we love Python (who doesn’t?), we’ve had it in the back of our minds for a while now that we should find a way to make it easier for EViews and Python to talk to each other, so Python programmers could use the econometric engine of EViews directly from Python.  So we did!  Using the Python **comtypes** package we’ve written a Python package, **pyeviews**, that uses COM to transfer data between Python and EViews.  
    
Here’s a simple example going from Python to EViews.  We’re going to use the popular Chow-Lin interpolation routine in EViews using data created in Python.  Chow-Lin interpolation is a regression-based technique to transform low-frequency data (in our example, annual) into higher-frequency data (in our example, quarterly).  It has the ability to use a higher-frequency series as a pattern for the interpolated series to follow.   The quarterly interpolated series is chosen to match the annual benchmark series in one of four ways: first (the first quarter value of the interpolated series matches the annual series), last (same, but for the fourth quarter value), sum (the sum of the first through fourth quarters matches the annual series), and average (the average of the first through fourth quarters matches the annual series).
 
We’re going to create two series in Python using the time series functionality of the **pandas** package, transfer it to EViews, perform Chow-Lin interpolation on our series, and bring it back into Python.  The data are taken from [BLO2001]_ in an example originally meant for Denton interpolation.

1.	If you don’t have Python, we recommend the `Anaconda distribution <https://www.continuum.io/downloads>`_, which will include most of the packages we’ll need.  Then head over to the `Python Package Index <https://pypi.python.org/pypi>`_ and follow the directions to download and install the `**pyeviews** module <http://w3.org>`_.

2.	Create two time series using pandas.  We’ll call the annual series “benchmark” and the quarterly series “indicator”:

.. code:: python

    >>> import pandas as pa
    >>> dtsa = pa.date_range('1998', periods = 3, freq = 'A')
    >>> benchmark = pa.Series([4000.,4161.4,np.nan], index=dtsa, name = 'benchmark')
    >>> dtsq = pa.date_range('1998q1', periods = 12, freq = 'Q')
    >>> indicator = pa.Series([98.2, 100.8, 102.2, 100.8, 99., 101.6, 102.7, 101.5, 100.5, 103., 103.5, 101.5], index = dtsq, name = 'indicator')`
    
3.	Load the **pyeviews** package and call the ``NewEViewsWF`` function:

.. code:: python

    >>> import eviewspython as evp
    >>> evp.NewEViewsWF(benchmark)
    >>> evp.NewEViewsWF(indicator) <<<< need NewEViewsPage function instead!!!!!!!!!
    
4.	Use the EViews ``copy`` command to copy the benchmark series in the annual page to the quarterly page, using the indicator series in the quarterly page as the high-frequency indicator and matching the sum of the benchmarked series for each year (four quarters) with the matching annual value of the benchmark series:

.. code:: python

    >>> copy(rho=.7, c=chowlins, overwrite) annual\benchmark benchmarked @indicator indicator
    
5.	Bring the new series back into Python:


References
----------
.. [BLO2001] Bloem, A.M, Dippelsman, R.J. and Maehle, N.O. 2001 Quarterly National Accounts Manual–Concepts, Data Sources, and Compilation. IMF. http://www.imf.org/external/pubs/ft/qna/2000/Textbook/index.htm