DRIP-vs-Manual
==============

A python script to compare DRIP and manual investments

This script compares a DRIP investment to a manual reinvestment strategy. An article on www.seekingalpha.com was written
based on the results (http://seekingalpha.com/article/1703972-reinvesting-dividends-vs-saving-them-and-buying-on-stock-price-drops-part-ii).

Although the ticker symbols in the 'tckrlist' include:
['APD','MO','T','BKH','ED','DBD','DOV','BEN','LANC','NC','NWN','PPG','RPM','SWK','UBSI','UVV','WGL','ERIE','ESS','NNN','PRE']

only 'ED', 'DBD', 'DOV', 'PPG', 'SWK', 'UVV', 'WGL', 'NNN' are included in the analysis in the article. (The inclusion of 'MO'
in the article was a typo). Also, although July 17th, 2000 was the reported back test start date, Sept 19, 2000 was the actual
start date. This shouldn't have any noticable effect on the results.

The list positions for these tickers are:

tckrlist[1], tckrlist[4], tckrlist[5], tckrlist[6], tckrlist[11], tckrlist[13], tckrlist[15], tckrlist[16], tckrlist[19]

respectively. 

To perform a DRIP versus manual backtest on one of those eight stocks:

    m,t = manualtest([4])
    d =  driptest([4])
    print 'The manual value was', m, 'and', t, 'transactions took place'
    print 'The DRIP value was', d

the values m and d are the manual value on the final date of the back test July 15, 2013. 

The manualtest function also returns the number of transactions after the final value.

To test portfolios with more than one stock enter the list positions separated by commas. For example, to back test compare
a portfolio consisting of ED, DBD, DOV and PPG:

    m,t = manualtest([4, 5, 6, 11, 13])
    d = driptest([4, 5, 6, 11, 13])
    print 'The manual value was', m, 'and', t, 'transactions took place'
    print 'The DRIP value was', d

Dependendies include:

xlrd

divpaydatedata.xls (this file must be created by the user by combining the three .csv files into one spread sheet with
pay date, ex-date and payment values as sheets 1, 2, and 3, respectively). Save the file as divpaydatedata.xls and put in the same
directory as the .py file.
