from __future__ import division
# -*- coding: utf-8 -*-
"""
Created on Sun Jul 28 14:14:31 2013 and Tue Sep 17 19:44:12 2013

@author: ryan
"""

import xlrd
import datetime
import calendar
import random

import pylab
import scipy
from matplotlib import finance as mf
tckrlist = ['APD','MO','T','BKH','ED','DBD','DOV','BEN','LANC','NC','NWN','PPG','RPM','SWK','UBSI','UVV','WGL','ERIE','ESS','NNN','PRE']


deeptckrlist = [tckrlist[1], tckrlist[4], tckrlist[5], tckrlist[6], tckrlist[11], tckrlist[13], tckrlist[15], tckrlist[16], tckrlist[19]]

def stockpricemasterlist(tckrlist):
    stockpricedates = []
    stockpricemaster = []
    for j in tckrlist:
        temporlist = []
        temporlist2  = []
        price = mf.fetch_historical_yahoo(j, (2000, 9, 19), (2013, 7, 15), dividends=False)
        for i in price:
    
            try:
                i =i.split(',')
                i[0] = i[0].split('-')
                if i[0][1][0] == str(0):
                    
                    m = i[0][1][1]
                    
                else:
                    m = i[0][1]
                if i[0][2][0] == str(0):
                    d = i[0][2][1]
                else:
                    d = i[0][2]
                y = i[0][0]
                date = str(m)+'/'+str(d)+'/'+str(y)
                temporlist.append(date)
                temporlist2.append(i[4])
            except:
                pass
        stockpricedates.append(temporlist)
        stockpricemaster.append(temporlist2)
    return stockpricedates, stockpricemaster
    
stockdatemaster, stockpricemaster = stockpricemasterlist(tckrlist)


def removerepeats(resultlist):
    '''
    input: a list containing duplicate values
    output: a list containing each unique item from input list once, i.e. with more 
    than 1 occurance of a value
    '''

    returnlist  = []
    for i in resultlist:
        if i not in returnlist:
            returnlist.append(i)
    return returnlist

def meanstdv(x): 
    
	from math import sqrt 
	n, mean, std = len(x), 0, 0 
	for a in x: 
         a = float(a)  
         mean = mean + a 
	mean = mean / float(n) 
	for a in x:
         a = float(a)
         std = std + (a - mean)**2 
	std = sqrt(std / float(n-1)) 
	return std, mean

def pickrandomnums(i):
    '''
    picks i random numbers between 0 and 7
    once a value is chosen, it is no longer available for subsequent selection
    '''
    returnlist = []
    count = 0
    while count < i:
        a = random.randint(0, 7)
        if a not in returnlist:
            returnlist.append(a)
            count += 1
        
    return returnlist

def figure(n, s):
    '''
    calculates how many unique combinations (p) of n items there are in a set (s) of integers
    e.g. s = 1, 2, 3, 4, 5
    n = 3
    unique possibilities (p) = 10
    '''
    
    ans1 = ((s**2)-s)/2
    if n >= 2:
        count = 2
        poss = ans1
        while count < n:
            count += 1
            nextp = ((poss*s)-(count-1)*poss)/count

            poss = nextp
        return poss
    else:
        return s

def daterange():
    '''
    returns a list of all  dates in mm/dd/yyyy format, without zeros
    from sept 19, 2000 to july 15, 2013
    '''
    date = '7/15/2013' #end date of July 17, 2013
    dateobj = datetime.datetime.strptime(date, '%m/%d/%Y') 
    numdays = 4683 #days minus 7/15/2013 to reach 9/19/2000
    dateList = []
    masterdlist = []
    for x in range (0, numdays):
        dateList.append(dateobj - datetime.timedelta(days = x)) #all days between start and end
    for i in dateList:
        
         i = str(i)
         yyyy = int(i[0:4])
         mm = int(i[5:7])
         dd = int(i[8:10])
         workabledate = str(mm)+'/'+str(dd)+'/'+str(yyyy)

         masterdlist.append(workabledate)
    count = -1
    returnlist = []
    while abs(count) <= len(masterdlist):
        returnlist.append(masterdlist[count])
        count -=1
    
    return returnlist
    
fulldatelist = daterange()
#print fulldatelist[0]
  
def generate_numlists(i):
    '''
    generates a list of nonrepeating num lists, length i 
    '''
    returnlist = []
    finlist = []
    overlist = []
    while len(returnlist) != figure(i, 8):
        a = pickrandomnums(i)
        a.sort()
        if a not in returnlist:
            returnlist.append(a)

    for q in returnlist:
        finlist = []
        for j in q:
            if j == 0:
                finlist.append(4)
            if j == 1:
                finlist.append(15)
            if j == 2:
                finlist.append(5)
            if j == 3:
                finlist.append(6)
            if j ==4:
                finlist.append(11)
            if j ==5:
                finlist.append(13)
            if j== 6:
                finlist.append(16)
            if j == 7:
                finlist.append(19)
        overlist.append(finlist)
    return overlist
    
book = xlrd.open_workbook('divpaydatedata.xls')

paydates = book.sheet_by_index(0)
exdates = book.sheet_by_index(1)
payments = book.sheet_by_index(2)

exdatemaster  = []
paydatemaster = [] 
paymentmaster = []

for i in range(1, 22):
    '''
    retrieve all pay date/ex date and payment date info from excel file
    '''
    
    exdatevals = exdates.col_values(i-1, start_rowx=1, end_rowx=96)
    paydatevals = paydates.col_values(i-1, start_rowx=1, end_rowx=96)
    paymentvals = payments.col_values(i-1, start_rowx=1, end_rowx=96)
    
    for p in exdatevals:
        if p == '':
            exdatevals.pop()
            
    for q in paydatevals:
        if q == '':
            paydatevals.pop()
        
    for r in paymentvals: 
        if r == '':
            paymentvals.pop()
        else:
            r = float(r)
    if i-1 ==1:
        paymentvals.reverse()
    exdatemaster.append(exdatevals)
    paydatemaster.append(paydatevals)
    paymentmaster.append(paymentvals)

def driptest(numlist):
    '''
    input: a list of numbers (nums corrsp to stocks listed in tckrlist)
    output: the value of a portfolio containing stocks specified  in numlist starting on 9/19/2000 and ending on 7/15/2013
            divs reinvested at price on close of pay dates
    '''
    #TODO add semi-annual contributions, invested evenly between all stocks, every 130 days
    cashperstock = 50000/len(numlist) #start with $50,000 dollars divided by number of stocks
    holdingslist = []

    for i in numlist: #buy initial shares
        price = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[i], (2000, 9, 19), (2000, 9, 19), dividends=False), adjusted=False)
        iprice = price[0][2]

        ishares = round(cashperstock/iprice, 2)
        holdingslist.append(ishares)

    for i in fulldatelist: #for every day in window
        hcount = 0
        p = i.split('/')
        d = int(p[1])
        yahoodate = (int(p[2]), int(p[0]), d) # a yahoo fetch compatible date
        hcount = 0
        for j in numlist: #look at every stock in the numlist
            paycount = 0
            icurrprice = 0
            d = int(p[1])
            yahoodate = (int(p[2]), int(p[0]), d)
            for k in paydatemaster[j]: #for each stock, look at payment date master list
                k = k.strip()
                if k == i: #check if current day is a dividend pay day
                    paymentmaster[j][paycount] = float(paymentmaster[j][paycount])
                    divpaymt = round(paymentmaster[j][paycount]*holdingslist[hcount], 2)

                    try:
                        currprice = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[j], yahoodate, yahoodate, dividends=False), adjusted=False)
                        icurrprice = currprice[0][2]
                        holdingslist[hcount] += round(divpaymt/icurrprice, 3) #reinvest, using yahoo data, tckrlist[j]
                    except: #sometimes paydates are on the weekend, in which case the next available day's close price is used     
                        while icurrprice == 0:
                            d+= 1
                            try:
                                yahoodate = (int(p[2]), int(p[0]), d)
                                currprice = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[j], yahoodate, yahoodate, dividends=False), adjusted=False)
                                icurrprice = currprice[0][2]
                                holdingslist[hcount] += round(divpaymt/icurrprice, 3)
                            except:
                                pass
                paycount += 1
            hcount += 1

    finaldate = (2013, 7, 15)
    count = 0
    value = 0
    for i in numlist:
        finalprice = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[i], finaldate, finaldate, dividends=False), adjusted=False)
        ifinalprice = finalprice[0][2]
        value += round(ifinalprice*holdingslist[count], 2) #calculate final value
        count += 1
    return value
    
    
def manualtest(numlist):
    '''
    input: a list of numbers (nums corrsp to stocks listed in tckrlist)
    output: the value of a portfolio containing stocks specified in numlist starting on 9/19/2000 and ending on 7/15/2013
            using a manual reinvestment strategy
    '''
    daycount = 0
    cash = 0
    transactions = 0
    cashperstock = 50000/len(numlist) #start with $50,000 dollars divided by number of stocks
    holdingslist = []

    for i in numlist: #buy initial shares

        price = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[i], (2000, 9, 19), (2000, 9, 19), dividends=False), adjusted=False)
        iprice = price[0][2]

        ishares = round(cashperstock/iprice, 2)
        holdingslist.append(ishares)

    for i in fulldatelist: #for every day in window
        fiftydaylist = []
        buyconditiontf =[]
        tendaylist= []
        mrdaystckprice = []

        hcount = 0
        for j in numlist: #look at every stock in the randomly chosen numlist
            paycount = 0
            for k in paydatemaster[j]: #for each stock, look at payment date master list
                k = k.strip()
                if k == i: #check if current day is a div pay day
                    paymentmaster[j][paycount] = float(paymentmaster[j][paycount])
                    divpaymt = round(paymentmaster[j][paycount]*holdingslist[hcount], 2)
                    cash += divpaymt # if it's a pay day, add the div payment to cash
                paycount += 1
            hcount += 1  

        if daycount > 50: # if more than 50 days have passed
            for u in numlist:
                tempolist = []
                dcount = 0
                for p in stockdatemaster[u]:
                    if p == i:
                        tempolist = stockpricemaster[u][dcount:dcount+50] #stock prices for last 50 days
                    dcount += 1
                if len(tempolist) != 0:
                    fstdv, favg = meanstdv(tempolist) #calculate fifty day moving average
                    tstdv, tavg = meanstdv(tempolist[:10]) #calculate 10 day moving average
                    fiftydaylist.append(favg)
                    tendaylist.append(tavg)
                    mrdaystckprice.append(tempolist[0])

        fcount = 0
        truelist = [] 

        for s in mrdaystckprice:
            
            if float(s) < fiftydaylist[fcount] and float(s) < tendaylist[fcount] and tendaylist[fcount] < fiftydaylist[fcount]: #buy criteria defined
                buyconditiontf.append(True)
                truelist.append(fcount)
#                print i
                
            else:
                buyconditiontf.append(False)
            
            fcount += 1
        fcount = 0

        if cash != 0 and len(truelist) != 0: #if there is cash available and one or more buy criteria have been met
            if cash/len(truelist) > 500: #if there's more than $500 cash
                for e in buyconditiontf: 
                    if e == True:

                        holdingslist[fcount] += (cash/len(truelist))/float(mrdaystckprice[fcount]) #for all stocks meeting buy criteria, add to holdings using available cash
                        transactions += 1 #add one transaction to the list
                    fcount += 1
                cash = 0 #reset cash to zero; all cash was spent adding to stocks meeting buy criteria

        daycount += 1
    finaldate = (2013, 7, 15)
    count = 0
    value = 0
    for i in numlist:
        finalprice = mf.parse_yahoo_historical(mf.fetch_historical_yahoo(tckrlist[i], finaldate, finaldate, dividends=False), adjusted=False)
        ifinalprice = finalprice[0][2]

        value += round((ifinalprice*holdingslist[count])+cash, 2) #calculate final value for each stock
        count += 1
        
    return value, transactions


#m,t = manualtest([4])
#d =  driptest([4])
#print 'The manual value was', m, 'and', t, 'transactions took place'


##### Note: the code below was used to generate the data for the seeking alpha article. 
# It takes about 40 minutes to run, and generates a plot and output file with a summary of results.
####

#
#sum1 =0
#for z in range(1, 9):
#    print 'there are ', figure(z, 8), 'combinations of ',z
#    sum1 += (figure(z, 8))
#
#f = open('dripman.txt', 'w')
#f.write('DRIP, MAN, Transactions, stocks')
#f.write('\n')
#completedrip = []
#completeman = []
#for z in range(1, 9):
#    f.write(str(z)+':')
#    f.write('\n')
#    driphigh = []
#    dripperc = []
#    manhigh = []
#    manperc = []
#    a = generate_numlists(z)
#    print a
#    for r in a:
#       
#        d= driptest(r)
#        m, t = manualtest(r)
#        f.write(str(d)+' , '+str(m)+' , '+str(t)+' , '+str(r))
#        f.write('\n')
#        print d, ',', m, t, '   ',r
#        driphigh.append(d)
#        completedrip.append(d)
#        completeman.append(m)
#        manhigh.append(m)
#        if d > m:
#            diff = d-m
#            more = (diff/m)*100
#            
#            dripperc.append(more)
#        if m >d:
#            diff = m -d
#            more = (diff/d)*100
#            
#            manperc.append(more)
#    print
#    print '***', 'drip:', meanstdv(driphigh), 'man: ',meanstdv(manhigh), 'drip%: ',len(dripperc)/len(a), 'man%: ', len(manperc)/len(a)
#    print
#    f.write('\n')
#    f.write(str(meanstdv(driphigh))+' , '+str(meanstdv(manhigh)))
#    f.write('\n')
#    f.write(str(meanstdv(dripperc))+' , '+str(meanstdv(manperc)))
#    f.write('\n')
#    f.write('DRIP% higher:'+str((len(dripperc)/len(a))*100))
#    f.write('\n')
#    f.write('MAN% higher:'+str((len(manperc)/len(a))*100))
#    f.write('\n')
#drip = scipy.zeros((len(completedrip), 1))
#man = scipy.zeros((len(completeman), 1))
#count = 0
#for i in completedrip:
#    drip[count] = i
#    count += 1
#count = 0
#for i in completeman:
#    man[count] = i
#    count += 1
#
#
#fig, ax = pylab.subplots()
#ax.set_xlabel('DRIP Results ($)', fontsize=15)
#ax.set_ylabel('Manual Results ($)', fontsize=15)
#ax.set_title('DRIP versus Manual Purchase Results: Size = '+str(z), fontsize=15)
#l = pylab.Line2D([0, 450000], [0, 450000])
#ax.plot(drip, man, 'o')
#ax.add_line(l)
#ax.axis([0.0,450000.0, 0.0,450000])
#
#ax.set_autoscale_on(False)
#ax.grid(True)
#pylab.plt.savefig('drip-man-comp')
#f.close()
#
